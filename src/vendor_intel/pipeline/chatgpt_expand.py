"""Full market expand via OpenAI ChatGPT SDK only (no run_query.py).

Steps:
  1) ChatGPT recalls known channel companies for the market
  2) OpenAI SDK web_search discovers companies (+ optional ddgs harvest)
  3) ChatGPT extracts + scores market relevance from ddgs hits (if used)
  4) ChatGPT verifies keep/reject (related to market or not)
  5) Fill columns: Google AI scraper → DDGS/SearXNG/Wikipedia/Owler first,
     then OpenAI SDK web_search residual for leftover Not publicly disclosed gaps
  6) Verify/correct per-field evidence (field_verify_fill.py): Core Categories,
     Specialty Focus, Key Brands Represented + channel-fit keep/remove,
     Online/Offline/Both classification, Contact Person/Role/LinkedIn via
     Apollo free search + web-search fallback, and multi-location geography
  7) ChatGPT final verify pass on filled rows + write wrap-formatted FINAL
     Excel + audit JSON
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

from openai import OpenAI

from vendor_intel.config import Settings
from vendor_intel.pipeline import openai_cost
from vendor_intel.pipeline.openai_expand_checkpoint import ExpandCheckpoint
from vendor_intel.pipeline.run_and_export import auto_seed_names, load_seeds_file
from vendor_intel.pipeline.web_expand import (
    HEADERS,
    _blank_row,
    _norm,
    default_output_dir,
    market_family,
    merge_candidates,
    read_final_rows,
    resolve_final_path,
    web_harvest_candidates,
    write_final_xlsx,
)


def _seed_candidates(
    query: str,
    country: str,
    *,
    seeds_path: str | None = None,
) -> list[dict[str, str]]:
    """Load curated seed companies (e.g. queries/seeds/global_avocado_oil_market.txt)."""
    rows = load_seeds_file(seeds_path) if seeds_path else auto_seed_names(query, country)
    out: list[dict[str, str]] = []
    seen: set[str] = set()
    for name, domain, section in rows:
        key = _norm(name)
        if not key or key in seen:
            continue
        seen.add(key)
        website = f"https://{domain}" if domain else ""
        out.append(
            {
                "name": name,
                "website": website,
                "domain": domain or "",
                "snippet": section or "curated seed",
                "source_query": "seed_file",
                "discovery_source": "seed_file",
            }
        )
    return out

_JSON_FENCE = re.compile(r"```(?:json)?\s*([\s\S]*?)```", re.I)

ROLE_HINTS = {
    "avocado": "distributors, wholesalers, importers, specialty oil suppliers, grocery/retail chains, foodservice suppliers of avocado oil / edible oils",
    "firewall": "distributors, VADs, resellers, MSSPs, system integrators, channel partners for firewalls / network security (NOT the OEM vendors themselves)",
    "smartwatch": "distributors, wholesalers, electronics retailers, channel partners for smartwatches / wearables (NOT Apple/Samsung/Garmin as OEMs)",
    "green": "specialty chemical distributors / ingredients distributors for green / bio-based chemicals (NOT pure petrochem manufacturers)",
    "renewable": (
        "distributors, wholesalers, dealers, importers, EPCs/system integrators, "
        "authorized channel partners for solar/wind/storage "
        "(NOT module/inverter/turbine OEMs; NOT utilities/IPPs/developers as primary)"
    ),
    "general": (
        "distributors, wholesalers, channel partners, VADs, resellers, dealers, "
        "system integrators relevant to the market "
        "(NOT pure OEMs/manufacturers; NOT media/associations)"
    ),
}


def _base_url(settings: Settings | None = None) -> str:
    if settings is not None:
        raw = (settings.openai_base_url or "").strip()
    else:
        raw = (os.getenv("OPENAI_BASE_URL") or "").strip()
    return (raw or "https://api.openai.com/v1").rstrip("/")


def _is_deepseek(settings: Settings | None = None) -> bool:
    """True when OPENAI_* points at DeepSeek (OpenAI-compatible, no hosted web_search)."""
    url = _base_url(settings).lower()
    model = (
        (settings.openai_model if settings else None)
        or os.getenv("OPENAI_MODEL")
        or ""
    ).lower()
    return "deepseek" in url or model.startswith("deepseek")


def _supports_hosted_web_search(settings: Settings | None = None) -> bool:
    """OpenAI Responses ``web_search`` is OpenAI-hosted only — not on DeepSeek."""
    return not _is_deepseek(settings)


def _llm_label(settings: Settings | None = None) -> str:
    return "DeepSeek" if _is_deepseek(settings) else "OpenAI"


def _client(settings: Settings) -> OpenAI:
    key = (settings.openai_api_key or "").strip()
    if not key:
        raise RuntimeError("OPENAI_API_KEY not set in .env")
    return OpenAI(
        api_key=key,
        base_url=_base_url(settings),
    )


def _model(settings: Settings) -> str:
    default = "deepseek-chat" if _is_deepseek(settings) else "gpt-4o-mini"
    return (settings.openai_model or default).strip() or default


def _log(msg: str) -> None:
    print(msg, flush=True)


def _repair_json(raw: str) -> str:
    """Best-effort repair when the model truncates a JSON string mid-response."""
    s = raw.strip()
    if not s:
        return "{}"
    # close open strings / brackets roughly
    if s.count('"') % 2 == 1:
        s += '"'
    # trim broken trailing field
    s = re.sub(r",\s*\"[^\"]*$", "", s)
    s = re.sub(r",\s*$", "", s)
    open_b = s.count("{") - s.count("}")
    open_a = s.count("[") - s.count("]")
    if open_a > 0:
        s += "]" * open_a
    if open_b > 0:
        s += "}" * open_b
    return s


def _parse_json(text: str) -> Any:
    raw = (text or "").strip()
    m = _JSON_FENCE.search(raw)
    if m:
        raw = m.group(1).strip()
    start = raw.find("{")
    start_arr = raw.find("[")
    if start_arr != -1 and (start == -1 or start_arr < start):
        raw = raw[start_arr:]
    elif start != -1:
        raw = raw[start:]
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        repaired = _repair_json(raw)
        return json.loads(repaired)


def _chat_json(
    client: OpenAI,
    model: str,
    system: str,
    user: str,
    *,
    label: str = "chat",
    max_tokens: int = 4096,
    retries: int = 3,
) -> Any:
    last_err: Exception | None = None
    provider = "DeepSeek" if "deepseek" in (model or "").lower() or _is_deepseek() else "OpenAI"
    for attempt in range(1, retries + 1):
        _log(f"      · {label}: calling {provider} ({model}) attempt {attempt}/{retries}…")
        try:
            kwargs: dict[str, Any] = {
                "model": model,
                "temperature": 0.15,
                "max_tokens": max_tokens,
                "messages": [
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
            }
            # DeepSeek + OpenAI both accept json_object; skip only if model forbids it
            try:
                resp = client.chat.completions.create(
                    response_format={"type": "json_object"},
                    **kwargs,
                )
            except Exception as fmt_err:  # noqa: BLE001
                if "response_format" in str(fmt_err).lower() or "json" in str(fmt_err).lower():
                    _log(f"      · {label}: json_object unsupported — plain chat retry")
                    resp = client.chat.completions.create(**kwargs)
                else:
                    raise
            rec = openai_cost.record_response_usage(
                resp, label=label, model=model, step=label.split("-")[0]
            )
            if rec:
                sess = openai_cost.get_session()
                run_tot = sess.estimated_cost_usd if sess else rec.estimated_cost_usd
                _log(
                    f"      · {label}: tokens in={rec.prompt_tokens} "
                    f"out={rec.completion_tokens} "
                    f"cost~${rec.estimated_cost_usd:.4f} "
                    f"(run total ~${run_tot:.4f})"
                )
            content = resp.choices[0].message.content or "{}"
            finish = getattr(resp.choices[0], "finish_reason", None)
            _log(
                f"      · {label}: got response "
                f"({len(content)} chars, finish={finish})"
            )
            if finish == "length":
                _log(f"      · {label}: response truncated — will repair/retry smaller")
            try:
                data = _parse_json(content)
                _log(f"      · {label}: JSON parse OK")
                return data
            except json.JSONDecodeError as err:
                last_err = err
                _log(f"      · {label}: JSON parse failed: {err}")
                # shrink ask on retry by appending instruction
                user = (
                    user
                    + "\n\nIMPORTANT: Return SHORTER valid JSON only. "
                    "Fewer items if needed. No trailing commas. Complete all brackets."
                )
                max_tokens = min(max_tokens + 1024, 8192)
        except Exception as err:  # noqa: BLE001
            last_err = err
            _log(f"      · {label}: API error: {type(err).__name__}: {err}")
    raise RuntimeError(f"{label} failed after {retries} attempts: {last_err}")


def _response_output_text(resp: Any) -> str:
    text = getattr(resp, "output_text", None)
    if isinstance(text, str) and text.strip():
        return text
    parts: list[str] = []
    for item in getattr(resp, "output", None) or []:
        for content in getattr(item, "content", None) or []:
            t = getattr(content, "text", None)
            if isinstance(t, str) and t.strip():
                parts.append(t)
            elif isinstance(content, dict) and content.get("text"):
                parts.append(str(content["text"]))
    return "\n".join(parts).strip()


def _web_search_json(
    client: OpenAI,
    model: str,
    system: str,
    user: str,
    *,
    label: str = "web-fill",
    max_tokens: int = 4096,
    retries: int = 3,
) -> Any:
    """OpenAI Responses API + hosted web_search → JSON; DeepSeek uses chat.completions.

    DeepSeek has no hosted ``web_search`` tool. Skip failed Responses retries and go
    straight to chat (pair with search-stack fill / ddgs harvest for live web facts).
    """
    if not _supports_hosted_web_search():
        _log(
            f"      · {label}: DeepSeek mode — chat.completions "
            f"(no OpenAI hosted web_search) ({model})"
        )
        return _chat_json(
            client,
            model,
            system=system
            + " Use best public knowledge and any SERP/evidence in the prompt; "
            "mark unknowns Not publicly disclosed. Never invent emails/phones.",
            user=user,
            label=f"{label}-deepseek-chat",
            max_tokens=max_tokens,
            retries=retries,
        )

    last_err: Exception | None = None
    for attempt in range(1, retries + 1):
        _log(
            f"      · {label}: OpenAI web_search ({model}) attempt {attempt}/{retries}…"
        )
        try:
            try:
                resp = client.responses.create(
                    model=model,
                    instructions=system,
                    input=user,
                    tools=[{"type": "web_search"}],
                    tool_choice={"type": "web_search"},
                    temperature=0.15,
                    max_output_tokens=max_tokens,
                    text={"format": {"type": "json_object"}},
                )
            except Exception:
                # Older SDK / model combo: no json_object + tool_choice variants
                resp = client.responses.create(
                    model=model,
                    instructions=system
                    + " Return ONLY a single JSON object. No markdown.",
                    input=user,
                    tools=[{"type": "web_search"}],
                    temperature=0.15,
                    max_output_tokens=max_tokens,
                )
            rec = openai_cost.record_response_usage(
                resp, label=label, model=model, step=label.split("-")[0]
            )
            if rec:
                sess = openai_cost.get_session()
                run_tot = sess.estimated_cost_usd if sess else rec.estimated_cost_usd
                _log(
                    f"      · {label}: tokens in={rec.prompt_tokens} "
                    f"out={rec.completion_tokens} "
                    f"cost~${rec.estimated_cost_usd:.4f} "
                    f"(run total ~${run_tot:.4f})"
                )
            content = _response_output_text(resp) or "{}"
            _log(f"      · {label}: got response ({len(content)} chars)")
            try:
                data = _parse_json(content)
                _log(f"      · {label}: JSON parse OK")
                return data
            except json.JSONDecodeError as err:
                last_err = err
                _log(f"      · {label}: JSON parse fail — retry")
                user = (
                    user
                    + "\n\nIMPORTANT: Return SHORTER valid JSON only with all required keys."
                )
        except Exception as err:  # noqa: BLE001
            last_err = err
            _log(f"      · {label}: API error: {type(err).__name__}: {err}")
            # Fall back to chat completions (no live web) on hard failure
            if attempt == retries:
                _log(f"      · {label}: falling back to chat.completions (no web_search)")
                return _chat_json(
                    client,
                    model,
                    system=system
                    + " Use best public knowledge; mark unknowns Not publicly disclosed.",
                    user=user,
                    label=f"{label}-fallback",
                    max_tokens=max_tokens,
                )
    raise RuntimeError(f"{label} failed after {retries} attempts: {last_err}")


_EMPTYISH = frozenset(
    {
        "",
        "n/a",
        "na",
        "none",
        "null",
        "unknown",
        "-",
        "—",
        "tbd",
        "todo",
    }
)

# Must be web-evidenced or "Not publicly disclosed" — never invent / never generic defaults
_FACT_COLUMNS = frozenset(
    {
        "Founded",
        "Headquarters",
        "Continent / Geography",
        "Operational Presence",
        "Ownership",
        "Employees",
        "Website",
        "Key Brands Represented",
        "Contact Person",
        "Role",
        "Email",
        "LinkedIn",
        "Office No.",
        "Country Code",
        "Region Code",
    }
)

def _looks_generic_default(header: str, value: str, family: str) -> bool:
    """No more template defaults to check against — just the empty/NPD gate.
    header/family kept for call-site compatibility."""
    _ = header, family
    v = (value or "").strip().lower()
    return not v or v in _EMPTYISH


def _ensure_complete_row(
    row: dict[str, str],
    *,
    family: str,
    name: str,
    website: str,
    source: str,
    real_facts_only: bool = True,
) -> dict[str, str]:
    """Keep only evidenced values — never invent MARKET_DEFAULTS / known_* text."""
    from vendor_intel.export.landscape_slim import finalize_slim_row

    base = _blank_row(name, website, family, source)
    out = dict(base)
    for h in HEADERS:
        val = str(row.get(h) or "").strip()
        if val.lower() in _EMPTYISH or val.lower() == "not publicly disclosed":
            # Drop NPD placeholders that were never real scrape facts
            if real_facts_only and h not in {
                "Company",
                "Website",
            }:
                val = ""
        if real_facts_only and h in _FACT_COLUMNS and _looks_generic_default(h, val, family):
            val = ""
        if val:
            out[h] = val
    # Legacy column aliases from older fills
    if not out.get("Office No.") and row.get("Phone / WhatsApp"):
        ph = str(row.get("Phone / WhatsApp") or "").strip()
        if ph and ph.lower() not in _EMPTYISH and ph.lower() != "not publicly disclosed":
            out["Office No."] = ph
    if not out.get("Continent / Geography") and row.get("Cities / Regions"):
        out["Continent / Geography"] = str(row.get("Cities / Regions") or "")
    if not out.get("Operational Presence") and row.get("Regions Served"):
        out["Operational Presence"] = str(row.get("Regions Served") or "")

    out["Company"] = name or out.get("Company") or ""
    if website and (
        not out.get("Website")
        or str(out.get("Website")).lower() in _EMPTYISH
        or out.get("Website") == "Not publicly disclosed"
    ):
        out["Website"] = website

    # Contacts stay empty unless evidenced
    for c in ("Contact Person", "Role", "Email", "LinkedIn", "Office No."):
        v = str(out.get(c) or "").strip()
        if not v or v.lower() in _EMPTYISH or v.lower() == "not publicly disclosed":
            out[c] = ""

    # Leave all other missing columns blank (real data only — no NPD paint)
    for h in HEADERS:
        if not str(out.get(h) or "").strip():
            out[h] = ""

    out = finalize_slim_row(out)
    # Soft confidence signals kept for internal merges (not in slim Excel)
    founded = str(out.get("Founded") or "")
    hq = str(out.get("Headquarters") or "")
    web = str(out.get("Website") or "")
    real_bits = 0
    if founded and founded != "Not publicly disclosed" and any(ch.isdigit() for ch in founded):
        real_bits += 1
    if hq and hq != "Not publicly disclosed" and len(hq) > 3:
        real_bits += 1
    if web.startswith("http") and "not publicly" not in web.lower():
        real_bits += 1
    brands = str(out.get("Key Brands Represented") or "")
    if brands and brands != "Not publicly disclosed" and not _looks_generic_default(
        "Key Brands Represented", brands, family
    ):
        real_bits += 1
    if real_bits >= 3:
        out["Data Confidence"] = "High"
        out["Quality Score"] = "4"
    elif real_bits >= 1:
        out["Data Confidence"] = "Medium"
        out["Quality Score"] = "3"
    else:
        out["Data Confidence"] = "Low"
        out["Quality Score"] = "2"
    src = out.get("Data Sources") or source
    if source and source not in src:
        out["Data Sources"] = f"{src}; {source}".strip("; ")
    return out


async def _local_web_evidence(name: str, website: str, market: str) -> str:
    """Extra SERP snippets (ddgs) to ground OpenAI fill with real public text."""
    try:
        from vendor_intel.clients.search_router import FreeSearchRouter

        router = FreeSearchRouter(Settings())
        # One fact per Google/DDGS query (same style as Google AI scraper gap-fill)
        qn = f'"{name}"' if " " in name else name
        queries = [
            f"founded year of {qn}",
            f"headquarters of {qn}",
            f"number of employees of {qn}",
            f"official website of {qn}",
            f"operational presence of {qn}",
            f"key brands represented of {qn}",
        ]
        if website:
            queries.append(
                f"site:{website.replace('https://', '').replace('http://', '').split('/')[0]} about"
            )
        bits: list[str] = []
        for q in queries[:4]:
            try:
                hits = await router.search(q, search_topic="general", validation_mode=True)
            except Exception:
                continue
            for h in (hits or [])[:4]:
                title = getattr(h, "title", "") or ""
                link = getattr(h, "link", "") or ""
                snip = getattr(h, "snippet", "") or ""
                bits.append(f"- {title} | {link} | {snip}")
        return "\n".join(bits[:12])
    except Exception as err:  # noqa: BLE001
        return f"(local evidence unavailable: {type(err).__name__})"


def _row_needs_web_refill(row: dict[str, str]) -> bool:
    """True if key firmographic / channel fields are still weak."""
    weak_keys = (
        "Founded",
        "Headquarters",
        "Business Type",
        "Core Categories",
        "Specialty Focus",
        "Website",
        "Regions Served",
        "Summary",
    )
    weak = 0
    for k in weak_keys:
        v = str(row.get(k) or "").strip().lower()
        if not v or v in _EMPTYISH or v == "not publicly disclosed":
            weak += 1
    return weak >= 3


def gpt_recall_companies(
    client: OpenAI,
    model: str,
    *,
    query: str,
    family: str,
    target: int,
    ckpt: ExpandCheckpoint | None = None,
) -> list[dict[str, str]]:
    """ChatGPT lists known real channel companies in small pages (avoids JSON truncate)."""
    roles = ROLE_HINTS.get(family, ROLE_HINTS["general"])
    want_total = min(max(target, 40), 1200)
    page_size = 35
    out: list[dict[str, str]] = list(ckpt.data("recalled") or []) if ckpt else []
    seen: set[str] = {_norm(str(r.get("name") or "")) for r in out if r.get("name")}
    start_page = int(ckpt.progress_get("recall_page", 0) or 0) if ckpt else 0
    pages = (want_total + page_size - 1) // page_size
    _log(
        f"    → substep 1a: recall up to {want_total} companies "
        f"in {pages} page(s) of {page_size}"
        + (f" (resume from page {start_page + 1}, have {len(out)})" if start_page else "")
    )

    for page in range(start_page, pages):
        exclude = list(seen)[-80:]  # avoid huge prompts
        _log(f"    → substep 1a.{page + 1}: recall page {page + 1}/{pages}…")
        data = _chat_json(
            client,
            model,
            system=(
                "You list REAL companies that operate as channel partners for a market "
                f"({roles}). No invented names. No pure OEMs/brands unless they also run "
                "a distribution channel. Return compact JSON only: "
                '{"companies":[{"name":"...","website":"https://...","role":"...",'
                '"country":"...","why_related":"..."}]}'
            ),
            user=(
                f"Market: {query}\n"
                f"List up to {page_size} REAL distributors/wholesalers/retailers/"
                "channel partners for this market (diverse countries).\n"
                f"Do NOT repeat these names: {json.dumps(exclude, ensure_ascii=False)}"
            ),
            label=f"recall-page-{page + 1}",
            max_tokens=3500,
        )
        rows = data.get("companies") if isinstance(data, dict) else []
        added = 0
        if isinstance(rows, list):
            for row in rows:
                if not isinstance(row, dict):
                    continue
                name = str(row.get("name") or "").strip()
                if not name or _norm(name) in seen:
                    continue
                seen.add(_norm(name))
                website = str(row.get("website") or "").strip()
                out.append(
                    {
                        "name": name,
                        "website": website,
                        "domain": website.replace("https://", "")
                        .replace("http://", "")
                        .split("/")[0],
                        "snippet": str(row.get("why_related") or row.get("role") or ""),
                        "source_query": str(row.get("country") or "chatgpt_recall"),
                        "discovery_source": "chatgpt_recall",
                    }
                )
                added += 1
        _log(f"    → substep 1a.{page + 1}: +{added} (total recalled={len(out)})")
        if ckpt:
            ckpt.bump(
                "1_recall",
                f"1a.{page + 1}",
                progress={"recall_page": page + 1},
                recalled=out,
                note=f"recall page {page + 1}",
            )
        if added == 0:
            _log("    → substep 1a: no new names — stop recall early")
            break
    _log(f"    → substep 1a done: {len(out)} companies from ChatGPT knowledge")
    if ckpt:
        ckpt.mark_step_done("1_recall", recalled=out)
    return out


def gpt_discover_companies(
    client: OpenAI,
    model: str,
    *,
    query: str,
    family: str,
    target: int,
    exclude: list[str] | None = None,
    max_pages: int = 16,
    page_size: int = 25,
    ckpt: ExpandCheckpoint | None = None,
) -> list[dict[str, str]]:
    """Find channel companies via OpenAI Responses API ``web_search`` (not ddgs).

    Keeps paging until ``target`` is reached or ``max_pages`` is exhausted.
    An empty page does **not** stop discovery (avoids stalling under target).
    """
    roles = ROLE_HINTS.get(family, ROLE_HINTS["general"])
    out: list[dict[str, str]] = list(ckpt.data("discovered") or []) if ckpt else []
    seen: set[str] = {_norm(x) for x in (exclude or []) if x}
    for r in out:
        seen.add(_norm(str(r.get("name") or "")))
    want = max(target, 40)
    # Use full page budget — do not shrink pages so hard that we stop short of target
    pages = max(1, max_pages)
    start_page = int(ckpt.progress_get("discover_page", 0) or 0) if ckpt else 0
    geos = [
        "United States",
        "Europe",
        "India",
        "Latin America",
        "Middle East",
        "Asia Pacific",
        "Africa",
        "Canada",
        "United Kingdom",
        "Germany",
        "Brazil",
        "Southeast Asia",
        "Australia / New Zealand",
        "Japan / Korea",
        "global / multi-region",
    ]
    angles = [
        "trade association member directories and accredited distributor lists",
        "\"top distributors\" / \"largest wholesalers\" ranking articles for this market",
        "B2B marketplace seller lists and authorized reseller / partner directories",
        "import/export trading companies and regional importers for this market",
        "national and regional retail chains / specialist retailers in this market",
        "foodservice / HORECA / IT channel / wholesale catalogs (as fits the market)",
        "LinkedIn company search results for distributors and wholesalers",
        "industry expo exhibitor lists and sponsor directories",
    ]
    empty_streak = 0
    _log(
        f"    → substep 2a: "
        f"{'DeepSeek chat' if _is_deepseek() else 'OpenAI web_search'} discover "
        f"up to {want} companies "
        f"in ≤{pages} page(s) (continue through empty pages until target)"
        + (f" (resume page {start_page + 1}, have {len(out)})" if start_page else "")
    )
    for page in range(start_page, pages):
        if len(out) >= want:
            _log(
                f"    → substep 2a: hit target ({len(out)}/{want}) — stop discover"
            )
            break
        geo = geos[page % len(geos)]
        angle = angles[page % len(angles)]
        exclude_tail = [x for x in seen][-120:]
        need = min(page_size, want - len(out))
        _log(
            f"    → substep 2a.{page + 1}: web discover page {page + 1}/{pages} "
            f"({geo}; {angle[:48]}…, need +{need})…"
        )
        data = _web_search_json(
            client,
            model,
            system=(
                "You find REAL companies using live web search. "
                f"Focus on channel roles: {roles}. "
                "Reject pure OEMs/brands that only manufacture, media, research firms, "
                "directories, and market-report publishers. "
                "Prefer distributors, wholesalers, importers, retailers, foodservice suppliers. "
                "Return compact JSON only: "
                '{"companies":[{"name":"...","website":"https://...","role":"...",'
                '"country":"...","why_related":"..."}]}'
            ),
            user=(
                f"Market: {query}\n"
                f"Geography focus this page: {geo}\n"
                f"Search angle this page: {angle}\n"
                f"Web-search and list up to {need} DISTINCT real companies that distribute / "
                "wholesale / import / retail channel this market.\n"
                "Prioritize named companies with websites. Avoid repeating prior names.\n"
                "Search listicles, association directories, trade sites, LinkedIn, company sites.\n"
                f"Do NOT repeat these names: {json.dumps(exclude_tail, ensure_ascii=False)}"
            ),
            label=f"web-discover-{page + 1}",
            max_tokens=3500,
        )
        rows = data.get("companies") if isinstance(data, dict) else None
        if not isinstance(rows, list) and isinstance(data, dict) and data.get("name"):
            rows = [data]
        added = 0
        if isinstance(rows, list):
            for row in rows:
                if not isinstance(row, dict):
                    continue
                name = str(row.get("name") or "").strip()
                if not name or _norm(name) in seen:
                    continue
                seen.add(_norm(name))
                website = str(row.get("website") or "").strip()
                domain = (
                    website.replace("https://", "")
                    .replace("http://", "")
                    .split("/")[0]
                    .lower()
                )
                out.append(
                    {
                        "name": name,
                        "website": website,
                        "domain": domain,
                        "snippet": str(row.get("why_related") or row.get("role") or ""),
                        "source_query": str(row.get("country") or geo),
                        "discovery_source": "openai_web_search_discover",
                        "relevance": "80",
                    }
                )
                added += 1
        _log(
            f"    → substep 2a.{page + 1}: +{added} "
            f"(openai discovered total={len(out)}/{want})"
        )
        if ckpt:
            ckpt.bump(
                "2_discover",
                f"2a.{page + 1}",
                progress={"discover_page": page + 1},
                discovered=out,
                note=f"discover page {page + 1}",
            )
        if added == 0:
            empty_streak += 1
            _log(
                f"    → substep 2a: +0 new names on page {page + 1} "
                f"(empty streak={empty_streak}) — continuing until target or page budget"
            )
        else:
            empty_streak = 0
    _log(f"    → substep 2a done: {len(out)} companies from "
         f"{'DeepSeek chat' if _is_deepseek() else 'OpenAI web_search'}")
    if ckpt:
        ckpt.mark_step_done("2_discover", discovered=out)
    return out


def gpt_extract_and_score(
    client: OpenAI,
    model: str,
    *,
    query: str,
    family: str,
    candidates: list[dict[str, str]],
) -> list[dict[str, str]]:
    """Extract companies from search hits and score market relevance 0-100."""
    if not candidates:
        return []

    roles = ROLE_HINTS.get(family, ROLE_HINTS["general"])
    kept: list[dict[str, str]] = []
    chunk_size = 15
    total_chunks = (len(candidates) + chunk_size - 1) // chunk_size
    _log(
        f"    → substep 3a: score {len(candidates)} web hits "
        f"in {total_chunks} ChatGPT chunk(s)"
    )
    for i in range(0, len(candidates), chunk_size):
        chunk = candidates[i : i + chunk_size]
        idx = i // chunk_size + 1
        _log(f"    → substep 3a.{idx}: extract+score chunk {idx}/{total_chunks} ({len(chunk)} hits)…")
        payload = [
            {
                "name": c.get("name"),
                "website": c.get("website"),
                "domain": c.get("domain"),
                "snippet": (c.get("snippet") or "")[:220],
                "source_query": c.get("source_query"),
            }
            for c in chunk
        ]
        data = _chat_json(
            client,
            model,
            system=(
                f"You extract channel companies for a market. Target roles: {roles}. "
                "Reject OEMs-only, media, research firms, directories, job boards, "
                "unrelated product pages, and companies not related to this market. "
                "Return compact JSON: {\"companies\":[{\"name\":\"...\",\"website\":\"https://...\","
                "\"role\":\"...\",\"country_hint\":\"...\",\"relevance\":0-100,"
                "\"why_related\":\"...\",\"keep\":true|false}]}"
            ),
            user=(
                f"Market: {query}\nFamily: {family}\n"
                f"Web search candidates:\n{json.dumps(payload, ensure_ascii=False)}\n"
                "Only keep companies clearly related to this market (relevance >= 60)."
            ),
            label=f"extract-{idx}",
            max_tokens=3500,
        )
        rows = data.get("companies") if isinstance(data, dict) else []
        if not isinstance(rows, list):
            continue
        n_keep = 0
        for row in rows:
            if not isinstance(row, dict):
                continue
            if row.get("keep") is False:
                continue
            try:
                rel = int(row.get("relevance") or 0)
            except (TypeError, ValueError):
                rel = 0
            if rel < 60:
                continue
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            website = str(row.get("website") or "").strip()
            kept.append(
                {
                    "name": name,
                    "website": website,
                    "domain": website.replace("https://", "").replace("http://", "").split("/")[0],
                    "snippet": str(row.get("why_related") or row.get("role") or ""),
                    "source_query": str(row.get("country_hint") or ""),
                    "discovery_source": "chatgpt_extract",
                    "relevance": str(rel),
                }
            )
            n_keep += 1
        _log(
            f"    → substep 3a.{idx}: kept {n_keep}/{len(chunk)} "
            f"(running total extracted={len(kept)})"
        )
    _log(f"    → substep 3a done: {len(kept)} market-related companies from web")
    return kept


def gpt_verify_market(
    client: OpenAI,
    model: str,
    *,
    query: str,
    family: str,
    companies: list[dict[str, str]],
    ckpt: ExpandCheckpoint | None = None,
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    """Second pass: keep only companies related to the market."""
    if not companies:
        return [], []

    roles = ROLE_HINTS.get(family, ROLE_HINTS["general"])
    kept: list[dict[str, str]] = list(ckpt.data("verified") or []) if ckpt else []
    rejected: list[dict[str, str]] = list(ckpt.data("rejected_mid") or []) if ckpt else []
    chunk_size = 20
    start_i = int(ckpt.progress_get("verify_offset", 0) or 0) if ckpt else 0
    total_chunks = (len(companies) + chunk_size - 1) // chunk_size or 1
    _log(
        f"    → substep 4a: verify {len(companies)} candidates "
        f"in {total_chunks} ChatGPT chunk(s)"
        + (f" (resume offset={start_i}, kept={len(kept)})" if start_i else "")
    )
    for i in range(start_i, len(companies), chunk_size):
        chunk = companies[i : i + chunk_size]
        idx = i // chunk_size + 1
        _log(f"    → substep 4a.{idx}: verify chunk {idx}/{total_chunks}…")
        payload = [
            {
                "name": c.get("name") or c.get("Company"),
                "website": c.get("website") or c.get("Website"),
                "snippet": (c.get("snippet") or c.get("Summary") or "")[:200],
                "role_hint": c.get("discovery_source"),
            }
            for c in chunk
        ]
        data = _chat_json(
            client,
            model,
            system=(
                f"You verify whether each company belongs in the CHANNEL landscape. "
                f"Valid roles: {roles}. "
                "KEEP: distributors, wholesalers, dealers, VADs, resellers, importers, "
                "system integrators, authorized channel partners. "
                "REJECT: pure OEMs/manufacturers (even if they also sell direct), "
                "utilities/IPPs/project developers as primary business, media, "
                "associations, and unrelated companies. "
                "When unsure but clearly a channel/distributor, keep=true. "
                "When clearly an OEM brand, related=false. "
                "Return compact JSON: {\"results\":[{\"name\":\"...\",\"related\":true|false,"
                "\"confidence\":0-100,\"reason\":\"...\"}]}"
            ),
            user=(
                f"Market: {query}\n"
                f"Verify these companies:\n{json.dumps(payload, ensure_ascii=False)}"
            ),
            label=f"verify-{idx}",
            max_tokens=3000,
        )
        results = data.get("results") if isinstance(data, dict) else []
        by_name = {_norm(str(c.get("name") or c.get("Company") or "")): c for c in chunk}
        if not isinstance(results, list):
            kept.extend(chunk)
            continue
        seen: set[str] = set()
        for r in results:
            if not isinstance(r, dict):
                continue
            name = str(r.get("name") or "").strip()
            key = _norm(name)
            src = by_name.get(key)
            if not src:
                continue
            seen.add(key)
            related = bool(r.get("related"))
            try:
                conf = int(r.get("confidence") or 0)
            except (TypeError, ValueError):
                conf = 0
            if related and conf >= 55:
                src = dict(src)
                src["verify_reason"] = str(r.get("reason") or "")
                src["verify_confidence"] = str(conf)
                kept.append(src)
            else:
                rejected.append(
                    {
                        "name": name,
                        "reason": str(r.get("reason") or "not related"),
                        "confidence": str(conf),
                    }
                )
        for key, src in by_name.items():
            if key not in seen:
                kept.append(src)
        _log(
            f"    → substep 4a.{idx}: running kept={len(kept)} "
            f"rejected={len(rejected)}"
        )
        if ckpt:
            ckpt.bump(
                "4_verify",
                f"4a.{idx}",
                progress={"verify_offset": i + chunk_size, "verify_chunk": idx},
                verified=kept,
                rejected_mid=rejected,
                note=f"verify chunk {idx}",
            )
    _log(
        f"    → substep 4a done: kept={len(kept)} rejected={len(rejected)}"
    )
    if ckpt:
        ckpt.mark_step_done("4_verify", verified=kept, rejected_mid=rejected)
    return kept, rejected


async def gpt_fill_rows(
    client: OpenAI,
    model: str,
    *,
    query: str,
    family: str,
    rows: list[dict[str, str]],
    web_fill: bool = True,
    ckpt: ExpandCheckpoint | None = None,
    use_linkedin_enrich: bool = True,
    use_apollo: bool = True,
) -> list[dict[str, str]]:
    """Fill all 35 landscape columns via ChatGPT.

    When ``web_fill`` is True (default), each company is researched with the
    OpenAI Responses API ``web_search`` tool so columns are grounded in live web
    results instead of model memory alone. Fact fields are never invented.

    After web fill, optional LinkedIn gap-fill (OpenAI web_search only — no MCP)
    and Apollo contact lookup (partial name → full name via OpenAI LinkedIn web search).
    """
    if not rows:
        return rows

    from vendor_intel.pipeline.openai_fill_enrich import enrich_row_linkedin_apollo

    # Hosted web_search = 1 company/call; DeepSeek chat batches more efficiently
    use_hosted_web = bool(web_fill) and _supports_hosted_web_search()
    chunk_size = 1 if use_hosted_web else 4
    filled: list[dict[str, str]] = list(ckpt.data("filled_part") or []) if ckpt else []
    done_names = {_norm(str(r.get("Company") or "")) for r in filled if r.get("Company")}
    pending = [
        r
        for r in rows
        if _norm(str(r.get("Company") or r.get("name") or "")) not in done_names
    ]
    total_chunks = (len(rows) + chunk_size - 1) // chunk_size or 1
    mode = (
        "OpenAI web_search (real facts only)"
        if use_hosted_web
        else (
            "DeepSeek chat + search evidence (batched)"
            if _is_deepseek()
            else "ChatGPT knowledge"
        )
    )
    _log(
        f"    → substep 5a: fill {len(rows)} companies via {mode} "
        f"in {total_chunks} call(s) — Founded/HQ/brands ONLY from web evidence"
        + (
            f" (resume: {len(filled)} done, {len(pending)} left)"
            if filled
            else ""
        )
    )
    rows = pending  # only process remaining
    if not rows:
        _log("    → substep 5a: all rows already filled (checkpoint) — skip")
        if ckpt:
            ckpt.mark_step_done("5_fill", filled_part=filled)
        return filled
    total_chunks = (len(rows) + chunk_size - 1) // chunk_size or 1
    already_done = len(filled)
    system = (
        "You are a strict vendor-intelligence researcher. "
        "Use LIVE WEB SEARCH. Prefer official company websites, LinkedIn company pages, "
        "reputable directories, and news. "
        "REAL-DATA RULES (critical):\n"
        "1) Founded, Headquarters, Ownership, Employees, Website, Key Brands Represented, "
        "Continent / Geography, Operational Presence — ONLY values you can support from "
        "web results. If not found, use exactly \"Not publicly disclosed\". NEVER invent "
        "a year, city, or brand list. For Founded: use the business founding year "
        "(when the company/trade began), NOT a later incorporation, IPO, rebrand, or "
        "subsidiary registration year. Never copy another brand's year "
        "(e.g. do not give Rexel WESCO's 1922).\n"
        "2) Continent / Geography = ALL countries with a PHYSICAL footprint (HQ, offices, "
        "warehouses, plants, retail stores), not just one — REGION(Country[, Country…]) "
        "e.g. APAC(India, Singapore); EU(Germany, France). "
        "Operational Presence = ALL countries where the company actually SELLS OR SERVES "
        "customers (commercial reach — distribution, resale, service/shipping coverage), "
        "as Country; Country. Operational Presence is usually the SAME AS OR BROADER than "
        "Continent / Geography, since any country with a physical footprint is obviously "
        "also served — include every Continent / Geography country PLUS any additional "
        "country you find commercial-reach evidence for, not only the first. A company "
        "with offices or sales reach in multiple countries/regions must show ALL of them; "
        "do not stop after finding one.\n"
        "3) Core Categories = [Main Product Category] + [Sub Categories] — broad product "
        "groups this distributor handles, main/broadest category first then narrower "
        "sub-categories, semicolon-separated: \"Category 1; Category 2; Category 3\" "
        "(3-8 items). Key Brands Represented = ONLY manufacturer/brand names THIS "
        "company actually sells/distributes/represents (from their site or reliable "
        "listing), semicolon-separated: \"Brand A; Brand B; Brand C\" — distributor "
        "partners often represent multiple manufacturers, list every one evidenced. "
        "Do NOT paste generic market brand lists or a single unrelated brand. "
        "Specialty Focus = EXACT structure, in order: [Product Category] + [Specific "
        "Application / Customer Segment] + [Value-added expertise (optional)], e.g. "
        "\"Industrial rupture discs for oil & gas pressure-relief systems, with in-house "
        "engineering and testing support.\" Include the third clause only when a real "
        "value-add is evidenced; otherwise stop after the second part.\n"
        "4) Contact Person, Role, Email, LinkedIn, Office No., Office Email — leave EMPTY "
        "string if not evidenced. Email = the sales/procurement contact's personal email; "
        "Office Email = the company's general/info/sales@ inbox — these are different "
        "fields, do not duplicate one into the other unless both are separately evidenced. "
        "Do NOT invent personal emails, office emails, phones, or LinkedIn URLs.\n"
        "5) Country Code = ISO2 from HQ (e.g. IN). Region Code = NA|EU|ME|AF|LATAM|APAC.\n"
        "6) Online / Offline / Both = Online | Offline | Both when evidenced — Online means "
        "customers can complete a purchase directly on the company's own website "
        "(cart/checkout); Offline means sales happen via phone/email/rep/branch visit only; "
        "Both means it has real online ordering AND meaningful offline/rep sales.\n"
        "7) Run one Google-style ask per column: \"founded year of {Company}\", "
        "\"headquarters of {Company}\", "
        "\"sales or procurement contact person name of {Company}\", "
        "\"general office email of {Company}\". "
        "For geography specifically, run SEVERAL asks, not one, and keep the two concepts "
        "separate: for physical footprint ask \"offices of {Company}\", \"{Company} "
        "manufacturing plants\", \"{Company} warehouses\", THEN separately probe major "
        "regions, e.g. \"{Company} offices Europe\", \"{Company} offices Asia\", "
        "\"{Company} offices North America\" — combine every country found into "
        "Continent / Geography. For commercial reach ask \"{Company} countries served\", "
        "\"{Company} distributors\", \"{Company} markets served\", \"{Company} ships to\" "
        "— combine every country found PLUS every Continent / Geography country into "
        "Operational Presence.\n"
        "Contact Person/Role = sales or procurement only — never CEO/founder/HR.\n"
        "Every row MUST include ALL keys (empty string for unknown contacts; "
        "Not publicly disclosed for other unknown facts). "
        f"Required keys exactly: {json.dumps(HEADERS)}. "
        "Summary must state how the company relates to the market using web evidence. "
        'Return JSON: {"rows":[{...one object per company...}]}'
    )

    for i in range(0, len(rows), chunk_size):
        chunk = rows[i : i + chunk_size]
        idx = i // chunk_size + 1
        names = ", ".join(
            str(r.get("Company") or r.get("name") or "?")[:40] for r in chunk
        )
        _log(f"    → substep 5a.{idx}: fill {idx}/{total_chunks}: {names}")
        slim = []
        evidence_blocks: list[str] = []
        for r in chunk:
            cname = str(r.get("Company") or r.get("name") or "")
            cweb = str(r.get("Website") or r.get("website") or "")
            slim.append(
                {
                    "Company": cname,
                    "Website": cweb,
                    "Headquarters": r.get("Headquarters") or "",
                    "why_related": r.get("verify_reason") or r.get("snippet"),
                    "relevance": r.get("relevance") or r.get("verify_confidence"),
                }
            )
            if web_fill and cname:
                ev = await _local_web_evidence(cname, cweb, query)
                if ev and not ev.startswith("(local evidence unavailable"):
                    evidence_blocks.append(f"### Evidence for {cname}\n{ev}")
                    _log(f"      · local SERP evidence: {len(ev.splitlines())} hits")
        user = (
            f"Market: {query}\n"
            f"Web-search each company and fill a COMPLETE {len(HEADERS)}-column row "
            "(slim landscape schema).\n"
            "FACT fields (Founded, Headquarters, Key Brands Represented, Continent / "
            "Geography, Operational Presence, etc.) must be REAL from the web or "
            "\"Not publicly disclosed\" — never guessed. Contacts stay empty if unknown.\n"
            f"Companies:\n{json.dumps(slim, ensure_ascii=False)}\n"
        )
        if evidence_blocks:
            user += (
                "\nExtra SERP evidence (use if consistent with your web_search):\n"
                + "\n".join(evidence_blocks)
                + "\n"
            )
        if web_fill:
            data = _web_search_json(
                client,
                model,
                system=system,
                user=user,
                label=f"web-fill-{idx}",
                max_tokens=3500,
            )
            source_tag = "openai_sdk_web_search+real_facts"
        else:
            data = _chat_json(
                client,
                model,
                system=system.replace(
                    "SEARCH THE LIVE WEB for each company",
                    "Use best public knowledge for each company",
                ),
                user=user,
                label=f"fill-{idx}",
                max_tokens=4500,
            )
            source_tag = "openai_chatgpt_sdk_fill"

        out_rows = data.get("rows") if isinstance(data, dict) else None
        if isinstance(data, dict) and not out_rows and data.get("Company"):
            out_rows = [data]
        if not isinstance(out_rows, list):
            for r in chunk:
                name = str(r.get("Company") or r.get("name") or "").strip()
                web = str(r.get("Website") or r.get("website") or "").strip()
                filled.append(
                    _ensure_complete_row(
                        {},
                        family=family,
                        name=name,
                        website=web,
                        source="chatgpt_fill_fallback",
                    )
                )
            continue

        by_name = {
            _norm(str(r.get("Company") or r.get("name") or "")): r for r in chunk
        }
        chunk_filled: list[dict[str, str]] = []
        for gr in out_rows:
            if not isinstance(gr, dict):
                continue
            name = str(gr.get("Company") or "").strip()
            base_src = by_name.get(_norm(name))
            if not name and base_src:
                name = str(base_src.get("Company") or base_src.get("name") or "")
            web = str(
                gr.get("Website")
                or (base_src or {}).get("Website")
                or (base_src or {}).get("website")
                or ""
            )
            merged = _ensure_complete_row(
                {**({k: base_src.get(k) for k in HEADERS if base_src and base_src.get(k)}), **gr},
                family=family,
                name=name,
                website=web,
                source=source_tag,
            )
            if base_src and base_src.get("verify_reason"):
                why = base_src["verify_reason"]
                if why and why not in (merged.get("Summary") or ""):
                    merged["Summary"] = (
                        f"{merged.get('Summary', '').rstrip()} Related: {why}".strip()
                    )
            # LinkedIn gap-fill + Apollo contact (partial → full name)
            if web_fill and (use_linkedin_enrich or use_apollo):
                _log(
                    f"      · enrich LinkedIn/Apollo for {merged.get('Company', '?')}…"
                )
                try:
                    merged = await enrich_row_linkedin_apollo(
                        client,
                        model,
                        query=query,
                        row=merged,
                        web_search_json=_web_search_json,
                        use_linkedin=use_linkedin_enrich,
                        use_apollo=use_apollo,
                        label_prefix=f"enrich-{idx}",
                    )
                    merged = _ensure_complete_row(
                        merged,
                        family=family,
                        name=str(merged.get("Company") or name),
                        website=str(merged.get("Website") or web),
                        source=str(merged.get("Data Sources") or source_tag),
                    )
                except Exception as err:  # noqa: BLE001
                    _log(
                        f"      · enrich failed: {type(err).__name__}: {err}"
                    )
            chunk_filled.append(merged)

        got = {_norm(r.get("Company", "")) for r in chunk_filled}
        for r in chunk:
            key = _norm(str(r.get("Company") or r.get("name") or ""))
            if key and key not in got:
                name = str(r.get("Company") or r.get("name") or "")
                web = str(r.get("Website") or r.get("website") or "")
                chunk_filled.append(
                    _ensure_complete_row(
                        r if r.get("Company") else {},
                        family=family,
                        name=name,
                        website=web,
                        source="chatgpt_fill_miss",
                    )
                )
        filled.extend(chunk_filled)
        _log(
            f"    → substep 5a.{idx}: filled {len(chunk_filled)} "
            f"(running total={len(filled)} / this pass {already_done + idx})"
        )
        if ckpt:
            ckpt.bump(
                "5_fill",
                f"5a.{len(filled)}",
                progress={"fill_count": len(filled)},
                filled_part=filled,
                to_fill_names=[
                    str(r.get("Company") or r.get("name") or "") for r in rows
                ],
                note=f"filled {len(filled)} rows",
            )

    # Second pass: re-web-search rows that still look sparse
    refill_done = set(ckpt.progress_get("refill_done_names") or []) if ckpt else set()
    if web_fill:
        weak_idx = [
            i
            for i, r in enumerate(filled)
            if _row_needs_web_refill(r)
            and _norm(str(r.get("Company") or "")) not in refill_done
        ]
        if weak_idx:
            _log(
                f"    → substep 5b: web re-fill {len(weak_idx)} sparse row(s)…"
            )
            for n, i in enumerate(weak_idx, 1):
                r = filled[i]
                name = r.get("Company") or ""
                _log(f"    → substep 5b.{n}: re-fill {name}")
                data = _web_search_json(
                    client,
                    model,
                    system=system,
                    user=(
                        f"Market: {query}\n"
                        f"This row has missing/weak fields. Web-search and COMPLETE every "
                        f"column for:\n{json.dumps({h: r.get(h) for h in HEADERS}, ensure_ascii=False)}\n"
                        'Return JSON: {"rows":[one complete row]}'
                    ),
                    label=f"web-refill-{n}",
                    max_tokens=3500,
                )
                out_rows = data.get("rows") if isinstance(data, dict) else None
                if isinstance(data, dict) and data.get("Company") and not out_rows:
                    out_rows = [data]
                if isinstance(out_rows, list) and out_rows and isinstance(out_rows[0], dict):
                    filled[i] = _ensure_complete_row(
                        out_rows[0],
                        family=family,
                        name=name,
                        website=str(r.get("Website") or ""),
                        source="openai_sdk_web_search+refill",
                    )
                refill_done.add(_norm(name))
                if ckpt:
                    ckpt.bump(
                        "5_fill",
                        f"5b.{n}",
                        progress={"refill_done_names": list(refill_done)},
                        filled_part=filled,
                        note=f"refill {name}",
                    )

    # Final guarantee: zero empty cells
    filled = [
        _ensure_complete_row(
            r,
            family=family,
            name=str(r.get("Company") or ""),
            website=str(r.get("Website") or ""),
            source=str(r.get("Data Sources") or "openai_sdk_fill"),
        )
        for r in filled
    ]
    _log(f"    → substep 5a done: {len(filled)} rows, all {len(HEADERS)} columns filled")
    if ckpt:
        ckpt.mark_step_done("5_fill", filled_part=filled)
    return filled


def gpt_final_verify_rows(
    client: OpenAI,
    model: str,
    *,
    query: str,
    family: str,
    rows: list[dict[str, str]],
    ckpt: ExpandCheckpoint | None = None,
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    """Final pass on filled rows: drop not-related; keep related."""
    if not rows:
        return [], []

    roles = ROLE_HINTS.get(family, ROLE_HINTS["general"])
    kept: list[dict[str, str]] = list(ckpt.data("final_kept") or []) if ckpt else []
    rejected: list[dict[str, str]] = list(ckpt.data("rejected_final") or []) if ckpt else []
    chunk_size = 10
    start_i = int(ckpt.progress_get("final_verify_offset", 0) or 0) if ckpt else 0
    total_chunks = (len(rows) + chunk_size - 1) // chunk_size or 1
    _log(
        f"    → substep 6a: final QA on {len(rows)} rows "
        f"in {total_chunks} ChatGPT chunk(s)"
        + (f" (resume offset={start_i})" if start_i else "")
    )
    for i in range(start_i, len(rows), chunk_size):
        chunk = rows[i : i + chunk_size]
        idx = i // chunk_size + 1
        _log(f"    → substep 6a.{idx}: final-verify chunk {idx}/{total_chunks}…")
        payload = [
            {
                "Company": r.get("Company"),
                "Website": r.get("Website"),
                "Core Categories": r.get("Core Categories"),
                "Specialty Focus": r.get("Specialty Focus"),
                "Summary": (r.get("Summary") or "")[:180],
            }
            for r in chunk
        ]
        try:
            data = _chat_json(
                client,
                model,
                system=(
                    f"Final QA for CHANNEL landscape. Valid roles: {roles}. "
                    "KEEP only distributors / wholesalers / dealers / VADs / resellers / "
                    "importers / system integrators / channel partners. "
                    "DROP pure OEMs/manufacturers, utilities/IPPs/developers-as-primary, "
                    "and rows not related to the market. "
                    "Return compact JSON: {\"results\":[{\"Company\":\"...\",\"keep\":true|false,"
                    "\"reason\":\"...\"}]}"
                ),
                user=f"Market: {query}\nRows:\n{json.dumps(payload, ensure_ascii=False)}",
                label=f"final-{idx}",
                max_tokens=2500,
            )
        except Exception as err:  # noqa: BLE001
            # DeepSeek/OpenAI outage (e.g. insufficient balance) must not crash the
            # whole pipeline run — keep this chunk's rows unverified rather than lose
            # everything already gathered in Steps 1-6.
            _log(f"    → substep 6a.{idx}: final-verify unavailable ({err}) — keeping chunk as-is")
            data = {}
        results = data.get("results") if isinstance(data, dict) else []
        by = {_norm(str(r.get("Company") or "")): r for r in chunk}
        if not isinstance(results, list):
            kept.extend(chunk)
            continue
        seen: set[str] = set()
        for r in results:
            if not isinstance(r, dict):
                continue
            name = str(r.get("Company") or "").strip()
            key = _norm(name)
            src = by.get(key)
            if not src:
                continue
            seen.add(key)
            if r.get("keep") is False:
                rejected.append({"company": name, "reason": str(r.get("reason") or "")})
            else:
                src = dict(src)
                reason = str(r.get("reason") or "").strip()
                if reason:
                    src["Data Confidence"] = src.get("Data Confidence") or "Medium"
                    src["Summary"] = (src.get("Summary") or "").rstrip()
                    if reason not in src["Summary"]:
                        src["Summary"] = f"{src['Summary']} Verify: {reason}".strip()
                kept.append(src)
        for key, src in by.items():
            if key not in seen:
                kept.append(src)
        _log(
            f"    → substep 6a.{idx}: running kept={len(kept)} "
            f"rejected={len(rejected)}"
        )
        if ckpt:
            ckpt.bump(
                "6_final",
                f"6a.{idx}",
                progress={"final_verify_offset": i + chunk_size, "final_chunk": idx},
                final_kept=kept,
                rejected_final=rejected,
                note=f"final verify chunk {idx}",
            )
    _log(
        f"    → substep 6a done: kept={len(kept)} rejected={len(rejected)}"
    )
    if ckpt:
        ckpt.mark_step_done("6_final", final_kept=kept, rejected_final=rejected)
    return kept, rejected


def _as_landscape_rows(
    verified: list[dict[str, str]], family: str
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for c in verified:
        if c.get("Company"):
            rows.append(c)
            continue
        name = str(c.get("name") or "").strip()
        web = str(c.get("website") or "").strip()
        row = _blank_row(name, web, family, c.get("discovery_source") or "chatgpt")
        if c.get("source_query") and c["source_query"] not in {
            "chatgpt_recall",
            "",
        }:
            row["Headquarters"] = c["source_query"]
            row["Continent / Geography"] = c["source_query"]
        if c.get("snippet"):
            row["Specialty Focus"] = c["snippet"][:180]
        rows.append(row)
    return rows


async def run_chatgpt_expand(
    query: str,
    *,
    country: str = "global",
    target: int = 320,
    batch: str = "all",
    max_queries: int = 60,
    merge_existing: bool = True,
    final_path: str | None = None,
    out_dir: str | None = None,
    fill_existing: bool = False,
    max_fill: int = 500,
    seeds_only: bool = False,
    seeds_path: str | None = None,
    skip_recall: bool = False,
    web_fill: bool = True,
    openai_discover: bool = True,
    ddgs_harvest: bool = False,
    resume: bool = True,
    use_linkedin_enrich: bool = True,
    use_apollo: bool = True,
    fill_backend: str = "search",
    force: bool = False,
    family: str | None = None,
) -> dict[str, Any]:
    """Full expand: OpenAI/DeepSeek SDK for find+verify; search stack (default) for fill.

    Same steps 0–7 for ANY market (batch of 1…N). No demo-market hardcoding.

    DeepSeek: no hosted web_search — auto-enables ddgs harvest, keeps fill_backend=search,
    and uses chat.completions for verify / residual (efficient, no wasted Responses retries).
    """
    family = market_family(query, family=family)

    # DeepSeek-efficient defaults (OpenAI-compatible chat; live web via free search stack)
    # EXPAND_SKIP_DDGS_HARVEST=1 or ddgs_harvest=False with skip env → never force Step 3.
    skip_ddgs = (os.getenv("EXPAND_SKIP_DDGS_HARVEST") or "").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )
    if skip_ddgs:
        ddgs_harvest = False

    if _is_deepseek():
        if str(fill_backend).lower() in ("openai", "web", "responses"):
            _log(
                "  [chatgpt] DeepSeek: fill_backend forced to 'search' "
                "(no OpenAI hosted web_search)"
            )
            fill_backend = "search"
        if not ddgs_harvest and not seeds_only and not skip_ddgs:
            ddgs_harvest = True
            _log(
                "  [chatgpt] DeepSeek: auto-enabled ddgs/SearXNG harvest for discovery"
            )
        elif skip_ddgs:
            _log("  [chatgpt] Step 3 ddgs harvest SKIPPED (EXPAND_SKIP_DDGS_HARVEST)")
        # LinkedIn via OpenAI web_search does not exist on DeepSeek — skip MCP-less path
        if use_linkedin_enrich and not _supports_hosted_web_search():
            use_linkedin_enrich = False
            _log(
                "  [chatgpt] DeepSeek: LinkedIn OpenAI-web_search enrich off "
                "(use Apollo / search stack instead)"
            )
        # Prefer higher residual parallelism on cheap DeepSeek chat
        if not (os.getenv("EXPAND_RESIDUAL_CONCURRENT") or "").strip():
            os.environ["EXPAND_RESIDUAL_CONCURRENT"] = "4"

    out_p = Path(out_dir) if out_dir else default_output_dir(query, country)
    out_p.mkdir(parents=True, exist_ok=True)
    xlsx = resolve_final_path(query, country, out_p, final_path)

    existing = read_final_rows(xlsx) if merge_existing and xlsx.exists() else []
    seeds = _seed_candidates(query, country, seeds_path=seeds_path)
    ckpt, resumed = ExpandCheckpoint.load_or_create(
        out_p,
        query=query,
        country=country,
        batch=batch,
        target=target,
        resume=resume,
        force=force,
        meta={
            "seeds_only": seeds_only,
            "web_fill": web_fill,
            "openai_discover": openai_discover,
            "ddgs_harvest": ddgs_harvest,
            "max_fill": max_fill,
            "use_linkedin_enrich": use_linkedin_enrich,
            "use_apollo": use_apollo,
            "fill_backend": fill_backend,
        },
    )
    # Skip finished markets (required for 10k batch efficiency)
    if resume and not force and ckpt.is_done():
        n_existing = len(existing) if existing else len(read_final_rows(xlsx) if xlsx.exists() else [])
        _log(
            f"  [chatgpt] SKIP done market: {query!r} "
            f"(checkpoint status=done, final_rows≈{n_existing}) → {xlsx}"
        )
        return {
            "skipped": True,
            "reason": "already_done",
            "query": query,
            "country": country,
            "family": family,
            "final_path": str(xlsx),
            "xlsx": str(xlsx),
            "companies": n_existing,
            "total_after": n_existing,
            "checkpoint": str(ckpt.path),
        }

    settings = Settings()
    client = _client(settings)
    model = _model(settings)
    session = openai_cost.start_run(
        market=query,
        model=model,
        family=family,
        batch=batch,
        target=target,
        max_queries=0 if seeds_only else max_queries,
        max_fill=max_fill,
        seeds=len(seeds),
        seeds_only=seeds_only,
        openai_discover=openai_discover and not seeds_only,
        ddgs_harvest=ddgs_harvest and not seeds_only,
        web_fill=web_fill,
        resume=resumed,
        checkpoint=str(ckpt.path),
    )
    _log(
        f"  [chatgpt] FULL {_llm_label(settings)} SDK pipeline (no run_query.py) | "
        f"model={model} | base={_base_url(settings)}"
    )
    _log(f"  [chatgpt] cost run_id={session.run_id} → {openai_cost.cost_dir()}")
    _log(f"  [chatgpt] existing={len(existing)} final={xlsx}")
    _log(
        f"  [chatgpt] curated seeds loaded={len(seeds)}"
        + (" (seeds-only mode)" if seeds_only else "")
    )
    if resumed:
        _log(f"  [chatgpt] RESUME checkpoint: {ckpt.summary_line()}")
    else:
        _log(f"  [chatgpt] checkpoint → {ckpt.path}")
    _log(
        "  [chatgpt] discover mode: "
        + (
            "skip (seeds-only)"
            if seeds_only
            else (
                (
                    "DeepSeek chat discover"
                    if _is_deepseek(settings)
                    else "OpenAI web_search"
                )
                + (" + ddgs harvest" if ddgs_harvest else "")
                if openai_discover
                else ("ddgs harvest" if ddgs_harvest else "none")
            )
        )
    )
    _log(
        f"  [chatgpt] fill mode: "
        + (
            "search stack first (DDGS + SearXNG + Google AI scraper), "
            f"then {_llm_label(settings)} residual for leftover gaps"
            if str(fill_backend).lower() in ("search", "searx", "ddgs", "stack")
            else (
                "OpenAI Responses API web_search (live web per company)"
                if web_fill and _supports_hosted_web_search(settings)
                else f"{_llm_label(settings)} chat knowledge / batched fill"
            )
        )
    )
    _log(
        "  [chatgpt] enrich: "
        + (
            "Step 5a = Google AI + DDGS + SearXNG + Exa + Wikipedia + Owler; "
            "Step 5b = OpenAI SDK residual for remaining NPD/gaps | "
            f"Apollo={'on' if use_apollo and (os.environ.get('APOLLO_API_KEY') or '').strip() else 'off'}"
            if str(fill_backend).lower() in ("search", "searx", "ddgs", "stack")
            else (
                f"LinkedIn web_search={'on' if use_linkedin_enrich else 'off'} (no MCP) | "
                f"Apollo={'on' if use_apollo and (os.environ.get('APOLLO_API_KEY') or '').strip() else 'off'}"
            )
        )
    )

    try:
        ckpt.bump("0_start", "seeds", seeds_count=len(seeds), note="start")

        recalled: list[dict[str, str]] = []
        if ckpt.is_step_done("1_recall"):
            recalled = list(ckpt.data("recalled") or [])
            _log(f"  [chatgpt] Step 1/7: SKIP (checkpoint) — {len(recalled)} recalled")
        elif skip_recall or (seeds_only and len(seeds) >= target):
            _log("  [chatgpt] Step 1/7: skip ChatGPT recall (seeds cover target)")
            ckpt.mark_step_done("1_recall", recalled=[])
        else:
            _log("  [chatgpt] Step 1/7: ChatGPT recall known companies…")
            recalled = gpt_recall_companies(
                client,
                model,
                query=query,
                family=family,
                target=target,
                ckpt=ckpt,
            )

        discovered: list[dict[str, str]] = []
        raw: list[dict[str, str]] = []
        extracted: list[dict[str, str]] = []
        if seeds_only:
            _log("  [chatgpt] Step 2/7: skip web discover (seeds-only)")
            _log("  [chatgpt] Step 3/7: skip ddgs extract (seeds-only)")
            if not ckpt.is_step_done("2_discover"):
                ckpt.mark_step_done("2_discover", discovered=[])
            if not ckpt.is_step_done("3_ddgs"):
                ckpt.mark_step_done("3_ddgs", raw=[], extracted=[])
            openai_cost.heartbeat("seeds_only", seeds=len(seeds), recalled=len(recalled))
        else:
            exclude_names = [
                *(r.get("name") or "" for r in seeds),
                *(r.get("name") or "" for r in recalled),
                *(r.get("Company") or "" for r in existing),
            ]
            need = max(0, target - len(seeds) - len(recalled))
            # Scale page budget with target (800–1000 needs ~50–70 pages @ 25/page)
            discover_pages = max(
                8,
                min(80, max(max_queries // 4 or 12, (need + 24) // 25 + 10)),
            )

            if ckpt.is_step_done("2_discover"):
                discovered = list(ckpt.data("discovered") or [])
                _log(
                    f"  [chatgpt] Step 2/7: SKIP (checkpoint) — "
                    f"{len(discovered)} openai-discovered"
                )
            elif openai_discover and need > 0:
                _log(
                    "  [chatgpt] Step 2/7: OpenAI SDK web_search — find companies…"
                )
                openai_cost.heartbeat(
                    "openai_discover",
                    discover_pages=discover_pages,
                    need=need,
                    recalled=len(recalled),
                    seeds=len(seeds),
                )
                discovered = gpt_discover_companies(
                    client,
                    model,
                    query=query,
                    family=family,
                    target=need,
                    exclude=exclude_names,
                    max_pages=discover_pages,
                    ckpt=ckpt,
                )
                (out_p / f"chatgpt_openai_discover_batch_{batch.lower()}.json").write_text(
                    json.dumps(discovered, indent=2, ensure_ascii=False),
                    encoding="utf-8",
                )
                openai_cost.heartbeat(
                    "openai_discover_done",
                    openai_discovered=len(discovered),
                )
            elif openai_discover:
                _log(
                    "  [chatgpt] Step 2/7: skip OpenAI discover "
                    f"(seeds+recall already ≥ target: {len(seeds)+len(recalled)})"
                )
                ckpt.mark_step_done("2_discover", discovered=[])
            else:
                _log("  [chatgpt] Step 2/7: OpenAI discover disabled")
                ckpt.mark_step_done("2_discover", discovered=[])

            if ckpt.is_step_done("3_ddgs"):
                raw = list(ckpt.data("raw") or [])
                extracted = list(ckpt.data("extracted") or [])
                _log(
                    f"  [chatgpt] Step 3/7: SKIP (checkpoint) — "
                    f"{len(extracted)} ddgs extracted"
                )
            elif ddgs_harvest:
                _log("  [chatgpt] Step 3/7: optional ddgs/SearXNG harvest + extract…")
                openai_cost.heartbeat(
                    "ddgs_harvest",
                    web_queries_planned=max_queries,
                )
                raw = await web_harvest_candidates(
                    query,
                    country=country,
                    batch=batch,
                    max_queries=max_queries,
                )
                ckpt.bump("3_ddgs", "3a_harvest", raw=raw, note="ddgs harvest done")
                (out_p / f"chatgpt_harvest_batch_{batch.lower()}.json").write_text(
                    json.dumps(raw, indent=2, ensure_ascii=False), encoding="utf-8"
                )
                extracted = gpt_extract_and_score(
                    client, model, query=query, family=family, candidates=raw
                )
                ckpt.mark_step_done("3_ddgs", raw=raw, extracted=extracted)
            else:
                _log(
                    "  [chatgpt] Step 3/7: skip ddgs harvest "
                    "(use --ddgs-harvest to enable; OpenAI discover is primary)"
                )
                ckpt.mark_step_done("3_ddgs", raw=[], extracted=[])

        combined = seeds + recalled + discovered + extracted
        if ckpt.is_step_done("4_verify"):
            verified = list(ckpt.data("verified") or [])
            rejected_mid = list(ckpt.data("rejected_mid") or [])
            _log(
                f"  [chatgpt] Step 4/7: SKIP (checkpoint) — "
                f"kept={len(verified)} rejected={len(rejected_mid)}"
            )
        else:
            _log(
                f"  [chatgpt] Step 4/7: ChatGPT verify market-related "
                f"({len(combined)} = {len(seeds)} seeds + {len(recalled)} recalled "
                f"+ {len(discovered)} openai-web + {len(extracted)} ddgs)…"
            )
            ckpt.bump("4_verify", "4a_start", combined_count=len(combined))
            verified, rejected_mid = gpt_verify_market(
                client,
                model,
                query=query,
                family=family,
                companies=combined,
                ckpt=ckpt,
            )
            _log(f"    → substep 4b: write rejected list ({len(rejected_mid)})")
            (out_p / f"chatgpt_rejected_mid_batch_{batch.lower()}.json").write_text(
                json.dumps(rejected_mid, indent=2, ensure_ascii=False), encoding="utf-8"
            )

        # Merge verified into landscape-shaped rows with existing FINAL
        if ckpt.is_step_done("4b_merge") and ckpt.data("merged") and ckpt.data("to_fill"):
            merged = list(ckpt.data("merged") or [])
            to_fill = list(ckpt.data("to_fill") or [])
            _log(
                f"    → substep 4c: SKIP merge (checkpoint) — "
                f"merged={len(merged)} to_fill={len(to_fill)}"
            )
        else:
            landscape_new = _as_landscape_rows(verified, family)
            merge_in = [
                {
                    "name": r["Company"],
                    "website": r.get("Website", ""),
                    "domain": "",
                    "snippet": r.get("Specialty Focus", ""),
                    "source_query": r.get("Headquarters", ""),
                    "discovery_source": r.get("Data Sources", "chatgpt"),
                }
                for r in landscape_new
                if r.get("Company")
            ]
            merged = merge_candidates(
                existing, merge_in, family=family, target=max(target * 2, target)
            )
            by = {_norm(r.get("Company", "")): r for r in merged}
            for r in landscape_new:
                k = _norm(r.get("Company", ""))
                if k in by:
                    for h in HEADERS:
                        if r.get(h) and (
                            not by[k].get(h)
                            or by[k].get(h) == "Not publicly disclosed"
                            or (by[k].get("Data Sources") or "").startswith("web_expand")
                        ):
                            by[k][h] = r[h]
            merged = list(by.values())

            if fill_existing:
                to_fill = list(merged)
            else:
                existing_names = {_norm(r.get("Company", "")) for r in existing}
                new_rows = [
                    r for r in merged if _norm(r.get("Company", "")) not in existing_names
                ]
                weak = [
                    r
                    for r in merged
                    if "chatgpt" not in (r.get("Data Sources") or "").lower()
                    or (r.get("Headquarters") or "") == "Not publicly disclosed"
                ]
                seen: set[str] = set()
                to_fill = []
                for r in new_rows + weak:
                    k = _norm(r.get("Company", ""))
                    if not k or k in seen:
                        continue
                    seen.add(k)
                    to_fill.append(r)
            fill_cap = max(int(max_fill or 0), int(target or 0), 1)
            if len(to_fill) > fill_cap:
                to_fill = to_fill[:fill_cap]
            _log(
                f"  [chatgpt] fill ALL table rows this pass: {len(to_fill)} "
                f"(merged={len(merged)}, cap={fill_cap})"
            )
            ckpt.mark_step_done("4b_merge", merged=merged, to_fill=to_fill)

        if ckpt.is_step_done("5_fill") and ckpt.data("filled_part"):
            filled_part = list(ckpt.data("filled_part") or [])
            _log(
                f"  [chatgpt] Step 5/7: SKIP (checkpoint) — {len(filled_part)} filled"
            )
        else:
            use_search_fill = str(fill_backend).lower() in (
                "search",
                "searx",
                "ddgs",
                "stack",
            )
            if use_search_fill:
                from vendor_intel.pipeline.search_stack_fill import search_stack_fill_rows

                _log(
                    f"  [chatgpt] Step 5/7: Google AI scraper fill FIRST (all SLIM columns), "
                    f"then {_llm_label()} residual for leftover empty / "
                    f"Not publicly disclosed gaps — {len(to_fill)} rows…"
                )
                filled_part = await search_stack_fill_rows(
                    query=query,
                    family=family,
                    rows=to_fill,
                    ckpt=ckpt,
                    sources=None,  # SEARCH_STACK_SOURCES (default: google_ai only)
                    use_apollo=use_apollo,
                )
            else:
                _log(
                    f"  [chatgpt] Step 5/7: "
                    f"{'OpenAI web_search + ' if web_fill else ''}"
                    f"fill all 35 columns ({len(to_fill)} rows)…"
                )
                filled_part = await gpt_fill_rows(
                    client,
                    model,
                    query=query,
                    family=family,
                    rows=to_fill,
                    web_fill=web_fill,
                    ckpt=ckpt,
                    use_linkedin_enrich=use_linkedin_enrich,
                    use_apollo=use_apollo,
                )
        by = {_norm(r.get("Company", "")): r for r in merged}
        for r in filled_part:
            by[_norm(r.get("Company", ""))] = r
        filled_all = [
            _ensure_complete_row(
                r,
                family=family,
                name=str(r.get("Company") or ""),
                website=str(r.get("Website") or ""),
                source=str(r.get("Data Sources") or "openai_sdk"),
            )
            for r in by.values()
        ]
        ckpt.bump("5_fill", "5c_merge", filled_all=filled_all, note="merged filled_all")
        _log(f"    → substep 5c: merged filled rows into list ({len(filled_all)} total)")

        if ckpt.is_step_done("5d_field_verify") and ckpt.data("field_verified_rows") is not None:
            filled_all = list(ckpt.data("field_verified_rows") or [])
            rejected_field_verify = list(ckpt.data("rejected_field_verify") or [])
            _log(
                f"  [chatgpt] Step 6/7: SKIP (checkpoint) — "
                f"kept={len(filled_all)} rejected={len(rejected_field_verify)}"
            )
        else:
            _log(
                f"  [chatgpt] Step 6/7: verify/correct Core Categories, Specialty Focus, "
                f"Key Brands, channel fit, Online/Offline/Both, Contact Person/Role/LinkedIn, "
                f"geography — {len(filled_all)} rows…"
            )
            try:
                from vendor_intel.pipeline.field_verify_fill import verify_and_correct_fields

                filled_all, rejected_field_verify = await verify_and_correct_fields(
                    filled_all,
                    query=query,
                    family=family,
                    ckpt=ckpt,
                    use_apollo=use_apollo,
                )
            except Exception as err:  # noqa: BLE001
                _log(f"    → Step 6/7: field verify/correct failed, keeping Step 5 output: {err}")
                rejected_field_verify = []
            ckpt.mark_step_done(
                "5d_field_verify",
                field_verified_rows=filled_all,
                rejected_field_verify=rejected_field_verify,
            )
            _log(
                f"    → Step 6/7 done: kept={len(filled_all)} "
                f"rejected={len(rejected_field_verify)}"
            )

        if ckpt.is_step_done("6_final") and ckpt.data("final_kept"):
            final_rows = list(ckpt.data("final_kept") or [])
            rejected_final = list(ckpt.data("rejected_final") or [])
            _log(
                f"  [chatgpt] Step 7/7: SKIP (checkpoint) — "
                f"kept={len(final_rows)} rejected={len(rejected_final)}"
            )
        else:
            _log("  [chatgpt] Step 7/7: ChatGPT final market-related verify…")
            final_rows, rejected_final = gpt_final_verify_rows(
                client,
                model,
                query=query,
                family=family,
                rows=filled_all,
                ckpt=ckpt,
            )
            all_rejected = list(rejected_field_verify) + list(rejected_final)
            _log(
                f"    → substep 6b: write final rejects "
                f"({len(rejected_final)} market-relevance + {len(rejected_field_verify)} field-verify)"
            )
            (out_p / f"chatgpt_rejected_final_batch_{batch.lower()}.json").write_text(
                json.dumps(all_rejected, indent=2, ensure_ascii=False), encoding="utf-8"
            )

        final_rows.sort(key=lambda r: _norm(r.get("Company", "")))

        # Channel-partner purity (discovery stayed wide; export is channel-only)
        try:
            from vendor_intel.intelligence.channel_purity import (
                channel_purity_enabled,
                filter_channel_pure_rows_async,
            )

            if channel_purity_enabled():
                before_n = len(final_rows)
                # Auto-map electrical distribution markets to electrical OEM blocklist
                purity_family = family
                try:
                    from vendor_intel.intelligence.channel_purity import _resolve_family

                    purity_family = _resolve_family(family, query)
                except Exception:
                    purity_family = family
                final_rows, purity_dropped = await filter_channel_pure_rows_async(
                    final_rows, family=purity_family, query=query
                )
                _log(
                    f"  [chatgpt] Channel purity ({purity_family}): kept {len(final_rows)}/{before_n} "
                    f"(dropped {len(purity_dropped)} OEM/utility/fake/non-channel)"
                )
                if purity_dropped:
                    (out_p / f"chatgpt_rejected_purity_batch_{batch.lower()}.json").write_text(
                        json.dumps(purity_dropped, indent=2, ensure_ascii=False),
                        encoding="utf-8",
                    )
                    rejected_final = list(rejected_final) + [
                        {"company": d.get("company"), "reason": f"purity:{d.get('reason')}"}
                        for d in purity_dropped
                    ]
        except Exception as exc:  # noqa: BLE001
            _log(f"  [chatgpt] Channel purity skipped: {type(exc).__name__}: {exc}")

        if len(final_rows) > target:

            def _rank(r: dict[str, str]) -> tuple[int, str]:
                dc = (r.get("Data Confidence") or "").lower()
                score = 2 if dc == "high" else 1 if dc == "medium" else 0
                return (-score, _norm(r.get("Company", "")))

            final_rows = sorted(final_rows, key=_rank)[:target]
            final_rows.sort(key=lambda r: _norm(r.get("Company", "")))

        sheet = {
            "firewall": "Firewall Companies",
            "smartwatch": "Smartwatch Companies",
            "green": "Green Chemical Companies",
            "avocado": "Avocado Oil Companies",
        }.get(family, "Companies")

        cost_summary = openai_cost.finish_run(
            companies=len(final_rows),
            xlsx=str(xlsx),
            rejected_mid=len(rejected_mid),
            rejected_final=len(rejected_final),
            status="ok",
        )
        openai_cost.print_session_summary(cost_summary)

        audit = {
            "backend": "chatgpt",
            "uses_run_query": False,
            "model": model,
            "query": query,
            "country": country,
            "batch": batch,
            "family": family,
            "target": target,
            "existing_before": len(existing),
            "seeds_loaded": len(seeds),
            "seeds_only": seeds_only,
            "web_fill": web_fill,
            "fill_backend": fill_backend,
            "openai_discover": openai_discover and not seeds_only,
            "ddgs_harvest": ddgs_harvest and not seeds_only,
            "chatgpt_recalled": len(recalled),
            "openai_discovered": len(discovered),
            "harvested_raw": len(raw),
            "chatgpt_extracted": len(extracted),
            "verified_mid": len(verified),
            "rejected_mid": len(rejected_mid),
            "chatgpt_filled": len(filled_part),
            "rejected_final": len(rejected_final),
            "total_after": len(final_rows),
            "xlsx": str(xlsx),
            "mode": "openai_sdk_full_find_verify_fill",
            "openai_cost": cost_summary,
        }
        audit["checkpoint"] = str(ckpt.path)
        audit["resumed"] = resumed
        _log(f"    → substep 7: write Excel ({len(final_rows)} rows) → {xlsx.name}")
        ckpt.bump("7_write", "7a", note="writing excel")
        write_final_xlsx(xlsx, final_rows, sheet, audit)
        (out_p / f"chatgpt_expand_batch_{batch.lower()}_audit.json").write_text(
            json.dumps(audit, indent=2), encoding="utf-8"
        )
        ckpt.mark_done()
        _log(f"  [chatgpt] DONE (OpenAI SDK only) → {len(final_rows)} rows → {xlsx}")
        _log(f"  [chatgpt] checkpoint marked done → {ckpt.path}")
        return audit
    except Exception:
        # persist partial cost if run crashes mid-way; checkpoint stays resumable
        try:
            ckpt.bump(
                str(ckpt.state.get("step") or "error"),
                "error",
                note="interrupted — re-run same command to --resume",
            )
            _log(f"  [chatgpt] checkpoint saved for resume → {ckpt.path}")
            _log(f"  [chatgpt] {ckpt.summary_line()}")
        except Exception:
            pass
        partial = {
            "run_id": session.run_id,
            "market": query,
            "model": model,
            "calls": len(session.calls),
            "prompt_tokens": session.prompt_tokens,
            "completion_tokens": session.completion_tokens,
            "total_tokens": session.total_tokens,
            "estimated_cost_usd": session.estimated_cost_usd,
        }
        openai_cost.finish_run(status="error")
        openai_cost.print_session_summary(partial)
        raise
