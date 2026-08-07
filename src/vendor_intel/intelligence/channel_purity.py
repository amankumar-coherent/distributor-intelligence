"""Channel-partner purity filter (shared by run_batch + OpenAI expand).

Design:
  - Discovery / recall stay WIDE (market coverage).
  - Before Excel export, drop OEMs, pure manufacturers, utilities/IPPs, and
    junk that are not distributors / VADs / resellers / dealers / wholesalers /
    system integrators / channel partners.

Env:
  CHANNEL_PURITY_FILTER=true|false   (default true)
  CHANNEL_PURITY_MODE=balanced|strict (default balanced)
  CHANNEL_PURITY_LLM=true|false      (default true — DeepSeek/OpenAI via OPENAI_*)
  CHANNEL_PURITY_LLM_CHUNK=20        (companies per LLM call)
  CHANNEL_PURITY_WEB=true|false      (default false; optional SearXNG confirm)
"""
from __future__ import annotations

import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass
from typing import Any


def channel_purity_enabled() -> bool:
    raw = (os.getenv("CHANNEL_PURITY_FILTER") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def channel_purity_mode() -> str:
    raw = (os.getenv("CHANNEL_PURITY_MODE") or "balanced").strip().lower()
    return raw if raw in ("balanced", "strict") else "balanced"


def channel_purity_llm_enabled() -> bool:
    raw = (os.getenv("CHANNEL_PURITY_LLM") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def channel_purity_web_enabled() -> bool:
    raw = (os.getenv("CHANNEL_PURITY_WEB") or "false").strip().lower()
    return raw in ("1", "true", "yes", "on")


def _llm_chunk_size() -> int:
    try:
        return max(5, min(40, int(os.getenv("CHANNEL_PURITY_LLM_CHUNK") or "20")))
    except ValueError:
        return 20


CHANNEL_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|dealer|"
    r"channel partner|value[- ]added distributor|\bvad\b|authorized dealer|"
    r"stockist|importer|system integrator|var\b|mssp|epc contractor|"
    r"solar installer|installation partner|merchant wholesaler|"
    r"electrical distributor|pv distributor|authorized partner)\b",
    re.I,
)
OEM_RE = re.compile(
    r"\b(manufacturer|manufactures|oem\b|module maker|inverter manufacturer|"
    r"turbine manufacturer|produces (?:solar|wind|pv|inverters|modules|devices)|"
    r"leading manufacturer|global manufacturer|tier[- ]?1 manufacturer|"
    r"original equipment)\b",
    re.I,
)
UTILITY_RE = re.compile(
    r"\b(utility|independent power producer|\bipp\b|power producer|asset owner|"
    r"project developer|owns and operates|energy retailer|electricity supplier)\b",
    re.I,
)
NAME_CHANNEL_RE = re.compile(
    r"\b(distributor|distributors|wholesale|wholesaler|dealer|dealers|reseller|"
    r"channel|vad|var)\b",
    re.I,
)

OEM_BLOCKLISTS: dict[str, set[str]] = {
    "renewable": {
        "sungrow", "sungrow power supply", "huawei", "huawei fusionsolar",
        "huawei technologies", "growatt", "growatt new energy", "delta electronics",
        "abb", "siemens", "siemens energy", "siemens gamesa", "general electric",
        "ge vernova", "mitsubishi electric", "toshiba", "panasonic", "lg electronics",
        "lg energy solution", "byd", "rec group", "solaria", "silfab solar",
        "q cells", "hanwha q cells", "hanwha corporation", "first solar",
        "jinko solar", "jinkosolar", "trina solar", "trinasolar", "canadian solar",
        "longi", "longi solar", "ja solar", "enphase", "enphase energy",
        "solaredge", "solaredge technologies", "tesla", "vestas", "nordex",
        "goldwind", "mingyang", "schneider electric", "eaton", "sma",
        "sma solar technology ag", "fronius", "fronius international", "goodwe",
        "catl", "samsung sdi", "orsted", "iberdrola", "nextera energy",
        "nextera energy resources", "enel", "enel green power", "engie",
        "edf renewables", "shell", "totalenergies", "equinor", "rwe",
        "rwe renewables", "acciona energy", "mission solar energy", "risen energy",
        "talesun solar", "yingli solar", "sunpower corporation", "vikram solar",
        "renewsys india", "kyocera corporation", "chint new energy", "tongwei solar",
        "waaree energies", "suzlon energy", "victron energy",
    },
    "firewall": {
        "palo alto networks", "fortinet", "cisco", "check point", "checkpoint",
        "sophos", "juniper", "juniper networks", "forcepoint", "zscaler",
        "crowdstrike", "sonicwall", "watchguard", "barracuda", "f5", "f5 networks",
        "aruba", "hp aruba", "huawei", "sangfor", "hillstone",
    },
    "smartwatch": {
        "apple", "samsung", "garmin", "fitbit", "google", "huawei", "xiaomi",
        "amazfit", "huami", "withings", "fossil", "tag heuer", "polar",
        "suunto", "oppo", "oneplus", "honor",
    },
    "avocado": {
        "unilever", "cargill", "adm", "bunge", "nestle", "procter gamble",
    },
    "green": {
        "basf", "dow", "dupont", "sabic", "lyondellbasell", "exxonmobil",
        "shell chemicals", "ineos", "covestro", "evonik",
    },
    "general": set(),
}

CHANNEL_ROLE_HINTS: dict[str, str] = {
    "renewable": (
        "distributors, wholesalers, dealers, importers, EPCs/system integrators, "
        "authorized channel partners for solar/wind/storage equipment "
        "(NOT module/inverter/turbine OEMs, NOT utilities/IPPs/developers as primary)"
    ),
    "firewall": (
        "distributors, VADs, resellers, MSSPs, system integrators, channel partners "
        "for firewalls / network security (NOT the OEM vendors themselves)"
    ),
    "smartwatch": (
        "distributors, wholesalers, electronics retailers, channel partners for "
        "smartwatches / wearables (NOT Apple/Samsung/Garmin as OEMs)"
    ),
    "avocado": (
        "distributors, wholesalers, importers, specialty oil suppliers, grocery/retail "
        "chains, foodservice suppliers of avocado oil / edible oils"
    ),
    "green": (
        "specialty chemical distributors / ingredients distributors for green / "
        "bio-based chemicals (NOT pure petrochem manufacturers)"
    ),
    "general": (
        "distributors, wholesalers, channel partners, VADs, resellers, dealers, "
        "system integrators relevant to the market (NOT pure OEMs/manufacturers)"
    ),
}


def _norm(s: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


def _scrub_boilerplate(text: str) -> str:
    t = text or ""
    t = re.sub(r"(?i)channel participant[^.]*\.?", " ", t)
    t = re.sub(r"(?i)verify:\s*", " ", t)
    t = re.sub(r"(?i)\bnot publicly disclosed\b", " ", t)
    return t


def _is_blocked_oem(name: str, family: str) -> bool:
    key = _norm(name)
    block: set[str] = set()
    block |= OEM_BLOCKLISTS.get(family, set())
    block |= OEM_BLOCKLISTS.get("general", set())
    if family == "general":
        block |= OEM_BLOCKLISTS.get("renewable", set())
        block |= OEM_BLOCKLISTS.get("firewall", set())
    if key in block:
        return True
    for m in block:
        if not m:
            continue
        if key == m or key.startswith(m + " "):
            return True
        if len(m.split()) == 1 and len(m) >= 5:
            if re.fullmatch(
                rf"{re.escape(m)}( solar| energy| power| technologies| technology| networks)?",
                key,
            ):
                return True
    return False


@dataclass
class PurityVerdict:
    keep: bool
    reason: str
    label: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def classify_channel_purity(
    name: str,
    *,
    family: str = "general",
    specialty: str = "",
    core_categories: str = "",
    summary: str = "",
    distribution_type: str = "",
    website: str = "",
    extra_text: str = "",
) -> PurityVerdict:
    """Fast heuristic classify (no network / no LLM)."""
    family = (family or "general").strip().lower() or "general"
    name = (name or "").strip()
    if not name:
        return PurityVerdict(False, "empty_name", "drop_unclear")

    if _is_blocked_oem(name, family):
        return PurityVerdict(False, "known_oem_blocklist", "drop_oem")

    if NAME_CHANNEL_RE.search(name):
        return PurityVerdict(True, "name_channel_marker", "keep")

    dist = (distribution_type or "").strip()
    dist_evidence = ""
    if dist and dist.lower() not in {"distributor", "not publicly disclosed", "n/a", "na", ""}:
        dist_evidence = dist

    blob = _scrub_boilerplate(
        "\n".join(
            [
                specialty or "",
                core_categories or "",
                summary or "",
                dist_evidence,
                extra_text or "",
                website or "",
            ]
        )
    )
    ch = len(CHANNEL_RE.findall(blob))
    oem = len(OEM_RE.findall(blob))
    util = len(UTILITY_RE.findall(blob))
    mode = channel_purity_mode()

    if oem >= 2 and ch == 0:
        return PurityVerdict(False, f"oem_signals={oem}", "drop_oem")
    if util >= 2 and ch == 0:
        return PurityVerdict(False, f"utility_signals={util}", "drop_utility")
    if oem >= 1 and ch == 0 and mode == "strict":
        return PurityVerdict(False, f"oem_only={oem}", "drop_oem")
    if ch >= 1 and ch >= oem:
        return PurityVerdict(True, f"channel_signals={ch} oem={oem}", "keep")
    if CHANNEL_RE.search(specialty or "") and oem <= 1:
        return PurityVerdict(True, "specialty_channel", "keep")
    if mode == "strict":
        if oem >= 1:
            return PurityVerdict(False, f"strict_oem={oem}", "drop_oem")
        if util >= 1:
            return PurityVerdict(False, f"strict_utility={util}", "drop_utility")
        return PurityVerdict(False, "strict_unclear", "drop_unclear")
    if oem >= 1 and ch == 0:
        return PurityVerdict(False, f"oem_leaning={oem}", "drop_oem")
    if util >= 1 and ch == 0:
        return PurityVerdict(False, f"utility_leaning={util}", "drop_utility")
    return PurityVerdict(True, "balanced_keep_coverage", "keep")


def _row_fields(row: dict[str, Any]) -> dict[str, str]:
    name = str(
        row.get("Company")
        or row.get("company")
        or row.get("brand")
        or row.get("name")
        or ""
    ).strip()
    crawl = row.get("_crawl") if isinstance(row.get("_crawl"), dict) else {}
    return {
        "name": name,
        "specialty": str(
            row.get("Specialty Focus")
            or crawl.get("specialty_focus")
            or row.get("specialty_focus")
            or ""
        ),
        "core": str(
            row.get("Core Categories")
            or crawl.get("core_categories")
            or row.get("key_products")
            or ""
        ),
        "summary": str(
            row.get("Summary")
            or row.get("company_summary")
            or row.get("polished_summary")
            or row.get("snippet")
            or row.get("role_description")
            or ""
        ),
        "distribution": str(
            row.get("Distribution Type")
            or row.get("Business Type")
            or row.get("role")
            or ""
        ),
        "website": str(row.get("Website") or row.get("website") or row.get("domain") or ""),
    }


def _resolve_family(family: str, query: str) -> str:
    family = (family or "general").strip().lower() or "general"
    q = (query or "").lower()
    if family == "general" and any(
        t in q for t in ("renewable", "solar", "wind energy", "photovoltaic", "energy storage")
    ):
        return "renewable"
    return family


def purity_role_hint(family: str) -> str:
    return CHANNEL_ROLE_HINTS.get((family or "general").lower(), CHANNEL_ROLE_HINTS["general"])


def _openai_client():
    """OpenAI-compatible client (DeepSeek when OPENAI_BASE_URL points there)."""
    from openai import OpenAI

    try:
        from vendor_intel.config import Settings

        settings = Settings.load()
        key = (settings.openai_api_key or os.getenv("OPENAI_API_KEY") or "").strip()
        base = (settings.openai_base_url or os.getenv("OPENAI_BASE_URL") or "").strip().rstrip("/")
        model = (settings.openai_model or os.getenv("OPENAI_MODEL") or "").strip()
    except Exception:
        key = (os.getenv("OPENAI_API_KEY") or "").strip()
        base = (os.getenv("OPENAI_BASE_URL") or "").strip().rstrip("/")
        model = (os.getenv("OPENAI_MODEL") or "").strip()

    if not key:
        return None, None
    if not base:
        base = "https://api.openai.com/v1"
    if not model:
        model = "deepseek-chat" if "deepseek" in base.lower() else "gpt-4o-mini"
    return OpenAI(api_key=key, base_url=base), model


def _parse_llm_json(text: str) -> Any:
    raw = (text or "").strip()
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.I)
    if m:
        raw = m.group(1).strip()
    start = raw.find("{")
    start_a = raw.find("[")
    if start_a != -1 and (start == -1 or start_a < start):
        raw = raw[start_a:]
    elif start != -1:
        raw = raw[start:]
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {}


def llm_classify_channel_batch(
    items: list[dict[str, str]],
    *,
    family: str,
    query: str,
) -> dict[str, PurityVerdict]:
    """Classify a batch via DeepSeek/OpenAI. Keyed by normalized company name."""
    if not items:
        return {}
    client, model = _openai_client()
    if client is None or not model:
        print("  [channel-purity] LLM skipped: OPENAI_API_KEY not set", flush=True)
        return {}

    roles = purity_role_hint(family)
    payload = [
        {
            "Company": it.get("name") or "",
            "Website": (it.get("website") or "")[:80],
            "Specialty Focus": (it.get("specialty") or "")[:120],
            "Core Categories": (it.get("core") or "")[:120],
            "Summary": _scrub_boilerplate(it.get("summary") or "")[:180],
        }
        for it in items
    ]
    system = (
        "You are a strict channel-partner analyst. "
        f"Market: {query or family}. Valid KEEP roles: {roles}. "
        "KEEP only if the company's PRIMARY business is distribution / wholesale / "
        "dealer / reseller / VAD / importer / system integrator / authorized channel partner. "
        "DROP (keep=false) if PRIMARY business is OEM/manufacturer of the core product, "
        "utility/IPP/project developer, pure brand owner, media, or unrelated. "
        "If a manufacturer also has a small channel arm, still DROP (OEM-primary). "
        "Authorized dealers and wholesalers KEEP even if they install. "
        "Return ONLY JSON: "
        '{"results":[{"Company":"...","keep":true|false,'
        '"label":"keep|drop_oem|drop_utility|drop_unclear|drop_unrelated","reason":"short"}]}'
    )
    user = (
        f"Classify each company for channel-partner purity.\n"
        f"Companies:\n{json.dumps(payload, ensure_ascii=False)}"
    )
    content = ""
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            response_format={"type": "json_object"},
            max_tokens=3500,
        )
        content = (resp.choices[0].message.content or "").strip()
    except Exception:
        try:
            resp = client.chat.completions.create(
                model=model,
                temperature=0,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user + "\nReturn JSON only."},
                ],
                max_tokens=3500,
            )
            content = (resp.choices[0].message.content or "").strip()
        except Exception as exc2:  # noqa: BLE001
            print(
                f"  [channel-purity] LLM error: {type(exc2).__name__}: {exc2}",
                flush=True,
            )
            return {}

    data = _parse_llm_json(content)
    results = data.get("results") if isinstance(data, dict) else data
    out: dict[str, PurityVerdict] = {}
    if not isinstance(results, list):
        return out
    for r in results:
        if not isinstance(r, dict):
            continue
        name = str(r.get("Company") or r.get("name") or "").strip()
        if not name:
            continue
        keep = r.get("keep")
        if keep is None:
            continue
        keep_b = bool(keep) if not isinstance(keep, str) else keep.strip().lower() in (
            "true",
            "1",
            "yes",
            "keep",
        )
        label = str(r.get("label") or ("keep" if keep_b else "drop_unclear")).strip()
        if keep_b:
            label = "keep"
        elif label == "keep":
            label = "drop_unclear"
        reason = str(r.get("reason") or "llm").strip() or "llm"
        out[_norm(name)] = PurityVerdict(keep_b, f"llm:{reason}", label)
    return out


def _is_hard_heuristic_drop(v: PurityVerdict) -> bool:
    if v.keep:
        return False
    if v.reason == "known_oem_blocklist":
        return True
    prefixes = (
        "oem_signals",
        "utility_signals",
        "oem_only",
        "oem_leaning",
        "utility_leaning",
        "strict_oem",
        "strict_utility",
        "strict_unclear",
    )
    return any(v.reason.startswith(p) for p in prefixes)


def _llm_filter_rows(
    rows: list[dict[str, Any]],
    *,
    family: str,
    query: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Heuristic pre-drop known OEMs, then DeepSeek classify the rest."""
    family = _resolve_family(family, query)
    hard_keep: list[dict[str, Any]] = []
    hard_drop: list[dict[str, Any]] = []
    to_llm: list[tuple[dict[str, Any], dict[str, str]]] = []

    for row in rows:
        fields = _row_fields(row)
        v = classify_channel_purity(
            fields["name"],
            family=family,
            specialty=fields["specialty"],
            core_categories=fields["core"],
            summary=fields["summary"],
            distribution_type=fields["distribution"],
            website=fields["website"],
        )
        if _is_hard_heuristic_drop(v):
            hard_drop.append(
                {"company": fields["name"], "verdict": v.label, "reason": v.reason}
            )
            continue
        if v.keep and v.reason == "name_channel_marker" and channel_purity_mode() == "balanced":
            hard_keep.append(row)
            continue
        to_llm.append((row, fields))

    if not channel_purity_llm_enabled() or not to_llm:
        for row, fields in to_llm:
            v = classify_channel_purity(
                fields["name"],
                family=family,
                specialty=fields["specialty"],
                core_categories=fields["core"],
                summary=fields["summary"],
                distribution_type=fields["distribution"],
                website=fields["website"],
            )
            if v.keep:
                hard_keep.append(row)
            else:
                hard_drop.append(
                    {"company": fields["name"], "verdict": v.label, "reason": v.reason}
                )
        return hard_keep, hard_drop

    client, model = _openai_client()
    provider = "DeepSeek" if model and "deepseek" in model.lower() else "OpenAI"
    if client is None:
        print("  [channel-purity] LLM off (no API key) — heuristic only", flush=True)
        for row, fields in to_llm:
            v = classify_channel_purity(
                fields["name"],
                family=family,
                specialty=fields["specialty"],
                core_categories=fields["core"],
                summary=fields["summary"],
                distribution_type=fields["distribution"],
                website=fields["website"],
            )
            if v.keep:
                hard_keep.append(row)
            else:
                hard_drop.append(
                    {"company": fields["name"], "verdict": v.label, "reason": v.reason}
                )
        return hard_keep, hard_drop

    print(
        f"  [channel-purity] {provider} ({model}) classifying {len(to_llm)} companies "
        f"(chunk={_llm_chunk_size()})…",
        flush=True,
    )
    chunk = _llm_chunk_size()
    llm_map: dict[str, PurityVerdict] = {}
    total_batches = (len(to_llm) + chunk - 1) // chunk
    for i in range(0, len(to_llm), chunk):
        batch = to_llm[i : i + chunk]
        items = [fields for _, fields in batch]
        print(
            f"  [channel-purity] LLM batch {i // chunk + 1}/{total_batches} "
            f"({len(items)} companies)",
            flush=True,
        )
        part = llm_classify_channel_batch(items, family=family, query=query)
        llm_map.update(part)

    kept = list(hard_keep)
    dropped = list(hard_drop)
    for row, fields in to_llm:
        key = _norm(fields["name"])
        v = llm_map.get(key)
        if v is None:
            v = classify_channel_purity(
                fields["name"],
                family=family,
                specialty=fields["specialty"],
                core_categories=fields["core"],
                summary=fields["summary"],
                distribution_type=fields["distribution"],
                website=fields["website"],
            )
            v = PurityVerdict(v.keep, f"heuristic_fallback:{v.reason}", v.label)
        if v.keep:
            kept.append(row)
        else:
            dropped.append(
                {"company": fields["name"], "verdict": v.label, "reason": v.reason}
            )
    return kept, dropped


def filter_channel_pure_rows(
    rows: list[dict[str, Any]],
    *,
    family: str = "general",
    query: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Filter landscape/pipeline rows (heuristic + DeepSeek when enabled)."""
    if not channel_purity_enabled():
        return list(rows), []
    return _llm_filter_rows(rows, family=family, query=query)


async def filter_channel_pure_rows_async(
    rows: list[dict[str, Any]],
    *,
    family: str = "general",
    query: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Async wrapper: runs LLM purity in a worker thread, then optional web pass."""
    import asyncio

    if not channel_purity_enabled():
        return list(rows), []

    loop = asyncio.get_running_loop()
    with ThreadPoolExecutor(max_workers=1) as pool:
        kept, dropped = await loop.run_in_executor(
            pool,
            lambda: _llm_filter_rows(rows, family=family, query=query),
        )

    if not channel_purity_web_enabled() or not kept:
        return kept, dropped

    try:
        from vendor_intel.clients.searxng import searxng_search_any, searxng_urls
    except Exception:
        return kept, dropped

    urls = searxng_urls((os.getenv("SEARXNG_BASE_URL") or "http://127.0.0.1:8080").strip())
    if not urls:
        return kept, dropped

    sem = asyncio.Semaphore(8)
    final_kept: list[dict[str, Any]] = []
    extra_drop: list[dict[str, Any]] = []

    async def check(row: dict[str, Any]) -> tuple[dict[str, Any], PurityVerdict | None]:
        fields = _row_fields(row)
        async with sem:
            q = (
                f'"{fields["name"]}" (manufacturer OR OEM OR distributor OR wholesaler '
                f"OR dealer OR utility)"
            )
            try:
                hits, _ = await searxng_search_any(q, urls, max_results=8)
            except Exception:
                return row, None
            blob = "\n".join(f"{h.title} {h.snippet}" for h in (hits or []))
            oem = len(OEM_RE.findall(blob))
            ch = len(CHANNEL_RE.findall(blob))
            util = len(UTILITY_RE.findall(blob))
            if oem >= 2 and ch == 0:
                return row, PurityVerdict(False, f"web_oem={oem}", "drop_oem")
            if util >= 2 and ch == 0:
                return row, PurityVerdict(False, f"web_utility={util}", "drop_utility")
            return row, None

    results = await asyncio.gather(*[check(r) for r in kept])
    for row, v in results:
        if v is not None and not v.keep:
            fields = _row_fields(row)
            extra_drop.append(
                {"company": fields["name"], "verdict": v.label, "reason": v.reason}
            )
        else:
            final_kept.append(row)
    return final_kept, dropped + extra_drop
