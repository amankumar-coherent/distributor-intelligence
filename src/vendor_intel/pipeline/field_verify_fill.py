"""Per-field verify + correct pipeline stage (Step 6/7 in chatgpt_expand.py).

Generalizes four scripts proven out manually on the Global Flexible Packaging
Market dataset into reusable, market-agnostic functions:
  - Core Categories / Specialty Focus / Key Brands Represented verify+correct,
    plus channel-partner keep/remove, using live evidence + DeepSeek. When the
    fixed-query evidence check is genuinely inconclusive, DeepSeek gets a
    bounded (max 2 rounds) agentic search tool — backed by the same Google AI
    scraper — to decide its own follow-up queries instead of us guessing
    another fixed one. Optional sharpening only: never required, never the
    only thing standing between a manufacturer and the export file.
  - Online/Offline/Both (or "Retail / E-commerce / Both") classification via a
    pure heuristic site-scan + evidence-text classifier — needs no LLM, so it
    still runs at zero DeepSeek balance.
  - Contact Person / Role / LinkedIn via Apollo's FREE people search, then a
    two-phase web-search fallback that cross-validates any name it proposes
    against the actual LinkedIn URL slug before writing anything.
  - Office No. / Office Email straight off the company's own Contact Us page
    (tries common paths — /contact, /contact-us, /about/contact, etc. — and
    reads tel:/mailto: links plus visible phone/email text). No LLM or search
    engine involved; runs unconditionally like the channel-mode classifier.
  - A light geography enrichment pass so Continent / Geography (physical
    footprint: HQ/offices/warehouses/plants/stores) and Operational Presence
    (commercial reach: countries sold to/served, a superset of physical
    footprint) each end up with every location found for THAT specific
    concept, not just one and not copied from each other (the structural
    truncation itself is fixed in landscape_slim.py / gap_fill/gaps.py — this
    pass adds targeted follow-up queries when a row still looks single-country
    after Step 5).

Every evidence source degrades gracefully: DeepSeek out of balance, DDGS down,
Google AI scraper down/CAPTCHA-blocked, or Apollo rate-limited all fall back
or skip rather than crashing the pipeline. Nothing is ever invented — a field
only changes when real evidence supports it.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any
from urllib.parse import urljoin, urlparse

import httpx

from vendor_intel.clients.ddgs_engines import normalize_text_backends
from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping
from vendor_intel.config import Settings
from vendor_intel.pipeline.openai_expand_checkpoint import ExpandCheckpoint

NPD = "Not publicly disclosed"
NPD_SET = frozenset(
    {"", "n/a", "na", "-", "—", "not publicly disclosed", "unknown", "tbd", "none", "null"}
)
CHANNEL_MODE_ALIASES = ("Online / Offline / Both", "Retail / E-commerce / Both", "Retail \\ E-commerce \\ Both")

BACKEND_PARAM, BACKEND_NAMES, _DROPPED = normalize_text_backends()
GAI_BASE_URL = (os.getenv("GAI_BASE_URL") or os.getenv("GOOGLE_AI_SCRAPER_URL") or "http://127.0.0.1:15551").rstrip("/")

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

BATCH_CONCURRENCY = 6


def _log(msg: str) -> None:
    print(f"    [field-verify] {msg}", flush=True)


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in NPD_SET


CANONICAL_CHANNEL_MODE_KEY = CHANNEL_MODE_ALIASES[0]


def _channel_mode_key(row: dict[str, Any]) -> str | None:
    """Find whichever alias is present, for READING an existing value only
    (older checkpoints/rows from before a column rename may still carry a
    legacy key name). Writes always go to CANONICAL_CHANNEL_MODE_KEY — see
    classify_channel_modes — so a value found under a legacy key gets
    migrated forward instead of silently dropped at export time."""
    for k in CHANNEL_MODE_ALIASES:
        if k in row:
            return k
    return None


# --------------------------------------------------------------------------
# Evidence gathering — Google AI scraper primary, DDGS opportunistic fallback
# --------------------------------------------------------------------------

_GAI_BAD_MARKERS = ("unusual traffic", "did not match", "ai overview is not available", "no ai overview")


def _host_of(website: str) -> str:
    return re.sub(r"^https?://", "", (website or "").strip(), flags=re.I).split("/")[0]


async def gai_ask_safe(query: str, *, llm_clean: bool = False) -> str:
    try:
        data = await google_ai_ask(query, base_url=GAI_BASE_URL, llm_clean=llm_clean)
    except Exception as exc:  # noqa: BLE001
        _log(f"gai error: {exc}")
        return ""
    md = str(data.get("markdown") or "").strip()
    if not md or any(m in md.lower() for m in _GAI_BAD_MARKERS):
        return ""
    return md


def _ddgs_raw(query: str, max_results: int = 6) -> list[dict]:
    """Same query ddg_search() runs, but RAISES instead of swallowing —
    ddg_search_safe()'s retry loop needs the real exception to decide
    whether a rate-limit retry is worth it."""
    from ddgs import DDGS

    with DDGS() as d:
        return list(d.text(query, backend=BACKEND_PARAM, max_results=max_results)) or []


def ddg_search(query: str, max_results: int = 6) -> list[dict]:
    try:
        return _ddgs_raw(query, max_results)
    except Exception as exc:  # noqa: BLE001
        _log(f"ddgs error ({query[:60]}...): {exc}")
        return []


_DDGS_SEM = asyncio.Semaphore(2)
_DDGS_RATE_LIMIT_MARKERS = (
    "ratelimit", "rate limit", "202", "403", "429",
    "too many requests", "blocked", "timeout",
)


_search_router: Any = None


def _get_search_router() -> Any:
    """Lazily build one shared FreeSearchRouter (DDGS + Bing HTML + SearXNG,
    whichever backends are actually reachable, with automatic fallback
    between them) instead of hammering DDGS alone. Returns None if it can't
    be constructed, in which case callers fall back to plain DDGS."""
    global _search_router
    if _search_router is None:
        try:
            from vendor_intel.clients.search_router import FreeSearchRouter

            _search_router = FreeSearchRouter(Settings())
        except Exception as exc:  # noqa: BLE001
            _log(f"FreeSearchRouter unavailable, falling back to raw DDGS: {exc}")
            _search_router = False
    return _search_router or None


async def web_search_safe(query: str, max_results: int = 6) -> list[dict]:
    """Multi-backend free search instead of DDGS alone. DDGS getting rate-
    limited even after throttling (it's a single source under one shared
    IP, so throttling only slows the failure, it doesn't prevent it) — this
    routes through the same FreeSearchRouter the rest of the pipeline
    already uses for discovery, which automatically falls back to Bing HTML
    or a local SearXNG instance when DDGS itself is unavailable. Falls back
    to plain throttled DDGS if the router can't be built at all.
    """
    router = _get_search_router()
    if router is None:
        return await ddg_search_safe(query, max_results)
    try:
        results = await router.search(query, discovery_mode=False)
    except Exception as exc:  # noqa: BLE001
        _log(f"search router error ({query[:60]}...): {exc}")
        return await ddg_search_safe(query, max_results)
    if not results:
        return await ddg_search_safe(query, max_results)
    return [{"title": r.title, "body": r.snippet, "href": r.link} for r in results[:max_results]]


async def ddg_search_safe(query: str, max_results: int = 6, *, retries: int = 2) -> list[dict]:
    """Throttled, retrying DDGS call.

    DuckDuckGo rate-limits by source IP, and this pipeline was firing up to
    BATCH_CONCURRENCY (6) simultaneous DDGS requests from the same machine —
    exactly the burst pattern that triggers it, which is what was showing up
    as frequent "ddgs error" log lines. This caps DDGS concurrency
    independently at 2 (regardless of how wide the rest of the pipeline
    fans out) and retries once with a short backoff specifically on
    rate-limit-looking errors before giving up. Still never raises and
    still returns [] on final failure — DDGS stays opportunistic-only, same
    as before; this only makes it more likely to actually succeed instead
    of getting rate-limited by its own concurrency.
    """
    async with _DDGS_SEM:
        for attempt in range(retries + 1):
            try:
                return await asyncio.to_thread(_ddgs_raw, query, max_results)
            except Exception as exc:  # noqa: BLE001
                msg = str(exc).lower()
                if attempt < retries and any(m in msg for m in _DDGS_RATE_LIMIT_MARKERS):
                    _log(f"ddgs rate-limited, retrying in {3 * (attempt + 1)}s ({query[:50]}...)")
                    await asyncio.sleep(3.0 * (attempt + 1))
                    continue
                _log(f"ddgs error ({query[:60]}...): {exc}")
                return []
    return []


async def gather_channel_fit_evidence(company: str, website: str, family: str) -> str:
    """Neutral evidence for the manufacturer-vs-distributor DECISION only.

    Deliberately does NOT use a leading query like "distributor categories
    brands" — confirmed by direct inspection that for a genuine OEM (ABB,
    Siemens, ...), that phrasing makes Google's AI Overview describe the
    company's OWN downstream distributor network ("ABB distributors supply
    ...") instead of the company itself, which reads exactly like real
    channel-partner evidence and causes a false keep. Asking a neutral,
    either/or question avoids presupposing the answer.
    """
    host = _host_of(website)
    q = (
        f"{company} {host} is it primarily a manufacturer/OEM that designs or "
        f"produces its own {family} products, or a distributor/wholesaler/reseller/"
        f"channel partner that sells other companies' {family} products"
    )
    return await gai_ask_safe(q)


async def gather_company_evidence(company: str, website: str, family: str) -> str:
    """Evidence for Core Categories / Specialty Focus / Key Brands text content
    ONLY — not for the manufacturer-vs-distributor decision itself, see
    gather_channel_fit_evidence() for why that needs a separate, neutral ask.
    """
    host = _host_of(website)
    q = f"{company} {host} {family} distributor categories brands represented specialty focus"
    gai = await gai_ask_safe(q)
    if gai and len(gai) >= 150:
        return gai
    hits = await web_search_safe(f'"{company}" {family} distributor brands categories', 6)
    if not hits:
        return gai
    lines = [f"- {h.get('title', '')}: {str(h.get('body', ''))[:280]} ({h.get('href', '')})" for h in hits]
    return (gai + "\n\n" + "\n".join(lines)).strip()


# --------------------------------------------------------------------------
# Section 1: Core Categories / Specialty Focus / Key Brands + keep/remove
# --------------------------------------------------------------------------

SYSTEM_FIELD_VERIFY = """You audit ONE row of a {family} market channel-partner spreadsheet.

A valid row is a genuine DISTRIBUTOR / RESELLER / CHANNEL PARTNER / IMPORTER
for the "{query}" market — NOT a manufacturer-only OEM, NOT an unrelated
business.

Using ONLY the evidence given (real web search / AI overview text), do the
following for THIS company:

1. "core_categories": verify or correct so it follows this structure:
   [Main Product Category] + [Sub Categories] — broad product groups this
   distributor handles, main/broadest category first then narrower
   sub-categories, semicolon-separated: "Category 1; Category 2; Category 3"
   (3-8 items). Keep the existing value if evidence supports it; correct it
   if evidence clearly contradicts it. Never invent categories unsupported
   by evidence or the existing value.
2. "specialty_focus": verify or correct so it follows this EXACT structure,
   in order (<= 28 words total):
   [Product Category] + [Specific Application / Customer Segment] + [Value-added
   expertise (optional)]
   e.g. "Industrial rupture discs for oil & gas pressure-relief systems, with
   in-house engineering and testing support." Always include the first two
   parts from real evidence; append the third clause ONLY when the evidence
   names a genuine value-add (custom engineering, technical support,
   certification, stocking program, etc.) — omit it entirely rather than
   inventing one. No company-name restatement, no trailing country/region.
3. "key_brands": verify or correct — ONLY manufacturer/brand names this
   company actually distributes, represents, or is an authorized reseller of,
   semicolon-separated: "Brand A; Brand B; Brand C". Distributor partners
   often represent multiple manufacturers — list every brand evidenced, not
   just one. Use a brand name ONLY if explicitly evidenced. Never invent. If
   the only brands found are from a clearly unrelated industry, do not
   include them. If nothing is evidenced, return "Not publicly disclosed".
4. "keep": true only if this company is genuinely a channel partner for this
   market AND (when any key brands exist) those brands are actually related
   to this market. Set keep=false with a short "reason" if: it's a
   manufacturer-only OEM, a pure logistics/3PL company, an unrelated
   business, or its only evidenced brands belong to a different industry.

Return STRICT JSON only:
{{"results":[{{"company":"exact name","core_categories":"...","specialty_focus":"...",
"key_brands":"...","keep":true,"reason":"short"}}]}}"""


def _ds_client_and_model(settings: Settings):
    from vendor_intel.pipeline.chatgpt_expand import _client, _model

    return _client(settings), _model(settings)


def resolve_deepseek(deepseek_ok: bool) -> tuple[Any, str | None]:
    """Resolve one DeepSeek client for the whole Step 6 run — every sub-stage
    below takes it as an OPTIONAL parameter and only uses it to refine an
    answer it already has from live evidence; none of them require it."""
    if not deepseek_ok:
        return None, None
    try:
        settings = Settings()
        client, model = _ds_client_and_model(settings)
        return client, model
    except Exception as exc:  # noqa: BLE001
        _log(f"DeepSeek unavailable, evidence-only for this run: {exc}")
        return None, None


async def _ds_json(client: Any, model: str, *, system: str, user: str, max_tokens: int = 500) -> dict:
    """Small shared helper for the optional one-shot DeepSeek calls below.
    Returns {} on any failure (including insufficient balance) — callers
    already have an evidence-based answer and just keep it."""
    try:
        resp = await asyncio.to_thread(
            client.chat.completions.create,
            model=model,
            temperature=0,
            max_tokens=max_tokens,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        return _parse_json(resp.choices[0].message.content or "{}")
    except Exception as exc:  # noqa: BLE001
        msg = str(exc).lower()
        if "insufficient balance" in msg or "402" in msg:
            _log("DeepSeek insufficient balance — keeping evidence-only answer")
        else:
            _log(f"DeepSeek call failed, keeping evidence-only answer: {exc}")
        return {}


_CHANNEL_FIT_SEARCH_TOOL = [
    {
        "type": "function",
        "function": {
            "name": "search",
            "description": (
                "Search the live web for evidence about a company. Returns a short "
                "text summary. Use a specific, targeted query — not a generic one."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "A specific web search query"},
                },
                "required": ["query"],
            },
        },
    }
]


async def _run_channel_fit_search_tool(tool_call: Any) -> str:
    try:
        args = json.loads(tool_call.function.arguments or "{}")
    except Exception:
        args = {}
    query = str(args.get("query") or "").strip()
    if not query:
        return "No query given."
    result = await gai_ask_safe(query)
    return result or "No results found for that query."


async def ds_agentic_channel_fit_check(
    company: str,
    website: str,
    family: str,
    ds_client: Any,
    ds_model: str,
    *,
    max_rounds: int = 2,
) -> tuple[bool | None, str]:
    """Bounded agentic escalation for the manufacturer-vs-distributor decision.

    Used ONLY when the fixed-query evidence-first check (classify_manufacturer_
    signal on gather_channel_fit_evidence) comes back genuinely inconclusive
    AND DeepSeek is reachable — never a requirement, never the primary path.
    DeepSeek gets a `search` tool (backed by the same Google AI scraper we
    already use elsewhere) and decides its own follow-up queries instead of
    us hand-guessing another fixed one — e.g. noticing evidence describes a
    company's own downstream distributor network and re-querying more
    specifically, which is exactly the failure mode fixed queries kept
    running into during testing. Capped at max_rounds search calls so a
    stubborn case can't loop forever or blow up cost/latency.
    """
    host = _host_of(website)
    market = family or "this market"
    system = (
        "You are researching ONE company to decide if it is primarily a "
        "MANUFACTURER/OEM (designs/produces its own products) or a "
        "DISTRIBUTOR/WHOLESALER/RESELLER/CHANNEL PARTNER (sells other "
        f"companies' products) for the {market} market. Use the search tool "
        "to find real evidence — specific, targeted queries (e.g. company "
        "name plus \"is a manufacturer\" or plus \"official distributor\"). "
        "Search again with a sharper query if the first result is unclear or "
        "turns out to be about the company's OWN downstream distributor "
        "network rather than the company itself. Stop searching once you "
        "have a confident answer."
    )
    messages: list[dict[str, Any]] = [
        {"role": "system", "content": system},
        {"role": "user", "content": f"Company: {company}\nWebsite/host: {host}"},
    ]

    for round_i in range(max_rounds):
        try:
            resp = await asyncio.to_thread(
                ds_client.chat.completions.create,
                model=ds_model,
                temperature=0,
                max_tokens=400,
                tools=_CHANNEL_FIT_SEARCH_TOOL,
                tool_choice="auto",
                messages=messages,
            )
        except Exception as exc:  # noqa: BLE001
            msg = str(exc).lower()
            if "insufficient balance" in msg or "402" in msg:
                _log("DeepSeek insufficient balance — skipping agentic search")
            else:
                _log(f"agentic search error (round {round_i + 1}): {exc}")
            return None, "agentic_search_unavailable"

        msg_obj = resp.choices[0].message
        tool_calls = getattr(msg_obj, "tool_calls", None)
        if not tool_calls:
            break  # model is done searching — fall through to the final answer

        messages.append(
            {
                "role": "assistant",
                "content": msg_obj.content or "",
                "tool_calls": [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {"name": tc.function.name, "arguments": tc.function.arguments},
                    }
                    for tc in tool_calls
                ],
            }
        )
        for tc in tool_calls:
            result_text = await _run_channel_fit_search_tool(tc)
            messages.append({"role": "tool", "tool_call_id": tc.id, "content": result_text[:3000]})

    messages.append(
        {
            "role": "user",
            "content": (
                "Based on everything found, answer now — no more searching. "
                'Return STRICT JSON only: {"role":"manufacturer"|"distributor"|'
                '"unclear","reason":"short"}'
            ),
        }
    )
    try:
        resp = await asyncio.to_thread(
            ds_client.chat.completions.create,
            model=ds_model,
            temperature=0,
            max_tokens=150,
            response_format={"type": "json_object"},
            messages=messages,
        )
        data = _parse_json(resp.choices[0].message.content or "{}")
    except Exception as exc:  # noqa: BLE001
        _log(f"agentic search final-answer error: {exc}")
        return None, "agentic_search_unavailable"

    role = str(data.get("role") or "").strip().lower()
    reason = str(data.get("reason") or "agentic_search").strip()
    if role == "manufacturer":
        return False, f"agentic_search:{reason}"
    if role == "distributor":
        return True, f"agentic_search:{reason}"
    return None, "agentic_search_unclear"


def _parse_json(text: str) -> dict:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I | re.S).strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return {}
    return {}


async def deepseek_verify_batch(
    client: Any, model: str, family: str, query: str, items: list[dict[str, str]]
) -> list[dict[str, Any]]:
    system = SYSTEM_FIELD_VERIFY.format(family=family, query=query)
    payload = {
        "companies": [
            {
                "company": it["company"],
                "existing_core": it.get("core", ""),
                "existing_specialty": it.get("specialty", ""),
                "existing_brands": it.get("brands", ""),
                "evidence": it.get("evidence", "")[:3500],
            }
            for it in items
        ]
    }
    for attempt in range(3):
        try:
            resp = await asyncio.to_thread(
                client.chat.completions.create,
                model=model,
                temperature=0,
                max_tokens=2200,
                response_format={"type": "json_object"},
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
                ],
            )
            data = _parse_json(resp.choices[0].message.content or "{}")
            results = data.get("results")
            if isinstance(results, list):
                return results
            return []
        except Exception as exc:  # noqa: BLE001
            msg = str(exc).lower()
            if "insufficient balance" in msg or "402" in msg:
                _log("DeepSeek insufficient balance — skipping brand/category verify for remaining rows")
                return []
            _log(f"deepseek error (attempt {attempt + 1}): {exc}")
            await asyncio.sleep(2.0 * (attempt + 1))
    return []


async def verify_brands_categories(
    rows: list[dict[str, Any]],
    *,
    family: str,
    query: str,
    deepseek_ok: bool = False,
    ds_client: Any = None,
    ds_model: str | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Returns (kept_rows, rejected_rows). Mutates rows in place for kept ones.

    The manufacturer-vs-channel-partner keep/reject decision never depends
    solely on DeepSeek: every row gets a fresh evidence-gather (Google AI /
    DDGS) and a regex-based manufacturer/distributor classification
    (channel_purity.classify_manufacturer_signal) on that live evidence. This
    alone decides keep/reject when DeepSeek is unavailable or out of balance
    — a company is only ever dropped on a positive OEM-signal-over-channel-
    signal reading of real evidence, never by default. DeepSeek, when
    reachable, additionally verifies/upgrades Core Categories, Specialty
    Focus, and Key Brands, and can add its own keep=false veto on top — either
    signal saying "manufacturer" is enough to drop a row.
    """
    from vendor_intel.intelligence.channel_purity import classify_manufacturer_signal

    if not rows:
        return rows, []

    client, model = ds_client, ds_model
    if client is None and deepseek_ok:
        client, model = resolve_deepseek(True)

    kept: list[dict[str, Any]] = []
    rejected: list[dict[str, Any]] = []
    deepseek_alive = client is not None

    for start in range(0, len(rows), 8):
        batch = rows[start : start + 8]
        evidences, fit_evidences = await asyncio.gather(
            asyncio.gather(
                *(
                    gather_company_evidence(
                        str(r.get("Company") or ""), str(r.get("Website") or ""), family
                    )
                    for r in batch
                )
            ),
            asyncio.gather(
                *(
                    gather_channel_fit_evidence(
                        str(r.get("Company") or ""), str(r.get("Website") or ""), family
                    )
                    for r in batch
                )
            ),
        )

        # DeepSeek-independent floor: pure regex over the NEUTRAL evidence
        # (not the categories/brands evidence — that phrasing describes a
        # manufacturer's own downstream distributors, not the company itself).
        # Genuinely inconclusive cases get one bounded agentic escalation
        # (DeepSeek decides its own follow-up searches) when DeepSeek is
        # reachable — this is optional sharpening, not a requirement: when
        # DeepSeek is down or the escalation also comes back unclear, the
        # row still defaults to keep rather than being dropped on a guess.
        async def _resolve_base(r: dict[str, Any], fit_ev: str) -> tuple[str, tuple[bool, str]]:
            name = str(r.get("Company") or "")
            keep2, reason2 = classify_manufacturer_signal(fit_ev, name)
            if keep2 is None and deepseek_alive:
                keep3, reason3 = await ds_agentic_channel_fit_check(
                    name, str(r.get("Website") or ""), family, client, model
                )
                if keep3 is not None:
                    return name, (keep3, reason3)
            return name, (keep2 if keep2 is not None else True, reason2)

        base_pairs = await asyncio.gather(
            *(_resolve_base(r, fit_ev) for r, fit_ev in zip(batch, fit_evidences))
        )
        base_verdict: dict[str, tuple[bool, str]] = dict(base_pairs)

        results: list[dict[str, Any]] = []
        if deepseek_alive:
            items = [
                {
                    "company": str(r.get("Company") or ""),
                    "core": str(r.get("Core Categories") or ""),
                    "specialty": str(r.get("Specialty Focus") or ""),
                    "brands": str(r.get("Key Brands Represented") or ""),
                    "evidence": ev,
                }
                for r, ev in zip(batch, evidences)
            ]
            results = await deepseek_verify_batch(client, model, family, query, items)
            if not results and start == 0:
                # Treat an empty first batch as balance/availability failure — stop trying.
                deepseek_alive = False

        by_name = {str(d.get("company") or "").strip().lower(): d for d in results}
        for r in batch:
            name = str(r.get("Company") or "")
            name_key = name.strip().lower()
            d = by_name.get(name_key)
            base_ok, base_reason = base_verdict.get(name, (True, "evidence_inconclusive"))
            if d:
                new_core = str(d.get("core_categories") or "").strip()
                new_spec = str(d.get("specialty_focus") or "").strip()
                new_brands = str(d.get("key_brands") or "").strip()
                if new_core and len(new_core) > 3:
                    r["Core Categories"] = new_core
                if new_spec and len(new_spec) >= 8:
                    r["Specialty Focus"] = new_spec
                if new_brands:
                    r["Key Brands Represented"] = new_brands
                llm_ok = bool(d.get("keep", True))
                final_ok = llm_ok and base_ok
                reason = base_reason if not base_ok else str(d.get("reason", "") or "llm")
            else:
                final_ok = base_ok
                reason = base_reason
            if final_ok:
                kept.append(r)
            else:
                rejected.append({"company": r.get("Company"), "reason": reason, "stage": "field_verify"})
        if start % 40 == 0:
            _log(f"brand/category verify {min(start + 8, len(rows))}/{len(rows)}")

    return kept, rejected


# --------------------------------------------------------------------------
# Section 2: Online/Offline/Both — pure heuristic, no LLM required
# --------------------------------------------------------------------------

ECOMMERCE_SIGNALS = [
    (re.compile(r"add[-_ ]?to[-_ ]?(cart|bag|basket)", re.I), "add_to_cart_button"),
    (re.compile(r"shopping[-_ ]?cart", re.I), "shopping_cart_widget"),
    (re.compile(r"/checkout|checkout[-_ ]?(page|process|flow)", re.I), "checkout_flow"),
    (re.compile(r"woocommerce", re.I), "woocommerce_platform"),
    (re.compile(r"cdn\.shopify|shopify\.com|/cart\.js", re.I), "shopify_platform"),
    (re.compile(r"bigcommerce", re.I), "bigcommerce_platform"),
    (re.compile(r"magento", re.I), "magento_platform"),
    (re.compile(r"buy[-_ ]?now", re.I), "buy_now_button"),
    (re.compile(r"your[-_ ]?cart|cart[-_ ]?(count|icon|total)", re.I), "cart_widget"),
    (re.compile(r"online[-_ ]?store|shop[-_ ]?online|order[-_ ]?online", re.I), "online_store_mention"),
]
OFFLINE_SIGNALS = [
    (re.compile(r"request\s+a\s+quote", re.I), "quote_request_only"),
    (re.compile(r"find\s+a\s+(branch|location|distributor)", re.I), "branch_locator"),
    (re.compile(r"call\s+(us\s+)?to\s+order", re.I), "call_to_order"),
    (re.compile(r"contact\s+(us|sales)\s+to\s+(order|purchase|buy)", re.I), "contact_to_order"),
    (re.compile(r"open\s+(a|an)\s+(trade|wholesale|business|credit)\s+account", re.I), "trade_account"),
    (re.compile(r"(our\s+)?sales\s+team|account\s+manager|sales\s+rep(resentative)?s?\b", re.I), "sales_rep_mention"),
    (re.compile(r"our\s+locations|branch\s+network|distributor\s+network|find\s+your\s+local", re.I), "location_network"),
    (re.compile(r'href=["\']tel:', re.I), "tel_link_present"),
    (re.compile(r"(call|phone|tel)[^.\n]{0,45}(order|sales|quote|wholesale|distributor)", re.I), "phone_near_sales_context"),
    (re.compile(r"physical\s+(branch(es)?|store(s)?|location(s)?)|brick[-\s]and[-\s]mortar", re.I), "physical_presence"),
    (re.compile(r"trade\s+centre(s)?|trade\s+center(s)?|showroom(s)?", re.I), "trade_centre_showroom"),
]
ONLINE_TEXT_RE = re.compile(
    r"\b(shop\s+online|order(s|ing)?\s+online|online\s+store|add\s+to\s+cart|buy\s+now|"
    r"e-?commerce\s+(checkout|store)|purchase(s|d)?\s+online|online\s+ordering|checkout\s+(page|process))\b",
    re.I,
)
OFFLINE_TEXT_RE = re.compile(
    r"\b(does\s+not\s+(offer|have|provide)\s+(an\s+)?online\s+(store|ordering|checkout)|"
    r"no\s+online\s+(store|ordering|checkout)|call\s+(us\s+)?to\s+order|"
    r"contact\s+(us|sales|a\s+rep)\s+to\s+(order|purchase|buy)|"
    r"(our\s+)?sales\s+team|account\s+manager|distributor\s+network|"
    r"physical\s+(branch(es)?|store(s)?)|brick[-\s]and[-\s]mortar|trade\s+centre|showroom)\b",
    re.I,
)


def _strip_html(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html or "")
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def scan_website_channel_signals(website: str) -> tuple[list[str], list[str], int]:
    url = (website or "").strip()
    if not url:
        return [], [], 0
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    paths = ["", "/contact", "/shop", "/contact-us", "/about", "/store", "/wholesale", "/cart"]
    ecom: set[str] = set()
    off: set[str] = set()
    fetched = 0
    try:
        with httpx.Client(
            headers={"User-Agent": UA}, verify=False, timeout=httpx.Timeout(10.0, connect=4.0), follow_redirects=True
        ) as client:
            for path in paths:
                try:
                    u = url if not path else urljoin(url.rstrip("/") + "/", path.lstrip("/"))
                    r = client.get(u)
                    if r.status_code >= 400:
                        continue
                    raw = r.text
                    for pattern, label in ECOMMERCE_SIGNALS:
                        if pattern.search(raw):
                            ecom.add(label)
                    text = _strip_html(raw)
                    for pattern, label in OFFLINE_SIGNALS:
                        if pattern.search(text):
                            off.add(label)
                    fetched += 1
                    if fetched >= 6:
                        break
                except Exception:
                    continue
    except Exception:
        return [], [], fetched
    return sorted(ecom), sorted(off), fetched


_TEL_HREF_RE = re.compile(r'href=["\']tel:([^"\']+)["\']', re.I)
_MAILTO_HREF_RE = re.compile(r'href=["\']mailto:([^"\'?]+)', re.I)
_EMAIL_TEXT_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
_OFFICE_EMAIL_PREFIX_RE = re.compile(
    r"^(info|contact|office|sales|enquir|inquir|hello|support|admin|mail|marketing|customerservice|customer)@",
    re.I,
)
_JUNK_EMAIL_RE = re.compile(
    r"@(example\.com|sentry\.io|wixpress\.com|schema\.org|localhost)$|"
    r"\.(png|jpg|jpeg|gif|svg|webp)(@|$)",
    re.I,
)
# Contact-us style pages first — the whole point is these facts are published
# there directly, which is far more reliable than inferring them from a
# generic web search.
CONTACT_PAGE_PATHS = (
    "/contact",
    "/contact-us",
    "/contactus",
    "/contact-us.html",
    "/about/contact",
    "/support/contact",
    "/about-us/contact-us",
    "",
)


def _looks_like_office_email(addr: str) -> bool:
    a = (addr or "").strip().strip(".").lower()
    if not a or "@" not in a or _JUNK_EMAIL_RE.search(a):
        return False
    local, _, domain = a.partition("@")
    return bool(local) and "." in domain


async def scan_website_contact_details(website: str) -> dict[str, str]:
    """Crawl the company's own Contact Us / About page(s) for a general office
    phone number and office email. This is the "look for a Contact Us page on
    the website" strategy — usually far more reliable than a generic web
    search, since companies publish this info directly on their own site."""
    url = (website or "").strip()
    if not url:
        return {}
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url

    phones: list[str] = []
    emails: list[str] = []

    def _fetch() -> None:
        try:
            with httpx.Client(
                headers={"User-Agent": UA},
                verify=False,
                timeout=httpx.Timeout(10.0, connect=4.0),
                follow_redirects=True,
            ) as client:
                for path in CONTACT_PAGE_PATHS:
                    try:
                        u = url if not path else urljoin(url.rstrip("/") + "/", path.lstrip("/"))
                        r = client.get(u)
                        if r.status_code >= 400:
                            continue
                        raw = r.text
                        for m in _TEL_HREF_RE.findall(raw):
                            cleaned = re.sub(r"[^\d+]", "", m)
                            if len(cleaned) >= 7:
                                phones.append(m.strip())
                        for m in _MAILTO_HREF_RE.findall(raw):
                            if _looks_like_office_email(m):
                                emails.append(m.strip())
                        text = _strip_html(raw)
                        for m in _EMAIL_TEXT_RE.findall(text):
                            if _looks_like_office_email(m):
                                emails.append(m.strip())
                        if phones and emails:
                            break
                    except Exception:
                        continue
        except Exception:
            pass

    await asyncio.to_thread(_fetch)

    phone_cands = list(dict.fromkeys(phones))
    email_cands = list(dict.fromkeys(e.lower() for e in emails))
    out: dict[str, Any] = {"_phone_candidates": phone_cands, "_email_candidates": email_cands}
    if phone_cands:
        out["Office No."] = phone_cands[0]
    if email_cands:
        # Prefer an obviously-generic office inbox (info@/contact@/sales@…)
        # over any other address the page happens to mention.
        generic = [e for e in email_cands if _OFFICE_EMAIL_PREFIX_RE.match(e)]
        out["Office Email"] = (generic or email_cands)[0]
    return out


async def ds_pick_office_contact(
    company: str, phone_cands: list[str], email_cands: list[str], ds_client: Any, ds_model: str
) -> dict[str, str]:
    """Disambiguate among MULTIPLE real candidates the site scan already
    found (e.g. several regional office numbers, or several email addresses
    with no obviously-generic prefix). Never invents a new value — only
    picks from the given lists, or returns empty if none clearly fits."""
    data = await _ds_json(
        ds_client,
        ds_model,
        system=(
            "Pick the single best GENERAL/MAIN office phone number and general "
            "office email (info@/contact@/sales@ style inbox, not a personal "
            "address) from the candidate lists given — choose ONLY from the exact "
            "strings provided, never invent a new one. Return STRICT JSON only: "
            '{"phone":"<one of the given phones or empty>","email":"<one of the '
            'given emails or empty>"}'
        ),
        user=json.dumps({"company": company, "phone_candidates": phone_cands, "email_candidates": email_cands}),
        max_tokens=150,
    )
    phone = str(data.get("phone") or "").strip()
    email = str(data.get("email") or "").strip().lower()
    out: dict[str, str] = {}
    if phone in phone_cands:
        out["Office No."] = phone
    if email in email_cands:
        out["Office Email"] = email
    return out


async def fill_office_contact_details(
    rows: list[dict[str, Any]], *, ds_client: Any = None, ds_model: str | None = None
) -> None:
    """Mutates rows in place. Fills Office No. / Office Email straight off the
    company's own Contact Us page — no LLM dependency for the base pass; an
    optional DeepSeek disambiguation runs only when several distinct
    candidates were found and none is an obviously-generic inbox."""

    def _needs(r: dict[str, Any]) -> bool:
        return emptyish(r.get("Office No.")) or emptyish(r.get("Office Email"))

    targets = [r for r in rows if _needs(r) and str(r.get("Website") or "").strip()]
    if not targets:
        return

    async def _one(r: dict[str, Any]) -> None:
        found = await scan_website_contact_details(str(r.get("Website") or ""))
        phone_cands = found.get("_phone_candidates") or []
        email_cands = found.get("_email_candidates") or []
        has_generic_email = any(_OFFICE_EMAIL_PREFIX_RE.match(e) for e in email_cands)
        ambiguous = ds_client is not None and (
            (len(phone_cands) > 1) or (len(email_cands) > 1 and not has_generic_email)
        )
        if ambiguous:
            picked = await ds_pick_office_contact(
                str(r.get("Company") or ""), phone_cands, email_cands, ds_client, ds_model
            )
            if picked.get("Office No."):
                found["Office No."] = picked["Office No."]
            if picked.get("Office Email"):
                found["Office Email"] = picked["Office Email"]
        if found.get("Office No.") and emptyish(r.get("Office No.")):
            r["Office No."] = found["Office No."]
        if found.get("Office Email") and emptyish(r.get("Office Email")):
            r["Office Email"] = found["Office Email"]

    for start in range(0, len(targets), BATCH_CONCURRENCY):
        batch = targets[start : start + BATCH_CONCURRENCY]
        await asyncio.gather(*(_one(r) for r in batch))
        if start % 40 == 0:
            _log(f"office contact-page scan {min(start + BATCH_CONCURRENCY, len(targets))}/{len(targets)}")


def classify_channel_mode(existing: str, ecom: list[str], off: list[str], site_fetched: int, text_evidence: str) -> str:
    online_from_site, offline_from_site = bool(ecom), bool(off)
    online_signal = online_from_site or bool(ONLINE_TEXT_RE.search(text_evidence))
    offline_signal = offline_from_site or bool(OFFLINE_TEXT_RE.search(text_evidence))
    valid = {"Online", "Offline", "Both"}

    if online_signal and offline_signal:
        return "Both"
    site_blocked = site_fetched == 0
    if site_blocked and existing in valid and not (online_from_site or offline_from_site):
        if online_signal or offline_signal:
            return existing  # weak text-only signal, site unreachable — keep verified value
    if online_signal:
        return "Online"
    if offline_signal:
        return "Offline"
    return existing if existing in valid else NPD


async def classify_channel_modes(
    rows: list[dict[str, Any]], *, ds_client: Any = None, ds_model: str | None = None
) -> None:
    """Mutates rows in place. No LLM dependency — always runs on every row.
    When ds_client is provided, rows the heuristic still can't resolve (NPD)
    get one optional DeepSeek read of the same evidence — never overrides a
    confident heuristic answer."""
    if not rows:
        return

    def _scan(r: dict[str, Any]) -> tuple[int, tuple[list[str], list[str], int]]:
        return id(r), scan_website_channel_signals(str(r.get("Website") or ""))

    scan_cache: dict[int, tuple[list[str], list[str], int]] = {}
    with ThreadPoolExecutor(max_workers=12) as ex:
        futs = [ex.submit(_scan, r) for r in rows]
        for fut in as_completed(futs):
            rid, result = fut.result()
            scan_cache[rid] = result

    gai_ok = False
    try:
        gai_ok = await google_ai_ping(base_url=GAI_BASE_URL)
    except Exception:
        gai_ok = False

    async def _classify_one(r: dict[str, Any]) -> None:
        legacy_key = _channel_mode_key(r)
        existing = str(r.get(legacy_key) or "").strip() if legacy_key else ""
        ecom, off, fetched = scan_cache.get(id(r), ([], [], 0))
        text_evidence = ""
        if gai_ok:
            host = _host_of(str(r.get("Website") or ""))
            company = str(r.get("Company") or "")
            q = f"{company} {host} does it sell online through its own website or only offline via phone quote sales rep"
            text_evidence = await gai_ask_safe(q)
        new_val = classify_channel_mode(existing, ecom, off, fetched, text_evidence)
        if new_val == NPD and ds_client is not None:
            evidence = "\n".join(
                [
                    f"site e-commerce signals found: {', '.join(ecom) or 'none'}",
                    f"site offline signals found: {', '.join(off) or 'none'}",
                    f"pages fetched from site: {fetched}",
                    f"web evidence text: {text_evidence[:1500]}",
                ]
            )
            data = await _ds_json(
                ds_client,
                ds_model,
                system=(
                    "Classify a distributor's sales channel from the evidence given. "
                    "Return STRICT JSON only: {\"value\":\"Online\"|\"Offline\"|\"Both\"|"
                    "\"Not publicly disclosed\"}. Online = customers can complete a "
                    "purchase on the company's own website. Offline = sales happen via "
                    "phone/email/rep/branch only. Both = real evidence of both. If the "
                    "evidence does not clearly show either, return \"Not publicly "
                    "disclosed\" — never guess."
                ),
                user=f"Company: {r.get('Company')}\nEvidence:\n{evidence}",
                max_tokens=80,
            )
            ds_val = str(data.get("value") or "").strip()
            if ds_val in {"Online", "Offline", "Both"}:
                new_val = ds_val
        if new_val:
            # Always write the canonical key so a value read from a legacy
            # alias (older checkpoint) still survives to export.
            r[CANONICAL_CHANNEL_MODE_KEY] = new_val
            if legacy_key and legacy_key != CANONICAL_CHANNEL_MODE_KEY:
                r[legacy_key] = new_val

    for start in range(0, len(rows), BATCH_CONCURRENCY):
        batch = rows[start : start + BATCH_CONCURRENCY]
        await asyncio.gather(*(_classify_one(r) for r in batch))
        if start % 40 == 0:
            _log(f"channel-mode classify {min(start + BATCH_CONCURRENCY, len(rows))}/{len(rows)}")


# --------------------------------------------------------------------------
# Section 3: Contact Person / Role / LinkedIn — Apollo free search + web fallback
# --------------------------------------------------------------------------

TITLE_WORDS = (
    "VP Sales", "Vice President of Sales", "Head of Sales", "Sales Director",
    "Director of Sales", "Commercial Director", "Chief Revenue Officer", "CRO",
    "Head of Business Development", "VP Business Development", "Channel Sales",
    "Head of Procurement", "VP Procurement", "Vice President of Procurement",
    "Procurement Director", "Director of Procurement", "Chief Procurement Officer",
    "Procurement Manager", "Purchasing Manager", "Head of Purchasing",
)
ROLE_OK = re.compile(
    r"(head of sales|sales head|\bvp\b.{0,25}sales|vice president.{0,25}sales|svp.{0,25}sales|"
    r"sales director|director of sales|director.{0,20}sales|"
    r"head of procurement|procurement head|\bvp\b.{0,25}procurement|vice president.{0,25}procurement|"
    r"procurement director|director of procurement|director.{0,20}procurement|"
    r"head of purchasing|purchasing director|chief (?:revenue|commercial|procurement)|\bcro\b|\bcpo\b|"
    r"commercial director|head of (?:business development|channel|partner)|national sales)",
    re.I,
)
ROLE_BAD = re.compile(r"\b(ceo|cfo|cto|coo|ciso|founder|chairman|owner|hr|marketing|accountant)\b", re.I)
BAD_NAME = re.compile(
    r"^(ceo|cfo|cto|coo|ciso|president|founder|leadership|board|team|contact|sales|procurement|manager|director|vp|head)\b",
    re.I,
)
LI_PERSON_RE = re.compile(r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/in/[A-Za-z0-9\-_%]+/?", re.I)
_NON_NAME_WORDS = {
    "at", "in", "on", "for", "by", "with", "from", "to", "as", "is", "are", "was", "were",
    "a", "an", "this", "that", "their", "its", "his", "her",
    "vice", "president", "director", "head", "sales", "procurement", "commercial",
    "revenue", "chief", "business", "development", "channel", "purchasing",
    "manager", "officer", "national", "regional", "svp", "vp", "cro", "cpo",
    "partner", "team", "department", "division",
}
_COMPANY_STOPWORDS = {
    "the", "and", "of", "co", "inc", "ltd", "llc", "gmbh", "group", "corp", "corporation",
    "company", "distribution", "distributors", "international", "global", "products", "industries",
}


def _role_is_target(role: str) -> bool:
    r = (role or "").strip()
    if not r:
        return False
    if ROLE_BAD.search(r) and not ROLE_OK.search(r):
        return False
    return bool(ROLE_OK.search(r))


def _name_ok(name: str) -> bool:
    s = re.sub(r"\s+", " ", (name or "").strip())
    if not s or BAD_NAME.match(s):
        return False
    first = s.split()[0]
    return len(re.sub(r"[^A-Za-z]", "", first)) >= 2


def _tokens(s: str) -> set[str]:
    return {t.lower() for t in re.findall(r"[A-Za-z']+", s or "") if len(t) > 1}


def _has_non_name_word(name: str) -> bool:
    return bool(_tokens(name) & _NON_NAME_WORDS)


def _is_company_fragment(name: str, company: str) -> bool:
    if _has_non_name_word(name):
        return True
    name_tokens = _tokens(name) - _COMPANY_STOPWORDS
    if not name_tokens:
        return True
    company_tokens = _tokens(company)
    if name_tokens.issubset(company_tokens):
        return True
    distinctive = {t for t in company_tokens if len(t) >= 5} - _COMPANY_STOPWORDS
    return bool(name_tokens & distinctive)


def name_from_linkedin_slug(url: str) -> str:
    m = re.search(r"/in/([A-Za-z0-9\-]+)", url)
    if not m:
        return ""
    parts = [p for p in m.group(1).split("-") if p and not re.fullmatch(r"[0-9a-f]{6,}", p, re.I) and not p.isdigit()]
    if len(parts) < 2:
        return ""
    return " ".join(p.capitalize() for p in parts[:3])


def _names_closely_match(a: str, b: str) -> bool:
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return False
    a_first = a.strip().split()[0].lower()
    if a_first not in tb:
        return False
    return len(ta & tb) >= min(len(ta), len(tb))


def _needs_contact(cp: str, role: str, li: str) -> bool:
    if emptyish(cp) or BAD_NAME.match(cp):
        return True
    if emptyish(role) or not _role_is_target(role):
        return True
    return "linkedin.com/in/" not in li.lower()


async def apollo_find_contact(company: str, website: str) -> dict[str, str] | None:
    try:
        from vendor_intel.enrichment.apollo_contacts import ApolloContactClient
    except Exception:
        return None
    try:
        async with ApolloContactClient(verbose=False, request_delay=2.0, timeout=30.0, max_retries=2) as apollo:
            hit = await apollo.find_contact({"company": company, "website": website}, enrich_match=False)
    except Exception:
        return None
    if not hit or hit.get("error") or not hit.get("contact_person"):
        return None
    name = str(hit.get("contact_person") or "").strip()
    role = str(hit.get("contact_role") or "").strip()
    if not _name_ok(name) or not _role_is_target(role):
        return None
    return {"contact_person": name, "contact_role": role, "linkedin": "", "source": "apollo"}


async def gai_find_name_role(company: str, website: str) -> tuple[str, str] | None:
    host = _host_of(website)
    q = (
        f"{company} {host} current VP of Sales OR Head of Sales OR Sales Director OR "
        f"Commercial Director OR VP of Procurement OR Procurement Director name"
    )
    md = await gai_ask_safe(q)
    if not md:
        return None
    role_found, role_idx = "", -1
    for tw in TITLE_WORDS:
        idx = md.lower().find(tw.lower())
        if idx != -1:
            role_found, role_idx = tw, idx
            break
    if not role_found or not _role_is_target(role_found):
        return None
    window = md[max(0, role_idx - 150) : role_idx + 50]
    for raw_cand in re.findall(r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z.\-]+){1,2})\b", window):
        cand = re.sub(r"\.\s.*$", "", raw_cand)
        cand = re.sub(r"\s+", " ", cand).strip().rstrip(".")
        if len(cand.split()) < 2 or not _name_ok(cand) or BAD_NAME.match(cand):
            continue
        if _is_company_fragment(cand, company):
            continue
        return cand, role_found
    return None


async def find_linkedin_url(name: str, company: str) -> tuple[str, str]:
    hits = await web_search_safe(f'"{name}" "{company}" site:linkedin.com/in', 5)
    for h in hits:
        m = LI_PERSON_RE.search(str(h.get("href") or ""))
        if not m:
            continue
        slug_name = name_from_linkedin_slug(m.group(0))
        if slug_name and _names_closely_match(name, slug_name):
            return m.group(0).rstrip("/"), slug_name
    md = await gai_ask_safe(f"{name} {company} LinkedIn profile URL")
    m = LI_PERSON_RE.search(md)
    if m:
        slug_name = name_from_linkedin_slug(m.group(0))
        if slug_name and _names_closely_match(name, slug_name):
            return m.group(0).rstrip("/"), slug_name
    return "", ""


async def ds_find_name_role(
    company: str, website: str, ds_client: Any, ds_model: str
) -> tuple[str, str] | None:
    """Fallback for when the regex-based gai_find_name_role() can't parse a
    name out of the evidence text. Reads the SAME evidence with DeepSeek
    instead — still subject to the LinkedIn-slug cross-validation before
    anything is written, same as the regex path."""
    host = _host_of(website)
    q = (
        f"{company} {host} current VP of Sales OR Head of Sales OR Sales Director OR "
        f"Commercial Director OR VP of Procurement OR Procurement Director name"
    )
    md = await gai_ask_safe(q)
    if not md:
        return None
    data = await _ds_json(
        ds_client,
        ds_model,
        system=(
            "Find the named sales or procurement leader for this company in the "
            "evidence text. Valid roles only: VP/Head/Director of Sales, Commercial "
            "Director, CRO, VP/Head/Director of Procurement or Purchasing. Never "
            "CEO/founder/HR/marketing. Return STRICT JSON only: "
            '{"name":"...","role":"..."} — both empty strings if no such named '
            "person is explicitly in the text. Never invent a name."
        ),
        user=f"Company: {company}\nEvidence:\n{md[:2500]}",
        max_tokens=120,
    )
    name = str(data.get("name") or "").strip()
    role = str(data.get("role") or "").strip()
    if not name or not role or not _name_ok(name) or not _role_is_target(role):
        return None
    if _is_company_fragment(name, company):
        return None
    return name, role


async def web_search_find_contact(
    company: str, website: str, *, ds_client: Any = None, ds_model: str | None = None
) -> dict[str, str] | None:
    found = await gai_find_name_role(company, website)
    if not found and ds_client is not None:
        found = await ds_find_name_role(company, website, ds_client, ds_model)
    if not found:
        return None
    name, role = found
    li_url, slug_name = await find_linkedin_url(name, company)
    if not li_url:
        return None  # no confirmed LinkedIn match — too risky to write
    final_name = slug_name if len(slug_name.split()) > len(name.split()) else name
    return {"contact_person": final_name, "contact_role": role, "linkedin": li_url, "source": "web"}


async def fill_contacts(
    rows: list[dict[str, Any]], *, use_apollo: bool, ds_client: Any = None, ds_model: str | None = None
) -> None:
    """Mutates rows in place."""
    targets = [
        r
        for r in rows
        if _needs_contact(
            str(r.get("Contact Person") or ""), str(r.get("Role") or ""), str(r.get("LinkedIn") or "")
        )
    ]
    if not targets:
        return

    apollo_misses: list[dict[str, Any]] = []
    if use_apollo:
        apollo_rate_limited = False
        for r in targets:
            if apollo_rate_limited:
                apollo_misses.append(r)
                continue
            company = str(r.get("Company") or "")
            website = str(r.get("Website") or "")
            hit = await apollo_find_contact(company, website)
            if hit:
                r["Contact Person"] = hit["contact_person"]
                r["Role"] = hit["contact_role"]
            else:
                apollo_misses.append(r)
    else:
        apollo_misses = targets

    gai_ok = False
    try:
        gai_ok = await google_ai_ping(base_url=GAI_BASE_URL)
    except Exception:
        gai_ok = False
    if not gai_ok or not apollo_misses:
        return

    async def _web_one(r: dict[str, Any]) -> None:
        hit = await web_search_find_contact(
            str(r.get("Company") or ""), str(r.get("Website") or ""),
            ds_client=ds_client, ds_model=ds_model,
        )
        if hit:
            r["Contact Person"] = hit["contact_person"]
            r["Role"] = hit["contact_role"]
            r["LinkedIn"] = hit["linkedin"]

    for start in range(0, len(apollo_misses), BATCH_CONCURRENCY):
        batch = apollo_misses[start : start + BATCH_CONCURRENCY]
        await asyncio.gather(*(_web_one(r) for r in batch))
        if start % 40 == 0:
            _log(f"contact web-search fallback {min(start + BATCH_CONCURRENCY, len(apollo_misses))}/{len(apollo_misses)}")


# --------------------------------------------------------------------------
# Section 4: Geography enrichment — targeted regional follow-up queries
# --------------------------------------------------------------------------

REGION_PROBES = ("Europe", "Asia", "North America", "Middle East", "Africa", "Latin America")


async def ds_decompose_regions(text: str, ds_client: Any, ds_model: str) -> list[str]:
    """Evidence often names a region GROUP ("the Nordics", "Benelux", "GCC",
    "DACH", "ASEAN") instead of individual countries — parse_country_list()
    alone can't expand those. Ask DeepSeek to decompose ONLY group terms
    actually present in this text, then still run the result back through
    parse_country_list() so nothing not on the real country list survives —
    DeepSeek can misread the text, but it can't invent a country that isn't
    in that canonical list."""
    from vendor_intel.export.landscape_slim import parse_country_list

    if not text:
        return []
    data = await _ds_json(
        ds_client,
        ds_model,
        system=(
            "The text may mention region GROUPS (Nordics, Benelux, DACH, GCC, "
            "ASEAN, Baltics, Balkans, Iberia, Maghreb, etc.) instead of naming "
            "individual countries. List the individual countries implied ONLY by "
            "group terms that actually appear in this text — do not add anything "
            "else, do not repeat countries already named individually. Return "
            'STRICT JSON only: {"countries":["...", "..."]} — empty list if no '
            "region group is mentioned."
        ),
        user=text[:2000],
        max_tokens=200,
    )
    countries = data.get("countries")
    if not isinstance(countries, list):
        return []
    return parse_country_list("; ".join(str(c) for c in countries))


async def _regional_probe_one(r: dict[str, Any], *, ds_client: Any = None, ds_model: str | None = None) -> None:
    """Second-tier fallback for a row still single-country after the general
    footprint/reach queries — evidence sometimes only gives an aggregate count
    ("operates in 60+ countries") without naming them. Probing region by
    region often surfaces the actual country names."""
    from vendor_intel.enrichment.gap_fill.gaps import merge_location_text
    from vendor_intel.export.landscape_slim import parse_country_list

    company = str(r.get("Company") or "")
    host = _host_of(str(r.get("Website") or ""))

    async def _probe(region: str) -> tuple[list[str], list[str]]:
        geo_md = await gai_ask_safe(f"{company} {host} offices, plants, or warehouses in {region}")
        reach_md = await gai_ask_safe(f"{company} {host} sells to or serves customers in {region}")
        geo_list = parse_country_list(geo_md) if geo_md else []
        reach_list = parse_country_list(reach_md) if reach_md else []
        if ds_client is not None:
            if geo_md:
                geo_list = list(dict.fromkeys(geo_list + await ds_decompose_regions(geo_md, ds_client, ds_model)))
            if reach_md:
                reach_list = list(dict.fromkeys(reach_list + await ds_decompose_regions(reach_md, ds_client, ds_model)))
        return geo_list, reach_list

    results = await asyncio.gather(*(_probe(region) for region in REGION_PROBES))
    geo_found: list[str] = []
    reach_found: list[str] = []
    for g, rc in results:
        for c in g:
            if c not in geo_found:
                geo_found.append(c)
        for c in rc:
            if c not in reach_found:
                reach_found.append(c)

    if geo_found:
        existing_geo = str(r.get("Continent / Geography") or "")
        merged = merge_location_text(existing_geo, "; ".join(geo_found))
        if merged != existing_geo:
            r["Continent / Geography"] = merged

    final_geo = str(r.get("Continent / Geography") or "")
    union_addition = merge_location_text("; ".join(reach_found), final_geo) if reach_found else final_geo
    if union_addition:
        existing_op = str(r.get("Operational Presence") or "")
        merged = merge_location_text(existing_op, union_addition)
        if merged != existing_op:
            r["Operational Presence"] = merged


async def enrich_geography(rows: list[dict[str, Any]], *, ds_client: Any = None, ds_model: str | None = None) -> None:
    """Continent / Geography = physical footprint (HQ, offices, warehouses,
    plants, stores). Operational Presence = commercial reach (countries the
    company sells to or serves) — always a superset of physical footprint,
    since anywhere with a physical presence is obviously also served.

    Asks two SEPARATE evidence questions and writes each to its own column;
    never copies one column's result into the other verbatim.
    """
    from vendor_intel.enrichment.gap_fill.gaps import merge_location_text
    from vendor_intel.export.landscape_slim import parse_country_list

    gai_ok = False
    try:
        gai_ok = await google_ai_ping(base_url=GAI_BASE_URL)
    except Exception:
        gai_ok = False
    if not gai_ok:
        return

    def _is_single_location(v: str) -> bool:
        return bool(v) and not emptyish(v) and ";" not in v and "," not in v

    targets = [
        r
        for r in rows
        if _is_single_location(str(r.get("Continent / Geography") or ""))
        or _is_single_location(str(r.get("Operational Presence") or ""))
    ]
    if not targets:
        return

    async def _enrich_one(r: dict[str, Any]) -> None:
        company = str(r.get("Company") or "")
        host = _host_of(str(r.get("Website") or ""))

        geo_q = (
            f"{company} {host} list every country where it has a physical presence: "
            f"offices, warehouses, manufacturing plants, or retail stores"
        )
        geo_md = await gai_ask_safe(geo_q)
        geo_found = parse_country_list(geo_md) if geo_md else []
        if geo_md and ds_client is not None:
            geo_found = list(dict.fromkeys(geo_found + await ds_decompose_regions(geo_md, ds_client, ds_model)))
        if geo_found:
            existing_geo = str(r.get("Continent / Geography") or "")
            addition = "; ".join(geo_found)
            if existing_geo and not emptyish(existing_geo):
                merged = merge_location_text(existing_geo, addition)
                if merged != existing_geo:
                    r["Continent / Geography"] = merged
            else:
                r["Continent / Geography"] = addition

        reach_q = (
            f"{company} {host} list every country or region it sells to, distributes "
            f"in, or serves customers — its commercial/market reach, which may be "
            f"broader than where it has offices"
        )
        reach_md = await gai_ask_safe(reach_q)
        reach_found = parse_country_list(reach_md) if reach_md else []
        if reach_md and ds_client is not None:
            reach_found = list(dict.fromkeys(reach_found + await ds_decompose_regions(reach_md, ds_client, ds_model)))

        # Operational Presence is a superset: reach evidence UNION the (now
        # possibly-updated) Continent / Geography countries — never the other
        # way around, and never just a copy of one field into the other.
        final_geo = str(r.get("Continent / Geography") or "")
        union_addition = "; ".join(reach_found) if reach_found else ""
        if final_geo and not emptyish(final_geo):
            union_addition = merge_location_text(union_addition, final_geo) if union_addition else final_geo
        if union_addition:
            existing_op = str(r.get("Operational Presence") or "")
            if existing_op and not emptyish(existing_op):
                merged = merge_location_text(existing_op, union_addition)
                if merged != existing_op:
                    r["Operational Presence"] = merged
            else:
                r["Operational Presence"] = union_addition

    for start in range(0, len(targets), BATCH_CONCURRENCY):
        batch = targets[start : start + BATCH_CONCURRENCY]
        await asyncio.gather(*(_enrich_one(r) for r in batch))

    # Tier 2: rows still single-country after the general queries — evidence
    # likely only gave an aggregate count without naming countries. Probe
    # region by region (fewer rows reach this point, so the extra query
    # volume stays bounded).
    still_stuck = [
        r
        for r in targets
        if _is_single_location(str(r.get("Continent / Geography") or ""))
        or _is_single_location(str(r.get("Operational Presence") or ""))
    ]
    for start in range(0, len(still_stuck), BATCH_CONCURRENCY):
        batch = still_stuck[start : start + BATCH_CONCURRENCY]
        await asyncio.gather(*(_regional_probe_one(r, ds_client=ds_client, ds_model=ds_model) for r in batch))


# --------------------------------------------------------------------------
# Orchestrator
# --------------------------------------------------------------------------


async def verify_and_correct_fields(
    rows: list[dict[str, str]],
    *,
    query: str,
    family: str,
    ckpt: ExpandCheckpoint | None = None,
    use_apollo: bool = True,
) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    """Returns (corrected_rows, rejected_rows). Never raises — every evidence
    source degrades gracefully so this step can't crash the pipeline."""
    if not rows:
        return rows, []

    deepseek_ok = bool((os.getenv("OPENAI_API_KEY") or "").strip())
    ds_client, ds_model = resolve_deepseek(deepseek_ok)
    ds_tag = "on" if ds_client is not None else "off"

    _log(f"brand/category verify + channel-fit keep/remove ({len(rows)} rows, deepseek={ds_tag})...")
    kept, rejected = await verify_brands_categories(
        rows, family=family, query=query, ds_client=ds_client, ds_model=ds_model
    )
    if ckpt:
        ckpt.bump("5d_field_verify", "5d1_brands", kept_after_brands=len(kept), rejected_so_far=len(rejected))

    _log(f"online/offline/both classify ({len(kept)} rows, deepseek={ds_tag})...")
    await classify_channel_modes(kept, ds_client=ds_client, ds_model=ds_model)
    if ckpt:
        ckpt.bump("5d_field_verify", "5d2_channel_mode")

    _log(f"contact person/role/linkedin ({len(kept)} rows, deepseek={ds_tag})...")
    await fill_contacts(kept, use_apollo=use_apollo, ds_client=ds_client, ds_model=ds_model)
    if ckpt:
        ckpt.bump("5d_field_verify", "5d3_contacts")

    _log(f"office phone/email from Contact Us pages ({len(kept)} rows, deepseek={ds_tag})...")
    await fill_office_contact_details(kept, ds_client=ds_client, ds_model=ds_model)
    if ckpt:
        ckpt.bump("5d_field_verify", "5d3b_office_contact")

    _log(f"geography enrichment pass ({len(kept)} rows, deepseek={ds_tag})...")
    await enrich_geography(kept, ds_client=ds_client, ds_model=ds_model)
    if ckpt:
        ckpt.bump("5d_field_verify", "5d4_geography")

    _log(f"done: kept={len(kept)} rejected={len(rejected)}")
    return kept, rejected
