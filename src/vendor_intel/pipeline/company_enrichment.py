"""Per-company, per-field progressive enrichment: Kompass -> ChatGPT scraper ->
Google AI Overview scraper -> DeepSeek -> Smart Crawl.

This is the ONE new orchestration layer this feature adds. Every actual data
source it calls already exists elsewhere in this repo and is reused as-is:

  Kompass          -> vendor_intel.discovery.kompass_distributors
  ChatGPT scraper  -> vendor_intel.scraping.chatgpt_scraper.ask_chatgpt
  Google AI        -> vendor_intel.enrichment.gap_fill.sources.fill_from_google_ai_scraper
  DeepSeek         -> vendor_intel.pipeline.chatgpt_expand._chat_json (OpenAI-compatible,
                       already routes to DeepSeek when OPENAI_BASE_URL points there)
  Smart Crawl      -> crawler.smart_crawl.smart_crawl(mode="company")

Design decisions (see conversation for why):
  - Runs entirely IN-MEMORY per company, before anything is written to the
    FINAL Excel — this is what avoids the provenance-loss bug where a
    Kompass row re-read from disk looks "weak" and gets fully re-filled
    (Data Sources is an Excel-internal-only column, never persisted).
  - One new is_missing()/merge_company_data() here, scoped to this module
    only. The three other missing-value checks already in this codebase
    (chatgpt_expand._EMPTYISH, search_stack_fill._NPD,
    enrichment.gap_fill.gaps.PLACEHOLDERS) are untouched — changing those
    risks behavior in unrelated code paths that already work.
  - Each stage only ever asks for the columns still missing after the prior
    stage — never the whole row — per the explicit performance requirement.
  - Field-source priority is tracked in-memory only (not written to the
    final Excel, since SLIM_COLUMNS has no source column and the schema is
    not being changed) via EnrichedCompany.field_sources, useful for logging
    and for tests.
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any

from vendor_intel.export.landscape_slim import SLIM_COLUMNS

# ---------------------------------------------------------------------------
# Missing-value detection — scoped to this orchestrator only.
# ---------------------------------------------------------------------------

_MISSING_SENTINELS = frozenset(
    {
        "",
        "n/a",
        "na",
        "none",
        "null",
        "not available",
        "not found",
        "unknown",
    }
)


def is_missing(value: Any) -> bool:
    """True when `value` should be treated as not-yet-found.

    Deliberately does NOT treat these as missing (per explicit instruction):
      - 0 / "0" / "0 employees" — a real, meaningful value
      - "Not publicly disclosed" — this codebase's existing convention for
        "we looked, there's genuinely nothing public" is NOT the same as
        "we haven't looked yet"; treating it as missing would make every
        enrichment stage re-search fields other stages already gave up on
        for a documented reason.
    """
    if value is None:
        return True
    s = str(value).strip()
    if not s:
        return True
    return s.lower() in _MISSING_SENTINELS


# Columns this orchestrator manages end-to-end. Matches SLIM_COLUMNS exactly
# except it excludes "Summary" (a derived free-text field, not something any
# single source is asked to "find") — Summary is left to existing pipeline
# logic (finalize_slim_row / _ensure_complete_row) untouched.
MANAGED_COLUMNS: tuple[str, ...] = tuple(c for c in SLIM_COLUMNS if c != "Summary")

# Fields Smart Crawl is authoritative for (spec's explicit priority override:
# "Smart Crawl should be considered the preferred source for website-derived
# fields... use that verified website information rather than an unsupported
# AI guess"). All other fields follow strict Kompass > ChatGPT > Google >
# DeepSeek > SmartCrawl order.
WEBSITE_DERIVED_COLUMNS: frozenset[str] = frozenset(
    {
        "Key Brands Represented",
        "Core Categories",
        "Specialty Focus",
        "Office No.",
        "Office Email",
    }
)

# Only these role types are acceptable for Contact Person/Role/LinkedIn —
# sales and procurement decision-makers relevant to vendor/distributor
# outreach, not just any employee a source happens to find. Checked against
# whatever "Role" text a source returns; a contact whose role doesn't match
# ANY of these patterns is rejected outright (name/role/LinkedIn all left
# empty) rather than kept with a mismatched or generic title.
ACCEPTED_CONTACT_ROLE_PATTERNS: tuple[str, ...] = (
    r"\bvp\b.*\bsales\b", r"\bsales\b.*\bvp\b", r"\bvice president\b.*\bsales\b",
    r"\bhead of sales\b", r"\bsales director\b", r"\bdirector of sales\b",
    r"\bsales manager\b", r"\bsales lead\b", r"\bchief sales officer\b", r"\bcso\b",
    r"\bhead of procurement\b", r"\bprocurement manager\b", r"\bprocurement director\b",
    r"\bprocurement lead\b", r"\bpurchasing manager\b", r"\bpurchasing director\b",
    r"\bchief procurement officer\b", r"\bcpo\b",
)


def is_accepted_contact_role(role: str) -> bool:
    """True only for sales/procurement leadership titles — see
    ACCEPTED_CONTACT_ROLE_PATTERNS. A blank/missing role is NOT accepted
    (nothing to validate against means the contact can't be confirmed as
    the right kind of person)."""
    import re

    r = (role or "").strip().lower()
    if not r:
        return False
    return any(re.search(pat, r) for pat in ACCEPTED_CONTACT_ROLE_PATTERNS)


_CONTACT_PERSON_COLUMNS: tuple[str, ...] = ("Contact Person", "Role", "LinkedIn")


def _enforce_contact_role_gate(new_data: dict[str, str]) -> dict[str, str]:
    """Applied to every source's `new_data` before it's merged (see
    merge_company_data) — if a contact-person field is present but the Role
    doesn't match an accepted sales/procurement title, strip ALL THREE
    contact fields from this update rather than merge a wrong-role contact.
    A source that returns a name/LinkedIn but no Role at all is also
    rejected — no role text means the type can't be confirmed."""
    has_any_contact_field = any(k in new_data for k in _CONTACT_PERSON_COLUMNS)
    if not has_any_contact_field:
        return new_data
    role = str(new_data.get("Role") or "")
    if is_accepted_contact_role(role):
        return new_data
    out = {k: v for k, v in new_data.items() if k not in _CONTACT_PERSON_COLUMNS}
    return out


# Long free-text fields worth a DeepSeek cleanup pass (spec: "polish the
# sentence and write in proper format"). Deliberately excludes short/
# structured fields (Website, Email, Phone, Founded, Employees, codes,
# Contact Person, Role) — those are either already correct as typed/scraped
# or risk being corrupted by a rewrite rather than improved.
POLISH_COLUMNS: tuple[str, ...] = (
    "Specialty Focus",
    "Core Categories",
    "Key Brands Represented",
    "Ownership",
    "Summary",
)

_SOURCE_ORDER = (
    "kompass", "chatgpt_scraper", "google_ai", "deepseek",
    "website_verify", "smart_crawl", "write_summary", "deepseek_polish",
)


def _log(msg: str) -> None:
    print(f"  [enrich] {msg}", flush=True)


def _detail(msg: str) -> None:
    print(f"    -> {msg}", flush=True)


@dataclass
class EnrichedCompany:
    row: dict[str, str]
    field_sources: dict[str, str] = field(default_factory=dict)
    stages_run: list[str] = field(default_factory=list)
    stages_skipped: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def missing_columns(row: dict[str, str], *, columns: tuple[str, ...] = MANAGED_COLUMNS) -> list[str]:
    return [c for c in columns if is_missing(row.get(c))]


def _union_geography_value(column: str, existing_value: str, new_value: str) -> str:
    """Combine two location answers for the same geography column into one
    complete list — reuses landscape_slim's own country parsing/formatting
    so the result matches the column's documented format exactly (see
    landscape_slim.py's module docstring: "Continent / Geography is
    formatted as REGION(Country[, Country...])."). Returns "" if nothing
    parseable came out of either value (caller then just keeps the existing
    value unchanged)."""
    from vendor_intel.export.landscape_slim import format_continent_geography, parse_country_list

    existing_countries = parse_country_list(existing_value)
    new_countries = parse_country_list(new_value)
    if not new_countries:
        return existing_value  # nothing new/parseable to add
    seen = {c.lower() for c in existing_countries}
    combined = list(existing_countries) + [c for c in new_countries if c.lower() not in seen]
    if len(combined) == len(existing_countries):
        return existing_value  # new value added nothing not already present

    if column == "Continent / Geography":
        return format_continent_geography(combined) or existing_value
    # Operational Presence — plain semicolon list, not grouped by region.
    return "; ".join(combined)


def merge_company_data(
    existing: dict[str, str],
    new_data: dict[str, str],
    *,
    source: str,
    field_sources: dict[str, str] | None = None,
    prefer_source_for: frozenset[str] = frozenset(),
) -> dict[str, str]:
    """Merge `new_data` onto `existing`, field by field.

    Rule 1/2: never overwrite a present value with a missing one; fill empty
    cells with a present new value.
    Rule 3/4: `prefer_source_for` lets a specific source (e.g. Smart Crawl)
    override an already-present value for a specific column set, because
    that source is documented as authoritative for those columns — this is
    the ONLY case a non-missing value is ever replaced. Never used for
    columns outside `prefer_source_for`.
    Rule 5: field_sources (if given) is updated in place to record which
    source last populated each column — internal bookkeeping only, never
    written to the exported Excel schema.
    Rule 6: Contact Person/Role/LinkedIn are only merged together and only
    when Role matches an accepted sales/procurement leadership title (see
    is_accepted_contact_role) — applied to every source, not just AI
    searches, so "Contact Person" always means the same thing regardless of
    which stage found it.
    Rule 7: Continent / Geography and Operational Presence are UNIONED
    across sources rather than fill-once-then-locked — every source may
    contribute additional countries it found that an earlier source missed,
    so the final value always lists every confirmed location, not just
    whichever source happened to answer first (per explicit instruction:
    write all locations where the company is present).
    """
    new_data = _enforce_contact_role_gate(new_data)
    out = dict(existing)
    for col, new_val in new_data.items():
        if col not in MANAGED_COLUMNS and col != "Summary":
            continue
        new_s = str(new_val or "").strip()
        if is_missing(new_s):
            continue  # never write a missing value over anything
        cur_s = str(out.get(col) or "").strip()
        if col in _GEOGRAPHY_COLUMNS and not is_missing(cur_s):
            unioned = _union_geography_value(col, cur_s, new_s)
            if unioned and unioned != cur_s:
                out[col] = unioned
                if field_sources is not None:
                    field_sources[col] = f"{field_sources.get(col, source)}+{source}"
            continue
        if is_missing(cur_s):
            out[col] = new_s
            if field_sources is not None:
                field_sources[col] = source
        elif col in prefer_source_for:
            out[col] = new_s
            if field_sources is not None:
                field_sources[col] = source
        # else: existing value is present and this source has no override
        # priority for this column — keep the existing value (Rule 4).
    return out


# ---------------------------------------------------------------------------
# Stage 1 — Kompass
# ---------------------------------------------------------------------------


def _kompass_row_for_company(company_name: str, market: str) -> dict[str, str] | None:
    """Look up a single company's Kompass profile by name, if a Kompass
    company URL can be found for it in the given market context. Reuses the
    existing verify_company() pipeline — no duplicate Kompass logic. Returns
    None (never raises) if Kompass can't be reached or doesn't have this
    company, so callers just fall through.

    Deliberately builds the slim row from the raw VerifiedDistributor fields
    rather than calling kd.to_slim_row() directly: to_slim_row() fills
    unfound fields with the literal string "Not publicly disclosed" as a
    *display* default for Kompass's own output, which is NOT the same thing
    as "Kompass genuinely searched and confirmed nothing exists" — treating
    that filler as a real value here would make every later enrichment
    stage think the field is already handled and skip it (e.g. Key Brands
    Represented always defaults to NOT_DISCLOSED when Kompass has no brand
    data, which must NOT block ChatGPT/Google/DeepSeek/SmartCrawl from
    actually searching for it)."""
    from vendor_intel.discovery import kompass_distributors as kd

    try:
        page_info = kd.find_market_product_page(market)
        if page_info is None:
            return None
        product_url, product_id = page_info
        candidates = kd.discover_distributor_candidates(
            product_url, market=market, max_candidates=kd.max_companies_per_market() or 500
        )
        target_key = company_name.strip().lower()
        for cand in candidates:
            # discover_distributor_candidates doesn't resolve names up front
            # (they come from the profile page itself) — verify_company()
            # fills cand.name from the page, so match after fetching.
            result = kd.verify_company(cand, market=market, kompass_product_id=product_id)
            if result is None:
                continue
            if result.company_name.strip().lower() == target_key:
                return _raw_kompass_fields(result)
        return None
    except Exception as exc:  # noqa: BLE001
        _detail(f"Kompass lookup error: {type(exc).__name__}: {exc}")
        return None


def _real_kompass_value(value: str) -> str:
    """Strip Kompass's own "Not publicly disclosed" display filler back to
    empty — see _kompass_row_for_company's docstring for why."""
    v = str(value or "").strip()
    return "" if v.lower() == "not publicly disclosed" else v


def _raw_kompass_fields(d: Any) -> dict[str, str]:
    """Map a VerifiedDistributor's REAL (non-filler) fields onto the SLIM
    column names this orchestrator manages, so merge_company_data() can
    treat genuinely-unfound Kompass fields as still missing."""
    return {
        "Company": _real_kompass_value(d.company_name),
        "Website": _real_kompass_value(d.website),
        "Founded": _real_kompass_value(d.founded_year),
        "Headquarters": _real_kompass_value(d.headquarters) or _real_kompass_value(d.country),
        "Employees": _real_kompass_value(d.employees),
        "Key Brands Represented": _real_kompass_value(d.brands),
        "Core Categories": _real_kompass_value(d.market),
        "Contact Person": _real_kompass_value(d.contact_person),
        "Email": _real_kompass_value(d.email),
        "Office No.": _real_kompass_value(d.phone),
    }


def enrich_from_kompass(row: dict[str, str], *, market: str) -> tuple[dict[str, str], bool]:
    """Stage 1. Returns (row, ran) — ran=False means Kompass was skipped
    entirely (nothing missing, or Kompass disabled/unreachable)."""
    gaps = missing_columns(row)
    if not gaps:
        _detail("Kompass: skipped (no missing fields)")
        return row, False

    company_name = str(row.get("Company") or "").strip()
    if not company_name:
        _detail("Kompass: skipped (no company name to look up)")
        return row, False

    try:
        kompass_row = _kompass_row_for_company(company_name, market)
    except Exception as exc:  # noqa: BLE001
        _detail(f"Kompass: error — {type(exc).__name__}: {exc} — continuing to next stage")
        return row, False

    if not kompass_row:
        _detail("Kompass: no matching company profile found")
        return row, False

    merged = merge_company_data(row, kompass_row, source="kompass")
    for col in gaps:
        _detail(f"{col}: {'FOUND' if not is_missing(merged.get(col)) else 'still missing'} (Kompass)")
    return merged, True


# ---------------------------------------------------------------------------
# Stage 2 — ChatGPT scraper (free, web-UI based)
# ---------------------------------------------------------------------------


def _parse_field_response(text: str, fields: list[str]) -> dict[str, str]:
    """Best-effort parse of a free-text/JSON-ish AI reply into {field: value}.
    Tries JSON first (AI was asked to reply in JSON); falls back to a simple
    "Field: value" line parser for scrapers that can't be forced into a
    strict JSON response format (e.g. the free ChatGPT web-UI scraper)."""
    import json
    import re

    text = (text or "").strip()
    if not text:
        return {}

    for candidate in (text, text.strip("`")):
        start, end = candidate.find("{"), candidate.rfind("}")
        if start >= 0 and end > start:
            try:
                data = json.loads(candidate[start : end + 1])
                if isinstance(data, dict):
                    return {str(k): str(v) for k, v in data.items() if v is not None}
            except (json.JSONDecodeError, ValueError):
                pass

    out: dict[str, str] = {}
    for line in text.splitlines():
        m = re.match(r"\s*[-*]?\s*([A-Za-z ./]+?)\s*[:=]\s*(.+)\s*$", line)
        if not m:
            continue
        key, val = m.group(1).strip(), m.group(2).strip()
        for f in fields:
            if key.lower() == f.lower() or key.lower().replace(" ", "") == f.lower().replace(" ", ""):
                out[f] = val
                break
    return out


_CONTACT_ROLE_PROMPT_HINT = (
    'For "Contact Person"/"Role"/"LinkedIn": ONLY return someone whose title is '
    "sales or procurement LEADERSHIP — e.g. VP of Sales, Head of Sales, Sales "
    "Director/Manager, Head of Procurement, Procurement/Purchasing Manager or "
    "Director, Chief Procurement Officer. Do NOT return a CEO, founder, generic "
    'manager, or any other role. If no one at this company holds one of those '
    'specific titles, leave "Contact Person", "Role", and "LinkedIn" all empty '
    "rather than return someone in the wrong role."
)

_GEOGRAPHY_COLUMNS: tuple[str, ...] = ("Continent / Geography", "Operational Presence")

_GEOGRAPHY_PROMPT_HINT = (
    'These are two DIFFERENT questions for a distributor company — answer '
    "each independently, do not assume one implies the other:\n"
    '"Continent / Geography" = where the company is physically BASED — its '
    'headquarters and any physical offices/facilities. Format as REGION('
    'Country, Country, ...) grouped by continent, e.g. "APAC(India, '
    'Singapore); EU(Germany, France)" — use region codes NA, EU, ME, AF, '
    "LATAM, APAC. List every physical office location you can confirm, not "
    "just the HQ, if the company has more than one.\n"
    '"Operational Presence" = the countries/regions where this company, AS A '
    "DISTRIBUTOR, actually DISTRIBUTES OR SELLS products — its sales/"
    "distribution territory. This is NOT automatically the same as where it "
    "is headquartered, and it is NOT automatically wherever its supplier "
    "brands/manufacturers are from — a distributor carrying international "
    "brands is not evidence it distributes internationally. Check the "
    "company's own distributor/dealer-network/coverage pages for the "
    "specific territory they state, e.g. \"authorized distributor for "
    '[country/region]" or "dealer network covering [territory]". Format as '
    'a plain semicolon-separated list of country names, e.g. "India; '
    'Singapore; Germany". If the company only distributes within its own '
    "home country, that single country is the correct and complete answer — "
    "do not pad it with other countries just because the company is "
    "internationally connected in other ways (brands carried, manufacturer "
    "partnerships, etc). List every distribution country you can actually "
    "confirm from the company's own stated coverage, not more, not fewer."
)


# Core Categories / Specialty Focus / Key Brands Represented mean something
# different for every market (e.g. "Core Categories" for a Flexible
# Packaging distributor vs. a Smartphone distributor is a completely
# different kind of answer) — same "define, then distill, then search"
# pattern as vendor_intel.discovery.kompass_ai_search.understand_market(),
# reused here for these 3 fields instead of a Kompass page.
_MARKET_FIELD_COLUMNS: tuple[str, ...] = ("Core Categories", "Specialty Focus", "Key Brands Represented")

_MARKET_FIELD_UNDERSTANDING_CACHE_NAMESPACE = "market_field_understanding"


def market_field_understanding_enabled() -> bool:
    return (
        os.getenv("MARKET_FIELD_UNDERSTANDING_ENABLED", "true") or "true"
    ).strip().lower() not in ("0", "false", "no", "off")


def _define_market_fields_via_chatgpt_scraper(market: str) -> str:
    """Step A: ask the free ChatGPT web-UI scraper what "Core Categories",
    "Specialty Focus", and "Key Brands Represented" concretely mean for
    companies/distributors in THIS specific market. Returns "" on any
    failure so the caller falls back to the generic (market-agnostic)
    prompt wording, same as before this feature existed."""
    from vendor_intel.scraping.google_ai_mode import ask_generative, generative_backend_enabled

    if not generative_backend_enabled():
        _detail("market field understanding: skipped (no generative web backend enabled)")
        return ""
    try:
        reply = ask_generative(
            f"For a company/distributor in the {market} market, explain concretely:\n"
            '1. What would "Core Categories" mean — the main product/service '
            "categories a company in this market would be classified under?\n"
            '2. What would "Specialty Focus" mean — what kinds of specific '
            "specializations/niches exist within this market?\n"
            '3. What would "Key Brands Represented" mean — what kind of '
            "brands/manufacturers would a distributor in this market typically "
            "carry (give examples of real, well-known brand names in this "
            "specific market if you know them)?\n\n"
            "Answer in 3 short paragraphs (one per question), no markdown, no "
            "preamble — just the explanation."
        )
        return (reply or "").strip()
    except Exception as exc:  # noqa: BLE001
        _detail(f"market field understanding: ChatGPT scraper error — {type(exc).__name__}: {exc}")
        return ""


def _distill_market_field_definitions_via_deepseek(market: str, raw_explanation: str) -> str:
    """Step B: DeepSeek (OpenAI-compatible dispatch, same as the rest of
    this repo) distills the raw multi-paragraph explanation into a short,
    prompt-ready block of market-specific guidance for the 3 fields. Returns
    "" on any failure — caller falls back to the generic prompt wording."""
    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        _detail("market field understanding: DeepSeek skipped (OPENAI_API_KEY not set)")
        return ""
    try:
        from vendor_intel.config import Settings
        from vendor_intel.pipeline.chatgpt_expand import _chat_json, _client, _model

        settings = Settings()
        client = _client(settings)
        model = _model(settings)
        system = (
            "You distill a raw explanation of market-specific terminology into "
            "short, precise search guidance. Do not invent categories, "
            "specializations, or brand names not implied by the explanation."
        )
        user = (
            f'Market: "{market}"\n'
            f'Raw explanation:\n"""{raw_explanation}"""\n\n'
            "Reply with a single JSON object with exactly these 3 keys, each a "
            "short 1-2 sentence guidance string for searching real companies in "
            "this market:\n"
            '{"core_categories_guidance": "...", "specialty_focus_guidance": "...", '
            '"key_brands_guidance": "..."}'
        )
        data = _chat_json(
            client, model, system, user, label="market-field-understand", max_tokens=400, retries=2
        )
        if not isinstance(data, dict):
            return ""
        parts = [
            f"Core Categories for this market: {data.get('core_categories_guidance', '')}".strip(),
            f"Specialty Focus for this market: {data.get('specialty_focus_guidance', '')}".strip(),
            f"Key Brands Represented for this market: {data.get('key_brands_guidance', '')}".strip(),
        ]
        parts = [p for p in parts if p and not p.endswith(":")]
        return "\n".join(parts)
    except Exception as exc:  # noqa: BLE001
        _detail(f"market field understanding: DeepSeek error — {type(exc).__name__}: {exc}")
        return ""


def understand_market_fields(market: str) -> str:
    """ChatGPT scraper explains what Core Categories/Specialty Focus/Key
    Brands Represented concretely mean for THIS market -> DeepSeek distills
    that into compact prompt guidance -> cached and reused for every company
    processed in this market (the meaning doesn't change per company).
    Returns "" (generic prompt wording used instead) if disabled or either
    step is unavailable/fails — this is a quality improvement, never a hard
    requirement for the pipeline to run."""
    if not market or not market_field_understanding_enabled():
        return ""

    from vendor_intel.clients.search_cache_disk import get_disk_cached_search, put_disk_cached_search

    cached = get_disk_cached_search(market, "", _MARKET_FIELD_UNDERSTANDING_CACHE_NAMESPACE)
    if cached:
        return str(cached[0].get("guidance") or "")

    raw = _define_market_fields_via_chatgpt_scraper(market)
    if not raw:
        return ""
    _detail(f"market field definition for {market!r}: {raw[:160]}{'…' if len(raw) > 160 else ''}")

    guidance = _distill_market_field_definitions_via_deepseek(market, raw)
    if not guidance:
        return ""
    _detail(f"market field guidance for {market!r} ready")

    put_disk_cached_search(
        market, "", _MARKET_FIELD_UNDERSTANDING_CACHE_NAMESPACE, [{"raw": raw, "guidance": guidance}]
    )
    return guidance


def _field_search_prompt(
    company_name: str,
    website: str,
    headquarters: str,
    fields: list[str],
    *,
    market_field_guidance: str = "",
) -> str:
    ctx = [f"Company: {company_name}"]
    if website:
        ctx.append(f"Website: {website}")
    if headquarters and not is_missing(headquarters):
        ctx.append(f"Headquarters/location: {headquarters}")
    ctx_text = "\n".join(ctx)
    field_list = "\n".join(f"- {f}" for f in fields)
    hints = []
    if any(f in _CONTACT_PERSON_COLUMNS for f in fields):
        hints.append(_CONTACT_ROLE_PROMPT_HINT)
    if any(f in _GEOGRAPHY_COLUMNS for f in fields):
        hints.append(_GEOGRAPHY_PROMPT_HINT)
    if market_field_guidance and any(f in _MARKET_FIELD_COLUMNS for f in fields):
        hints.append(
            "Market-specific guidance for these fields (use this to make sure "
            f"your answer is correct for THIS market, not generic):\n{market_field_guidance}"
        )
    contact_hint = ("\n\n" + "\n\n".join(hints)) if hints else ""
    return (
        f"{ctx_text}\n\n"
        f"Find ONLY the following specific pieces of information about this company "
        f"(a real, verifiable company — do not guess or invent facts):\n{field_list}"
        f"{contact_hint}\n\n"
        "Reply with a JSON object using exactly these field names as keys. "
        'If you cannot find a value for a field, use an empty string "" for it — '
        "never invent a plausible-sounding value. No markdown, no commentary, JSON only."
    )


def enrich_from_chatgpt_scraper(
    row: dict[str, str], *, market: str = "", timeout: float = 90
) -> tuple[dict[str, str], bool]:
    """Stage 2. Uses the free ChatGPT web-UI scraper (vendor_intel.scraping
    .chatgpt_scraper), gated by CHATGPT_SCRAPER_ENABLED same as elsewhere."""
    gaps = missing_columns(row)
    if not gaps:
        _detail("web LLM: skipped (no missing fields)")
        return row, False

    from vendor_intel.scraping.google_ai_mode import ask_generative, generative_backend_enabled

    if not generative_backend_enabled():
        _detail("web LLM: skipped (no generative web backend enabled)")
        return row, False

    company_name = str(row.get("Company") or "").strip()
    if not company_name:
        return row, False

    _detail(f"web LLM: searching {len(gaps)} missing field(s): {', '.join(gaps)}")
    market_field_guidance = understand_market_fields(market) if market else ""
    prompt = _field_search_prompt(
        company_name,
        str(row.get("Website") or ""),
        str(row.get("Headquarters") or ""),
        gaps,
        market_field_guidance=market_field_guidance,
    )
    try:
        reply = ask_generative(prompt, timeout=timeout)
    except Exception as exc:  # noqa: BLE001
        _detail(f"web LLM: error — {type(exc).__name__}: {exc} — continuing to next stage")
        return row, False

    found = _parse_field_response(reply, gaps)
    merged = merge_company_data(row, found, source="chatgpt_scraper")
    for col in gaps:
        _detail(f"{col}: {'FOUND' if not is_missing(merged.get(col)) else 'MISSING'} (ChatGPT)")
    return merged, True


# ---------------------------------------------------------------------------
# Stage 3 — Google AI Overview scraper
# ---------------------------------------------------------------------------

# SLIM_COLUMNS <-> gap_fill crawl-key names, matching search_stack_fill's
# existing landscape_to_company/company_to_landscape mapping exactly so the
# same conversion logic is reused rather than re-derived here.
_SLIM_TO_CRAWL_KEY: dict[str, str] = {
    "Founded": "founded_year",
    "Headquarters": "headquarters",
    "Continent / Geography": "offices",
    "Operational Presence": "regions_served",
    "Ownership": "ownership",
    "Employees": "employee_count",
    "LinkedIn": "linkedin",
    "Core Categories": "core_categories",
    "Specialty Focus": "specialty_focus",
    "Key Brands Represented": "key_brands_represented",
    "Online / Offline / Both": "retail_ecommerce",
    "Contact Person": "contact_person",
    "Role": "contact_role",
    "Email": "email",
    "Office No.": "phone",
    "Office Email": "office_email",
    "Website": "website",
    "Summary": "summary",
}
_CRAWL_KEY_TO_SLIM: dict[str, str] = {v: k for k, v in _SLIM_TO_CRAWL_KEY.items()}


async def enrich_from_google_ai(row: dict[str, str], *, market: str) -> tuple[dict[str, str], bool]:
    """Stage 3. Reuses enrichment.gap_fill.sources.fill_from_google_ai_scraper
    unchanged — that function already does true per-field targeting (one
    query per missing column) against the existing Google AI Overview
    scraper relay."""
    gaps = missing_columns(row)
    if not gaps:
        _detail("Google AI Overview: skipped (no missing fields)")
        return row, False

    from vendor_intel.clients.google_ai_scraper import google_ai_gap_fill_enabled, google_ai_ping
    from vendor_intel.enrichment.gap_fill.sources import fill_from_google_ai_scraper
    from vendor_intel.pipeline.search_stack_fill import landscape_to_company, company_to_landscape

    if not google_ai_gap_fill_enabled():
        _detail("Google AI Overview: skipped (GOOGLE_AI_SCRAPER_ENABLED / gap fill disabled)")
        return row, False
    try:
        if not await google_ai_ping():
            _detail("Google AI Overview: skipped (relay not reachable)")
            return row, False
    except Exception as exc:  # noqa: BLE001
        _detail(f"Google AI Overview: ping failed — {type(exc).__name__}: {exc}")
        return row, False

    _detail(f"Google AI Overview: searching {len(gaps)} missing field(s): {', '.join(gaps)}")
    company_dict = landscape_to_company(row, market=market)
    try:
        filled_keys = await fill_from_google_ai_scraper(company_dict)
    except Exception as exc:  # noqa: BLE001
        _detail(f"Google AI Overview: error — {type(exc).__name__}: {exc} — continuing to next stage")
        return row, False

    updated_row = company_to_landscape(company_dict, family=market, base=row, source="google_ai")
    merged = merge_company_data(row, updated_row, source="google_ai")
    for col in gaps:
        _detail(f"{col}: {'FOUND' if not is_missing(merged.get(col)) else 'MISSING'} (Google AI)")
    _ = filled_keys
    return merged, True


# ---------------------------------------------------------------------------
# Stage 4 — DeepSeek (OpenAI-compatible), targeted field fill only
# ---------------------------------------------------------------------------


def enrich_from_deepseek(row: dict[str, str], *, market: str) -> tuple[dict[str, str], bool]:
    """Stage 4. Uses chatgpt_expand._chat_json — the same low-level
    DeepSeek/OpenAI dispatch already used throughout the expand pipeline —
    but with a NEW targeted-field prompt asking only for the still-missing
    columns, instead of the existing gpt_fill_rows' whole-row prompt."""
    gaps = missing_columns(row)
    if not gaps:
        _detail("DeepSeek: skipped (no missing fields)")
        return row, False

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        _detail("DeepSeek: skipped (OPENAI_API_KEY not set)")
        return row, False

    from vendor_intel.config import Settings
    from vendor_intel.pipeline.chatgpt_expand import _chat_json, _client, _model

    company_name = str(row.get("Company") or "").strip()
    if not company_name:
        return row, False

    _detail(f"DeepSeek: searching {len(gaps)} missing field(s): {', '.join(gaps)}")
    try:
        settings = Settings()
        client = _client(settings)
        model = _model(settings)
    except Exception as exc:  # noqa: BLE001
        _detail(f"DeepSeek: client setup failed — {type(exc).__name__}: {exc}")
        return row, False

    system = (
        "You are a company-research assistant. You only report facts you are "
        "reasonably confident are true and verifiable. You NEVER invent or "
        "guess plausible-sounding values for fields you don't actually know."
    )
    market_field_guidance = understand_market_fields(market)
    user = _field_search_prompt(
        company_name,
        str(row.get("Website") or ""),
        str(row.get("Headquarters") or ""),
        gaps,
        market_field_guidance=market_field_guidance,
    ) + '\n\nRespond with a single JSON object: {"field_name": "value", ...}.'

    try:
        data = _chat_json(client, model, system, user, label="enrich-deepseek", max_tokens=1024, retries=2)
    except Exception as exc:  # noqa: BLE001
        _detail(f"DeepSeek: error — {type(exc).__name__}: {exc} — continuing to next stage")
        return row, False

    found = {str(k): str(v) for k, v in (data or {}).items()} if isinstance(data, dict) else {}
    merged = merge_company_data(row, found, source="deepseek")
    for col in gaps:
        _detail(f"{col}: {'FOUND' if not is_missing(merged.get(col)) else 'MISSING'} (DeepSeek)")
    return merged, True


# ---------------------------------------------------------------------------
# Website verification — reachability + identity check, with ChatGPT-scraper
# domain correction for dead/wrong websites. Runs after the 4 discovery
# stages (verifying/fixing whatever website any of them found) and before
# Smart Crawl (which needs a confirmed-real domain, not a dead/wrong one).
# ---------------------------------------------------------------------------


def verify_website_stage(row: dict[str, str]) -> tuple[dict[str, str], bool]:
    """Not a fallback-chain data source in the same sense as the 5 stages —
    always runs when a Website value exists or a company name is known,
    regardless of missing_columns(), because a dead/wrong website is a
    correctness problem even if the "Website" column is technically
    non-empty. Never invents a website: if nothing verifiable is found, the
    field is left empty rather than kept wrong."""
    company_name = str(row.get("Company") or "").strip()
    if not company_name:
        _detail("Website verify: skipped (no company name)")
        return row, False

    from vendor_intel.pipeline.website_verification import verify_and_fix_website

    current_website = str(row.get("Website") or "")
    headquarters = str(row.get("Headquarters") or "")
    try:
        result = verify_and_fix_website(company_name, headquarters, current_website)
    except Exception as exc:  # noqa: BLE001
        _detail(f"Website verify: error — {type(exc).__name__}: {exc} — keeping existing value")
        return row, False

    out = dict(row)
    if result.was_corrected:
        out["Website"] = result.website
        _detail(f"Website: CORRECTED -> {result.website or '(none found)'}")
        return out, True
    if not result.website and current_website:
        # Confirmed dead with no replacement — clear it rather than export a
        # known-bad domain (per spec: fix dead domains, don't just leave them).
        out["Website"] = ""
        _detail("Website: confirmed dead, no replacement found — cleared")
        return out, True

    _detail(f"Website: verified OK ({result.reason or 'reachable + identity match'})")
    return out, False


# ---------------------------------------------------------------------------
# Stage 5 — Smart Crawl (website-derived fields only)
# ---------------------------------------------------------------------------


def _domain_from_website(website: str) -> str:
    from urllib.parse import urlparse

    w = (website or "").strip()
    if not w:
        return ""
    if "://" not in w:
        w = f"https://{w}"
    host = urlparse(w).netloc.lower()
    return host[4:] if host.startswith("www.") else host


def _extract_website_fields_from_intel(intel: dict[str, Any]) -> dict[str, str]:
    """Map crawler/smart_crawl.py's INTEL_SCHEMA onto the website-derived
    SLIM columns. No schema field exists for "Key Brands" or "Specialty
    Focus" directly — derived from the closest existing fields
    (relationships.partnerships/distributors/suppliers, intel.strategic_focus
    + business.competitive_advantages) rather than modifying smart_crawl's
    schema (keeps Phase 3 / --full-crawl behavior untouched)."""
    business = intel.get("business") or {}
    relationships = intel.get("relationships") or {}
    contact = intel.get("contact") or {}
    intel_block = intel.get("intel") or {}

    def _join(items: Any, limit: int = 8) -> str:
        if not isinstance(items, list):
            return ""
        vals = [str(x).strip() for x in items if str(x or "").strip()]
        return "; ".join(vals[:limit])

    products = business.get("products") or []
    key_products = "; ".join(
        str(p.get("name") or "").strip() for p in products if isinstance(p, dict) and p.get("name")
    )
    core_categories = _join(business.get("industries")) or _join(business.get("key_markets"))
    if not core_categories and products:
        core_categories = _join([p.get("category") for p in products if isinstance(p, dict)])

    brand_sources = list(relationships.get("partnerships") or []) + list(
        relationships.get("distributors") or []
    ) + list(relationships.get("suppliers") or [])
    key_brands = _join(brand_sources)

    specialty = _join(intel_block.get("strategic_focus")) or _join(business.get("competitive_advantages"))

    out = {
        "Key Products": key_products,
        "Core Categories": core_categories,
        "Key Brands Represented": key_brands,
        "Specialty Focus": specialty,
        "Office No.": str(contact.get("phone") or "").strip(),
        "Office Email": str(contact.get("email") or "").strip(),
    }
    return {k: v for k, v in out.items() if v}


async def enrich_from_smart_crawl(row: dict[str, str]) -> tuple[dict[str, str], bool]:
    """Stage 5. Uses crawler.smart_crawl.smart_crawl(mode="company") unchanged.
    Only runs for the website-derived fields (spec section on Smart Crawl),
    and only if a website is known by this point in the chain — from Kompass,
    ChatGPT, Google AI, or DeepSeek, whichever found it first."""
    website_gaps = [c for c in WEBSITE_DERIVED_COLUMNS if is_missing(row.get(c))]
    if not website_gaps:
        _detail("Smart Crawl: skipped (no missing website-derived fields)")
        return row, False

    website = str(row.get("Website") or "").strip()
    domain = _domain_from_website(website)
    if not domain:
        _detail("Smart Crawl: skipped (no website discovered by any prior stage)")
        return row, False

    _detail(f"Smart Crawl: crawling {domain} for {len(website_gaps)} field(s): {', '.join(website_gaps)}")
    try:
        from crawler.smart_crawl import smart_crawl

        result = await smart_crawl(domain, mode="company")
    except Exception as exc:  # noqa: BLE001
        _detail(f"Smart Crawl: error — {type(exc).__name__}: {exc} — preserving prior data")
        return row, False

    intel = result.get("data") if isinstance(result, dict) else None
    if not isinstance(intel, dict):
        _detail("Smart Crawl: no usable data returned")
        return row, False

    found = _extract_website_fields_from_intel(intel)
    # Smart Crawl is authoritative for website-derived columns (spec Rule 3):
    # it may override a weaker AI-guessed value for these specific columns,
    # never for anything else.
    merged = merge_company_data(
        row, found, source="smart_crawl", prefer_source_for=WEBSITE_DERIVED_COLUMNS
    )
    for col in website_gaps:
        _detail(f"{col}: {'FOUND' if not is_missing(merged.get(col)) else 'MISSING'} (Smart Crawl)")
    return merged, True


# ---------------------------------------------------------------------------
# Final stage — DeepSeek text polish (formatting only, not a data source)
# ---------------------------------------------------------------------------


def _needs_polish(value: str) -> bool:
    """Skip an already-clean value rather than spend an API call on it —
    e.g. a short, already-well-formed semicolon list. Anything with raw
    scraper noise (double spaces, stray newlines, no capitalization, overly
    long run-on text) is worth sending."""
    v = value.strip()
    if not v or is_missing(v):
        return False
    if "\n" in v or "  " in v:
        return True
    if len(v) > 400:
        return True
    return False


def polish_text_fields(row: dict[str, str]) -> tuple[dict[str, str], bool]:
    """Final formatting-only pass: cleans up phrasing/whitespace/casing on
    the long free-text columns using DeepSeek, once per company, in a single
    batched call. This is NOT a data source — it must never add, remove, or
    change the factual content of a field, only its wording/formatting. If
    DeepSeek is unavailable or the call fails, the raw (unpolished) values
    from earlier stages are kept as-is — polishing is cosmetic, never
    required for correctness."""
    candidates = {
        col: str(row.get(col) or "").strip()
        for col in POLISH_COLUMNS
        if _needs_polish(str(row.get(col) or ""))
    }
    if not candidates:
        _detail("DeepSeek polish: skipped (nothing needs cleanup)")
        return row, False

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        _detail("DeepSeek polish: skipped (OPENAI_API_KEY not set)")
        return row, False

    from vendor_intel.config import Settings
    from vendor_intel.pipeline.chatgpt_expand import _chat_json, _client, _model

    try:
        settings = Settings()
        client = _client(settings)
        model = _model(settings)
    except Exception as exc:  # noqa: BLE001
        _detail(f"DeepSeek polish: client setup failed — {type(exc).__name__}: {exc}")
        return row, False

    _detail(f"DeepSeek polish: cleaning {len(candidates)} field(s): {', '.join(candidates)}")
    system = (
        "You are a copy editor for company-intelligence data. You ONLY fix "
        "grammar, capitalization, spacing, and formatting (e.g. turn a messy "
        "list into a clean '; '-separated list, fix run-on sentences). You "
        "NEVER add, remove, invent, or change any factual content — every "
        "name, number, and claim in your output must already be present in "
        "the input. If a field is already clean, return it unchanged."
    )
    user = (
        "Clean up the formatting/wording of these fields (JSON object, same "
        "keys in your reply):\n" + "\n".join(f'- {k}: "{v}"' for k, v in candidates.items())
        + '\n\nReply with a single JSON object: {"field_name": "cleaned text", ...}.'
    )

    try:
        data = _chat_json(client, model, system, user, label="enrich-polish", max_tokens=1024, retries=2)
    except Exception as exc:  # noqa: BLE001
        _detail(f"DeepSeek polish: error — {type(exc).__name__}: {exc} — keeping unpolished text")
        return row, False

    if not isinstance(data, dict):
        _detail("DeepSeek polish: no usable response — keeping unpolished text")
        return row, False

    out = dict(row)
    changed = 0
    for col, cleaned in data.items():
        if col not in candidates:
            continue  # ignore anything the model returns that we didn't ask for
        cleaned_s = str(cleaned or "").strip()
        if not cleaned_s or is_missing(cleaned_s):
            continue  # never let polishing turn a real value into empty
        out[col] = cleaned_s
        changed += 1
        _detail(f"{col}: polished")
    return out, changed > 0


# ---------------------------------------------------------------------------
# Summary — composed by DeepSeek from the row's own already-verified data,
# after all enrichment stages have run. Never invents facts: only asked to
# write up what's already in the row, grounded and cited back to those
# fields, never new information.
# ---------------------------------------------------------------------------

_SUMMARY_SOURCE_COLUMNS: tuple[str, ...] = (
    "Founded", "Headquarters", "Continent / Geography", "Operational Presence",
    "Ownership", "Employees", "Core Categories", "Specialty Focus",
    "Key Brands Represented", "Online / Offline / Both",
)


def _looks_like_real_summary(value: str) -> bool:
    """Distinguish an actual written summary from Kompass's raw
    'Key: value | Key: value' metadata dump (see kompass_distributors
    .to_slim_row) — the latter is not prose and should still be rewritten
    into a real summary once enough enrichment data exists to write one."""
    v = (value or "").strip()
    if not v or is_missing(v):
        return False
    if v.count("|") >= 2 and ":" in v:
        return False  # looks like the raw "Key: value | Key: value" format
    return True


def write_company_summary(row: dict[str, str], *, market: str = "") -> tuple[dict[str, str], bool]:
    """Compose (or improve) the Summary column via DeepSeek, using only the
    company's own already-populated, already-verified fields as source
    material — never invents new facts. Skipped if DeepSeek isn't
    configured, or if there isn't enough real data yet to summarize."""
    company_name = str(row.get("Company") or "").strip()
    if not company_name:
        return row, False

    current = str(row.get("Summary") or "")
    if _looks_like_real_summary(current):
        _detail("Summary: already a real written summary — skipped")
        return row, False

    facts = {
        col: str(row.get(col) or "").strip()
        for col in _SUMMARY_SOURCE_COLUMNS
        if not is_missing(row.get(col))
    }
    if len(facts) < 2:
        _detail("Summary: skipped (not enough verified data yet to summarize)")
        return row, False

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        _detail("Summary: skipped (OPENAI_API_KEY not set)")
        return row, False

    from vendor_intel.config import Settings
    from vendor_intel.pipeline.chatgpt_expand import _chat_json, _client, _model

    try:
        settings = Settings()
        client = _client(settings)
        model = _model(settings)
    except Exception as exc:  # noqa: BLE001
        _detail(f"Summary: client setup failed — {type(exc).__name__}: {exc}")
        return row, False

    system = (
        "You write short, factual company summaries for a B2B intelligence "
        "database. You ONLY use facts explicitly given to you — you NEVER add, "
        "infer, or invent any fact (founding year, location, employee count, "
        "products, brands, etc.) that isn't already present in the input. If "
        "the given facts are thin, write a short summary using only what's "
        "given rather than padding it with invented detail."
    )
    facts_text = "\n".join(f"- {k}: {v}" for k, v in facts.items())
    user = (
        f"Company: {company_name}\n"
        + (f"Market: {market}\n" if market else "")
        + f"Known facts:\n{facts_text}\n\n"
        "Write a concise 2-3 sentence company summary using ONLY these facts. "
        'Reply with a single JSON object: {"summary": "..."}.'
    )

    try:
        data = _chat_json(client, model, system, user, label="enrich-summary", max_tokens=400, retries=2)
    except Exception as exc:  # noqa: BLE001
        _detail(f"Summary: error — {type(exc).__name__}: {exc} — keeping existing value")
        return row, False

    summary_text = str((data or {}).get("summary") or "").strip() if isinstance(data, dict) else ""
    if not summary_text or is_missing(summary_text):
        _detail("Summary: no usable response — keeping existing value")
        return row, False

    out = dict(row)
    out["Summary"] = summary_text
    _detail("Summary: written")
    return out, True


# ---------------------------------------------------------------------------
# Orchestrator
# ---------------------------------------------------------------------------


async def enrich_company(
    row: dict[str, str],
    *,
    market: str,
    company_index: int | None = None,
    total_companies: int | None = None,
) -> EnrichedCompany:
    """Run the full Kompass -> ChatGPT -> Google AI -> DeepSeek -> Smart Crawl
    fallback chain for ONE company row, in memory. Every stage only searches
    for columns still missing after the previous stage; a stage with nothing
    left to search for is skipped without any network/API call."""
    company_name = str(row.get("Company") or "").strip() or "(unnamed)"
    prefix = f"[Company {company_index}]" if company_index is not None else "[Company]"
    if total_companies is not None and company_index is not None:
        prefix = f"[Company {company_index}/{total_companies}]"
    _log(f"{prefix} {company_name}")

    result = EnrichedCompany(row=dict(row))
    field_sources: dict[str, str] = {}

    stage_fns: list[tuple[str, Any]] = [
        ("kompass", lambda r: enrich_from_kompass(r, market=market)),
        ("chatgpt_scraper", lambda r: enrich_from_chatgpt_scraper(r, market=market)),
        ("google_ai", lambda r: enrich_from_google_ai(r, market=market)),
        ("deepseek", lambda r: enrich_from_deepseek(r, market=market)),
        # Runs AFTER the 4 discovery stages (so it verifies/fixes whatever
        # website any of them found) but BEFORE Smart Crawl (which needs a
        # confirmed-real domain to crawl, not a dead/wrong one).
        ("website_verify", lambda r: verify_website_stage(r)),
        ("smart_crawl", lambda r: enrich_from_smart_crawl(r)),
        # Runs LAST (before the polish pass) so the summary reflects
        # everything found by every prior stage, including Smart Crawl's
        # website-derived fields.
        ("write_summary", lambda r: write_company_summary(r, market=market)),
    ]

    current = dict(row)
    for source, fn in stage_fns:
        before = dict(current)
        try:
            outcome = fn(current)
            if hasattr(outcome, "__await__"):
                outcome = await outcome
            current, ran = outcome
        except Exception as exc:  # noqa: BLE001
            # One failed source must never stop the whole pipeline.
            result.errors.append(f"{source}: {type(exc).__name__}: {exc}")
            _detail(f"{source}: unexpected error — {type(exc).__name__}: {exc} — continuing")
            continue

        if not ran:
            result.stages_skipped.append(source)
            continue

        result.stages_run.append(source)
        for col in MANAGED_COLUMNS:
            if is_missing(before.get(col)) and not is_missing(current.get(col)):
                field_sources[col] = source

    # Final formatting-only pass — not a data source, so it never claims
    # field_sources credit for a column (that tracks who FOUND a value, not
    # who reformatted it); tracked only in stages_run/skipped for visibility.
    try:
        current, polished = polish_text_fields(current)
    except Exception as exc:  # noqa: BLE001
        result.errors.append(f"deepseek_polish: {type(exc).__name__}: {exc}")
        _detail(f"deepseek_polish: unexpected error — {type(exc).__name__}: {exc} — keeping unpolished text")
        polished = False
    (result.stages_run if polished else result.stages_skipped).append("deepseek_polish")

    result.row = current
    result.field_sources = field_sources
    remaining = missing_columns(current)
    if remaining:
        _detail(f"still missing after all stages: {', '.join(remaining)}")
    _log(f"{prefix} {company_name} — enrichment complete")
    return result


@dataclass
class EnrichmentSummary:
    total_companies: int = 0
    stage_run_counts: dict[str, int] = field(default_factory=lambda: {s: 0 for s in _SOURCE_ORDER})
    stage_skip_counts: dict[str, int] = field(default_factory=lambda: {s: 0 for s in _SOURCE_ORDER})
    fields_filled_by_source: dict[str, int] = field(default_factory=lambda: {s: 0 for s in _SOURCE_ORDER})
    still_missing_total: int = 0
    errors: list[str] = field(default_factory=list)
    websites_discovered: int = 0

    def record(self, enriched: EnrichedCompany, *, had_website_before: bool) -> None:
        self.total_companies += 1
        for s in enriched.stages_run:
            self.stage_run_counts[s] = self.stage_run_counts.get(s, 0) + 1
        for s in enriched.stages_skipped:
            self.stage_skip_counts[s] = self.stage_skip_counts.get(s, 0) + 1
        for col, src in enriched.field_sources.items():
            self.fields_filled_by_source[src] = self.fields_filled_by_source.get(src, 0) + 1
        self.still_missing_total += len(missing_columns(enriched.row))
        self.errors.extend(enriched.errors)
        if not had_website_before and not is_missing(enriched.row.get("Website")):
            self.websites_discovered += 1

    def log_summary(self) -> None:
        _log("=== Enrichment summary ===")
        _log(f"  Total companies: {self.total_companies}")
        for s in _SOURCE_ORDER:
            _log(
                f"  {s}: ran={self.stage_run_counts.get(s, 0)} "
                f"skipped={self.stage_skip_counts.get(s, 0)} "
                f"fields_filled={self.fields_filled_by_source.get(s, 0)}"
            )
        _log(f"  Websites discovered during enrichment: {self.websites_discovered}")
        _log(f"  Fields still missing (all companies): {self.still_missing_total}")
        _log(f"  Errors: {len(self.errors)}")
        for err in self.errors[:20]:
            _log(f"    · {err}")


def _company_checkpoint_key(row: dict[str, str], index: int) -> str:
    name = str(row.get("Company") or "").strip().lower()
    # index disambiguates duplicate/blank company names so no two rows collide
    return f"{index}:{name}" if name else f"{index}:(unnamed)"


_ENRICH_CKPT_SAVE_EVERY = 10  # persist every N companies — bounds rework on kill without excessive I/O


async def enrich_company_rows(
    rows: list[dict[str, str]],
    *,
    market: str,
    ckpt: Any = None,
) -> tuple[list[dict[str, str]], EnrichmentSummary]:
    """Run enrich_company() over a list of rows sequentially (each stage
    inside a company can be I/O-bound but stages are ordered/dependent, so
    companies are processed one at a time here; callers wanting company-level
    parallelism can wrap this with asyncio.gather over batches externally).

    ckpt (optional): an ExpandCheckpoint-shaped object (bump()/data()/save()).
    When given, resume is at COMPANY-BOUNDARY granularity — a company whose
    full 7-stage chain hasn't finished is redone from scratch on resume, but
    every already-completed company is skipped entirely rather than
    re-enriched. Progress (completed rows + running summary) is saved every
    _ENRICH_CKPT_SAVE_EVERY companies, not after every single one."""
    total = len(rows)
    done_keys: set[str] = set(ckpt.data("enrich_done_keys", []) if ckpt else [])
    out: list[dict[str, str]] = list(ckpt.data("enrich_out_rows", []) if ckpt else [])
    summary = EnrichmentSummary()
    if ckpt and done_keys:
        _log(f"=== Enrichment chain: RESUMED (checkpoint) — {len(done_keys)}/{total} companies already done ===")
        saved_summary = ckpt.data("enrich_summary_state")
        if isinstance(saved_summary, dict):
            summary.total_companies = saved_summary.get("total_companies", 0)
            summary.stage_run_counts.update(saved_summary.get("stage_run_counts", {}))
            summary.stage_skip_counts.update(saved_summary.get("stage_skip_counts", {}))
            summary.fields_filled_by_source.update(saved_summary.get("fields_filled_by_source", {}))
            summary.still_missing_total = saved_summary.get("still_missing_total", 0)
            summary.errors = list(saved_summary.get("errors", []))
            summary.websites_discovered = saved_summary.get("websites_discovered", 0)

    since_save = 0
    for i, row in enumerate(rows, start=1):
        key = _company_checkpoint_key(row, i)
        if key in done_keys:
            continue
        had_website = not is_missing(row.get("Website"))
        enriched = await enrich_company(row, market=market, company_index=i, total_companies=total)
        summary.record(enriched, had_website_before=had_website)
        out.append(enriched.row)
        done_keys.add(key)

        since_save += 1
        if ckpt and since_save >= _ENRICH_CKPT_SAVE_EVERY:
            ckpt.bump(
                "enrich_company_chain", f"company_{len(done_keys)}_of_{total}",
                enrich_done_keys=list(done_keys),
                enrich_out_rows=out,
                enrich_summary_state=_summary_to_dict(summary),
                note=f"enriched {len(done_keys)}/{total} companies",
            )
            since_save = 0

    if ckpt and since_save > 0:
        ckpt.bump(
            "enrich_company_chain", "chain_loop_done",
            enrich_done_keys=list(done_keys),
            enrich_out_rows=out,
            enrich_summary_state=_summary_to_dict(summary),
            note="company enrichment loop complete",
        )

    summary.log_summary()
    return out, summary


def _summary_to_dict(summary: EnrichmentSummary) -> dict[str, Any]:
    return {
        "total_companies": summary.total_companies,
        "stage_run_counts": dict(summary.stage_run_counts),
        "stage_skip_counts": dict(summary.stage_skip_counts),
        "fields_filled_by_source": dict(summary.fields_filled_by_source),
        "still_missing_total": summary.still_missing_total,
        "errors": list(summary.errors),
        "websites_discovered": summary.websites_discovered,
    }
