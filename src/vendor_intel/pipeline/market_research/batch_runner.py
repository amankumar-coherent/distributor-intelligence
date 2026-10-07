"""Orchestrates market-research batches: discover+draft (one combined
prompt) -> assemble up to N rows -> render Markdown table -> persist state.
Also drives the "keep going until TARGET total companies" loop and the
final Excel export.

BACKEND: Google AI Mode, via market_research.ai_mode_client (which wraps
scraping.google_ai_mode's Patchright-driven udm=50 client). It replaced the
ChatGPT-account backend because that one hit "usage limit reached" on a
free plan and stalled --target 500 runs at ~110 companies; AI Mode needs no
account, no login and has no per-account quota, and answers the same
batched JSON prompt in ~13s instead of ~166s. ai_mode_client keeps
chatgpt_oauth_client's chat()/chat_json() signatures, so reverting is a
one-line import change in discovery.py.

NOTE: per explicit user instruction, this pipeline does not independently
re-verify fields against real fetched sources (no live website fetch, no
Contact/About page fetch, no LinkedIn fetch) — it trusts the research
draft's own field values directly. verify.py's real-fetch functions still
exist (used previously, and available again if this is reverted) but are
not called from the default flow below. What IS still enforced: the
distributor-only verdict gate (never manufacturers/OEMs/system
integrators) and schema.finalize_row's "no blank/N/A/placeholder cell"
rule — a field the backend genuinely didn't know (returned "") still drops
the company rather than exporting a blank cell, since that is a separate
rule from "verify against a real source" and wasn't asked to change.

Two deterministic normalizers below (_normalize_employees,
_normalize_linkedin) repair formatting the prompt asks for but the model
doesn't always honour, rather than dropping an otherwise-complete company
over a formatting slip.
"""
from __future__ import annotations

import os
import re
import sys
import time

from vendor_intel.pipeline.market_research import discovery, schema
from vendor_intel.pipeline.market_research import state as state_module
from vendor_intel.pipeline.market_research.state import MarketResearchState


def _log(msg: str) -> None:
    """Print a progress line, never crashing the run over an encoding issue.

    A Windows cp1252 console cannot encode most non-Western characters, and
    print() RAISES on an unmappable one — which killed a run mid-batch on a
    Turkish company name containing "S" with cedilla (U+015E). Since this
    pipeline deliberately searches worldwide, non-ASCII names are normal
    input, so logging must degrade rather than abort.
    """
    text = f"  [market-research] {msg}"
    try:
        print(text, flush=True)
    except UnicodeEncodeError:
        enc = (getattr(sys.stdout, "encoding", None) or "ascii")
        print(text.encode(enc, errors="replace").decode(enc), flush=True)

# A single discover_and_draft call asking for a full 20-company response
# (each with ~15 fields, including exhaustive per-country geography lists)
# measured at ~267s against the real ChatGPT-OAuth proxy during
# integration — slow enough to risk timing out, and worse the more
# companies are requested per call. Capping each individual prompt to a
# smaller chunk keeps each call fast and reliable; _MAX_DRAFT_ROUNDS is
# raised accordingly so a full batch_size can still be reached across more,
# smaller rounds.
_MAX_DRAFT_REQUEST_SIZE = 10

# Verification will still reject some drafts (dead site, no real contact
# page found, mislabeled verdict, etc.) — so one round alone usually won't
# fill a full batch. Re-ask (still excluding everything seen so far,
# including this batch's own already-accepted rows) up to this many times
# before giving up and returning fewer than batch_size, per the spec's
# "accuracy over reaching exactly N" rule.
_MAX_DRAFT_ROUNDS = 8
# Profile resets allowed per batch when AI Mode answers in prose. Two
# is enough to clear a drifted session; more means the phrasing itself
# is the problem and retrying only burns time.
_MAX_PROSE_RESETS = 2


# A bare single number ("23,000" / "about 23,000" / "~5000") is not a
# range — the prompt asks for a range, but a model can still slip a single
# figure through. This is a deterministic safety net: convert it into a
# sensible surrounding range rather than either rejecting the whole company
# over a formatting slip, or silently exporting a bare number the user
# explicitly said not to want.
_BARE_EMPLOYEE_NUMBER_RE = re.compile(r"^\D*([\d,]{2,7})\+?\D*$")

# LinkedIn must be the CONTACT PERSON's own profile (/in/<slug>).
#
# This NO LONGER rebuilds the URL from the contact person's name. That
# construction produced URLs that looked researched but were generated
# from a name, so many pointed at a non-existent profile (or a different
# person with the same name). Per explicit user instruction, only a real
# /in/ URL supplied by the source is kept; anything else becomes "".
_LINKEDIN_PROFILE_RE = re.compile(r"linkedin\.com/in/[A-Za-z0-9\-_%]+", re.I)


def _real_or_blank(value: object) -> str:
    """Pass a contact value through, or return "" if it is a placeholder.

    Contact fields are optional now, so "N/A"/"unknown"/"none" must be
    normalised to a true blank — otherwise a placeholder string would be
    exported as though it were data.
    """
    text = str(value or "").strip()
    return "" if schema.is_placeholder(text) else text


def _real_linkedin_or_blank(raw: object) -> str:
    """Keep only a genuine /in/<slug> profile URL; never synthesise one.

    A bare "linkedin.com", a /company/ page, or a name-derived slug are
    all rejected to "" rather than repaired, because the repaired value
    could not be distinguished from a real one downstream.
    """
    text = str(raw or "").strip()
    if schema.is_placeholder(text):
        return ""
    match = _LINKEDIN_PROFILE_RE.search(text)
    if not match:
        return ""
    url = match.group(0)
    return url if url.lower().startswith("http") else f"https://www.{url}"


def _normalize_employees(raw: str) -> str:
    text = (raw or "").strip()
    if "-" in text or "–" in text or "to " in text.lower():
        return text  # already looks like a range
    m = _BARE_EMPLOYEE_NUMBER_RE.match(text)
    if not m:
        return text  # not a recognizable bare number — leave as-is
    n = int(m.group(1).replace(",", ""))
    if n <= 0:
        return text
    # Round down/up to a sensible band around n, e.g. 23,000 -> 20,000-25,000.
    step = 5_000 if n >= 10_000 else (1_000 if n >= 1_000 else 100 if n >= 100 else 10)
    low = (n // step) * step
    high = low + step
    return f"{low:,}-{high:,}"


def _echoes_market(value: str, market: str) -> bool:
    """True when a field just restates the market name.

    Such a value carries no information: "Core Categories = Global Solar
    Rooftop Market - Distributors Only" says only that the company is in
    the market we asked about, and it hides whether the company sells the
    right product -- which is exactly what the product scope is meant to
    establish. Measured before this guard: 55% of Dry Eyes rows and 41% of
    Solar Rooftop rows.

    Compares on alphanumerics only, so punctuation and the
    "- Distributors Only" filter suffix cannot smuggle it past.
    """
    def norm(text: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", str(text or "").lower())

    v = norm(value)
    if not v:
        return False
    m = norm(market)
    # Also compare against the market with the filter suffix removed, and
    # against the market's distinctive words alone.
    base = norm(re.sub(r"-\s*distributors?\s*only\s*$", "", market or "",
                       flags=re.I))
    return v in (m, base) or (len(v) > 12 and (v in m or m in v))


def _clean_category(value: str, market: str) -> str:
    """Blank a category that merely restates the market name.

    Returns "" rather than dropping the row here: an echoed category is a
    bad VALUE, not proof the company is wrong. Core Categories is in
    schema.REQUIRED_FIELDS, so finalize_row() rejects the row afterwards
    and the export audit records the reason -- which is the right outcome,
    because a company whose product category is unknown cannot be checked
    against the market's product scope.
    """
    if market and _echoes_market(value, market):
        return ""
    return value


def _build_row(item: dict[str, str], *, market: str = "") -> dict[str, str] | None:
    company = item["company"]

    # GENUINE WHOLESALE/AUTHORIZED DISTRIBUTORS ONLY, for a PHYSICAL-GOODS
    # market: manufacturers/OEMs, system integrators, channel partners/
    # resellers/VARs/dealers/agents are all dropped there, which the user
    # explicitly ruled out as NOT distributors. But that "distributor"-
    # only gate was hardcoded here regardless of market kind, so a SERVICE
    # market (inspection, testing, training, ...) had no verdict value it
    # could ever pass -- the prompt told the model to answer
    # "service_provider" while this gate accepted only "distributor",
    # silently dropping every company a service market was collecting.
    # Measured live: drone_testing passed 18% of what it discovered,
    # worst of every drone market, before this fix. accepted_verdicts_for
    # returns the SAME kind-specific set the prompt itself was built from
    # (verdict_instructions_for), so the two can no longer disagree.
    #
    # This is the one gate kept even with real-source verification
    # removed — checked against the draft's own self-reported verdict
    # field, not an external fetch.
    accepted = discovery.accepted_verdicts_for(market)
    if item.get("verdict", "").strip().lower() not in accepted:
        return None

    website = item["website"]
    if not website:
        return None

    # Office Website: no distinct regional office page is researched
    # separately, so this mirrors the company website itself, same as when
    # verification was active — not a new assumption introduced here.
    return {
        "Company": company,
        "Website": website,
        "Founded": item["founded"],
        "Headquarters": item["headquarters"],
        "Continent / Geography": item["continent_geography"],
        "Operational Presence": item["operational_presence"],
        "Ownership": item["ownership"],
        "Employees": _normalize_employees(item["employees"]),
        "Core Categories": _clean_category(item["core_categories"], market),
        "Specialty Focus": _clean_category(item["specialty_focus"], market),
        "Key Products / Technologies": item["key_products_technologies"],
        "Online / Offline / Both": item["online_offline_both"],
        # Contact fields pass through UNCHANGED and are never synthesised.
        # Previously LinkedIn was rebuilt from the contact's name into a
        # /in/<slug> URL, and Office Email fell back to Email — both
        # manufactured a value that looked researched. Per explicit user
        # instruction these must contain real data or nothing, so an
        # unknown field stays empty and schema.finalize_row keeps the row.
        "Contact Person": _real_or_blank(item.get("contact_person")),
        "Role": _real_or_blank(item.get("role")),
        "Email": _real_or_blank(item.get("email")),
        "LinkedIn": _real_linkedin_or_blank(item.get("linkedin")),
        "Office No.": _real_or_blank(item.get("office_no")),
        "Office Website": website,
        "Office Email": _real_or_blank(item.get("office_email")),
    }



# Cities that identify their country on their own. The model sometimes
# returns a bare city as the whole Headquarters value ("Osaka"), and
# treating that as an unknown country rejected Showa Boeki Co., Ltd. --
# a Japanese company, in a market where Japan is a target country.
_CITY_COUNTRY = {
    # Japan
    "tokyo": "Japan", "osaka": "Japan", "nagoya": "Japan",
    "yokohama": "Japan", "kobe": "Japan", "kyoto": "Japan",
    "fukuoka": "Japan", "sapporo": "Japan",
    # South Korea
    "seoul": "South Korea", "busan": "South Korea",
    "incheon": "South Korea", "daegu": "South Korea",
    "gwangju": "South Korea", "daejeon": "South Korea",
    "wanju": "South Korea", "jeonju": "South Korea",
    # Greater China
    "shanghai": "China", "beijing": "China", "guangzhou": "China",
    "shenzhen": "China", "qingdao": "China", "tianjin": "China",
    "hangzhou": "China", "chengdu": "China",
    "taipei": "Taiwan", "kaohsiung": "Taiwan", "taichung": "Taiwan",
    "hong kong": "Hong Kong", "kowloon": "Hong Kong",
    # Southeast Asia
    "singapore": "Singapore",
    "kuala lumpur": "Malaysia", "penang": "Malaysia",
    "johor bahru": "Malaysia", "selangor": "Malaysia",
    "shah alam": "Malaysia", "petaling jaya": "Malaysia",
    "bangkok": "Thailand", "chiang mai": "Thailand",
    "jakarta": "Indonesia", "surabaya": "Indonesia",
    "bandung": "Indonesia", "medan": "Indonesia",
    "manila": "Philippines", "quezon city": "Philippines",
    "makati": "Philippines", "cebu": "Philippines",
    "ho chi minh city": "Vietnam", "hanoi": "Vietnam",
    "saigon": "Vietnam", "da nang": "Vietnam",
    "yangon": "Myanmar", "mandalay": "Myanmar",
    # South Asia
    "mumbai": "India", "delhi": "India", "new delhi": "India",
    "bangalore": "India", "bengaluru": "India", "chennai": "India",
    "hyderabad": "India", "kolkata": "India", "pune": "India",
    "colombo": "Sri Lanka",
    # Oceania
    "sydney": "Australia", "melbourne": "Australia",
    "brisbane": "Australia", "perth": "Australia",
    "auckland": "New Zealand", "wellington": "New Zealand",
    # Western markets in the brief
    "new york": "United States", "los angeles": "United States",
    "chicago": "United States", "seattle": "United States",
    "toronto": "Canada", "vancouver": "Canada",
    "london": "United Kingdom", "hamburg": "Germany",
    "berlin": "Germany", "amsterdam": "Netherlands",
    "rotterdam": "Netherlands", "dubai": "United Arab Emirates",
    "abu dhabi": "United Arab Emirates",
}


def _hq_country(row: dict) -> str:
    """Country for the HQ of `row`, or "" when it cannot be established.

    Tries, in order: the text after the last comma (the normal
    "city, country" form), the whole string when it is itself a country
    or a city that identifies one, and finally the first country named in
    the row's own geography fields.

    The fallbacks exist because the country gate is fail-closed: without
    them a bare "Osaka" reads as an unknown country and a valid Japanese
    company is discarded.
    """
    hq = str(row.get("Headquarters") or "").strip()
    if hq:
        tail = hq.split(",")[-1].strip()
        if tail:
            city = _CITY_COUNTRY.get(tail.lower())
            if city:
                return city
            # A comma-separated tail that is not a known city is taken as
            # the country, which is the common and correct case.
            if "," in hq:
                return tail
        whole = _CITY_COUNTRY.get(hq.lower())
        if whole:
            return whole
        if tail:
            return tail

    # Last resort: the row states its own geography elsewhere.
    for field in ("Continent / Geography", "Operational Presence"):
        text = str(row.get(field) or "")
        if not text:
            continue
        inner = re.search(r"\(([^)]+)\)", text)
        candidate = (inner.group(1) if inner else text).split(",")[0]
        candidate = candidate.split(";")[0].strip()
        if candidate:
            return candidate
    return ""


def _render_markdown_table(rows: list[dict[str, str]], *,
                           market: str = "") -> str:
    # Market-aware columns: software/cloud markets omit the physical
    # sales-channel column (schema.columns_for). Keyword-only with a
    # default so callers that do not care keep the full list.
    columns = schema.columns_for(market) if market else schema.COLUMNS
    header = "| " + " | ".join(columns) + " |"
    sep = "| " + " | ".join("---" for _ in columns) + " |"
    lines = [header, sep]
    for row in rows:
        lines.append("| " + " | ".join(row.get(c, "") for c in columns) + " |")
    return "\n".join(lines)


def _source_verification_note(rows: list[dict[str, str]]) -> str:
    if not rows:
        return "**Source Verification**\n\nNo companies in this batch passed the distributor-verdict gate."
    return (
        "**Source Verification**\n\n"
        f"Distributor/channel-partner only — companies whose research draft "
        f"did not explicitly verdict as \"distributor\" were excluded. "
        f"NOTE: per explicit user instruction, all other fields (website, "
        f"contact, email, phone, LinkedIn, etc.) are used directly from the "
        f"Google AI Mode research draft as given — no independent live "
        f"website/contact-page/LinkedIn fetch is performed to confirm them."
    )


def _run_one_batch(state: MarketResearchState, market: str, *, batch_size: int) -> list[dict[str, str]]:
    """Fills up to `batch_size` rows for one batch, re-asking the backend
    (excluding everything seen so far, including rows already accepted
    within this batch) up to _MAX_DRAFT_ROUNDS times if the first round
    doesn't yield enough survivors of the distributor-verdict gate."""
    verified_rows: list[dict[str, str]] = []
    accepted_names_lower: set[str] = set()
    # website+country keys accepted in THIS batch. State covers earlier
    # batches; this covers two name variants arriving together.
    accepted_identities: set[str] = set()
    # Consecutive rounds that produced no drafts at all. One is a hiccup,
    # two in a row is evidence the market is actually dry -- see the
    # comment at the break below.
    empty_rounds = 0
    # Profile resets spent on prose answers in this batch. Bounded
    # so a backend that will only ever narrate cannot loop forever.
    prose_resets = 0

    for round_num in range(1, _MAX_DRAFT_ROUNDS + 1):
        remaining = batch_size - len(verified_rows)
        if remaining <= 0:
            break
        excluded = state.excluded_names() + [r["Company"] for r in verified_rows]
        request_size = min(remaining, _MAX_DRAFT_REQUEST_SIZE)
        _log(
            f"round {round_num}/{_MAX_DRAFT_ROUNDS}: asking for {request_size} more "
            f"companies ({len(verified_rows)}/{batch_size} filled so far)..."
        )
        drafts = discovery.discover_and_draft(
            market, excluded=excluded, count=request_size,
            # Pass everything collected so far so the prompt can steer
            # toward regions with no coverage yet, rather than repeating
            # a generic "be global" instruction the model ignores.
            collected=list(state.companies) + verified_rows,
        )
        _log(f"round {round_num}: got {len(drafts)} draft(s) back")
        if not drafts:
            # An empty answer is NOT proof the market is exhausted. It is
            # equally consistent with a transient failure or a refusal of
            # this particular phrasing, and ask() has already done its own
            # internal retries by the time we see this.
            #
            # Live cost of treating one empty round as exhaustion: Dry Eyes
            # batch 10 stopped at 6/20 companies with six of eight rounds
            # unused, dropping batch yield from +19 to +6.
            #
            # Two consecutive empties still exits promptly on a genuinely
            # dry market -- one extra query -- while recovering the far more
            # common one-off failure.
            # A PROSE answer is not an empty market. The model narrated
            # what it would have returned ("...have been identified and
            # detailed in the market research dataset ... review the
            # complete firmographic list above") while returning no array.
            # It still has candidates; the session has drifted into
            # summarising. A profile reset clears that, which is why the
            # verifier already does exactly this. Counting it as an empty
            # round instead made Polyurethane declare itself exhausted at
            # 412 discovered with companies still to find.
            if discovery.LAST_ANSWER_WAS_PROSE and prose_resets < _MAX_PROSE_RESETS:
                prose_resets += 1
                _log(
                    f"round {round_num}: answered in prose instead of JSON "
                    f"(describing a dataset it did not return) - resetting "
                    f"profile/cookies and retrying "
                    f"({prose_resets}/{_MAX_PROSE_RESETS})"
                )
                try:
                    from vendor_intel.scraping import google_ai_mode
                    google_ai_mode.reset_session(
                        reason="discovery answered in prose instead of JSON")
                except Exception as exc:  # noqa: BLE001
                    _log(f"  reset failed ({type(exc).__name__}: {exc}) - "
                         "continuing without it")
                # Deliberately NOT counted as an empty round: a recoverable
                # session fault is not evidence about the market.
                continue

            empty_rounds += 1
            if empty_rounds >= 2:
                _log(
                    f"round {round_num}: empty twice in a row — treating the "
                    "market as exhausted for this batch"
                )
                break
            _log(
                f"round {round_num}: empty answer — retrying once before "
                "giving up on this batch"
            )
            continue

        made_progress = False
        for item in drafts:
            if len(verified_rows) >= batch_size:
                break
            name_lower = item["company"].strip().lower()
            if state.is_seen(item["company"]) or name_lower in accepted_names_lower:
                _log(f"  skip (duplicate): {item['company']}")
                continue
            # Keyword-only so a test double taking just `item` keeps
            # working; several tests patch this with a 1-arg lambda.
            candidate = _build_row(item, market=market)
            if candidate is None:
                _log(f"  reject (not a distributor / no website): {item['company']}")
                continue
            # Website identity, checkable only now that the row exists.
            # The name check above cannot catch "PowernSun" vs "Power &
            # Sun" vs "Powernsun Group" — seven rows for one Dubai company
            # in the live data, and 13-19% of collected rows overall. A
            # shared domain in the SAME COUNTRY means one company; a
            # different country is a distinct national distributor and is
            # kept (Krannich Solar legitimately has five).
            identity = state_module.website_identity(candidate)
            if identity and (identity in accepted_identities
                             or identity in state.website_keys):
                _log(f"  skip (same website as an existing company): "
                     f"{item['company']}")
                continue

            # Country restriction, where the market defines one. The
            # prompt already asks for these countries only, but a prompt
            # is a request: this is the gate that makes it a guarantee.
            # An out-of-scope company here is not a weaker answer, it is
            # the wrong answer, so it is rejected rather than ranked down.
            if not discovery.country_is_allowed(
                    market, _hq_country(candidate)):
                _log(f"  reject (outside target countries): "
                     f"{item['company']} — {_hq_country(candidate) or 'unknown'}")
                continue

            index = state.last_index + len(verified_rows) + 1
            final = schema.finalize_row(candidate, index=index, market=market)
            if final is None:
                _log(f"  reject (incomplete row): {item['company']}")
                continue
            verified_rows.append(final)
            accepted_names_lower.add(name_lower)
            if identity:
                accepted_identities.add(identity)
            made_progress = True
            _log(f"  accepted #{index}: {item['company']}")

        if made_progress:
            # The counter means "consecutive rounds that produced no new
            # company", so only an accepted company clears it. Resetting
            # merely because drafts arrived was wrong: a round returning
            # ten drafts and accepting none would reset it, so an
            # all-duplicate batch never reached the stop condition and
            # burned all eight rounds.
            empty_rounds = 0
        else:
            # Same reasoning as the empty-drafts case above: a round that
            # was all duplicates or all rejects does not mean the next one
            # will be, because every accepted company is added to the
            # exclusion list and the prompt changes accordingly. Tolerate
            # one such round.
            empty_rounds += 1
            if empty_rounds >= 2:
                _log(
                    f"round {round_num}: no new companies twice in a row — "
                    "stopping this batch"
                )
                break
            _log(
                f"round {round_num}: no new companies accepted — retrying "
                "once with the updated exclusion list"
            )
            continue

    return verified_rows


def run_batch(market: str, *, batch_size: int = 20) -> str:
    """Runs one batch for `market`, returns the Markdown table + Source
    Verification section as a single string, and persists state so the next
    call (the "give me more N" continuation) excludes everything returned
    here."""
    batch_size = max(1, min(20, batch_size))
    state = MarketResearchState.load(market)

    _log(f'starting batch for "{market}" (already have {len(state.companies)} companies)')
    verified_rows = _run_one_batch(state, market, batch_size=batch_size)

    state.add_batch(verified_rows)
    state.save()
    _log(f"batch done: {len(verified_rows)} new companies (total now {len(state.companies)})")

    table = _render_markdown_table(verified_rows, market=market)
    note = _source_verification_note(verified_rows)
    return f"{table}\n\n{note}"


# A batch returning 0 new companies can mean two very different things:
# (1) ChatGPT genuinely has no more distinct distributors to suggest for
#     this market — a real "nothing more to find" stop condition, or
# (2) a transient failure (proxy timeout, usage limit temporarily hit —
#     chatgpt_oauth_client already retries several times internally before
#     giving up) happened to hit every round of that one batch — NOT
#     exhaustion, and should be retried rather than ending a --target 500
#     run at 90
#     companies just because one batch had bad luck with network timing.
# Only stop for reason (1) after this many CONSECUTIVE empty batches in a
# row — a single empty batch is treated as a possible transient failure and
# retried (with a short pause) instead of ending the run.
_MAX_CONSECUTIVE_EMPTY_BATCHES = 3

# A market can dry up without ever producing a fully EMPTY batch, which
# _MAX_CONSECUTIVE_EMPTY_BATCHES cannot detect. Polyurethane ground along
# at +1 and +2 per batch while every round returned names already held.
#
# The threshold is deliberately low (2 of a 20-company batch = 10%) and
# needs three consecutive batches to fire: a single weak batch is normal
# noise, and stopping early would leave real companies uncollected. Three
# in a row at <=2 has only ever meant the model is recycling.
_LOW_YIELD_THRESHOLD = 2
_MAX_LOW_YIELD_BATCHES = 3
_EMPTY_BATCH_RETRY_PAUSE_SECONDS = 15.0

# A CAPTCHA block needs a far longer pause than a transient empty batch:
# 15s does nothing against an IP-level rate limit. Cool-off grows with each
# blocked round (5min, 10min, 15min) to give the throttle time to expire.
# Cool-off after a batch came back empty while Google was blocking.
#
# Overridable because the wait is not always warranted: a CAPTCHA that was
# SOLVED mid-batch still marks the run as "recently blocked", so a batch
# that then came back empty for an unrelated reason (every suggestion a
# duplicate) would pay the full cool-off for nothing. Set
# MARKET_RESEARCH_BLOCK_COOLOFF=0 to retry immediately instead.
_DEFAULT_BLOCKED_COOL_OFF_SECONDS = 300.0
_DEFAULT_MAX_BLOCKED_ROUNDS = 3


def block_cool_off_seconds() -> float:
    """Cool-off after a batch came back empty while Google was blocking.

    Read at CALL time, not import time. Module-level os.getenv() ran
    before apply_env_overrides() had loaded .env, so the constant always
    captured the default and MARKET_RESEARCH_BLOCK_COOLOFF was silently
    ignored — the setting appeared to do nothing.

    0 disables the wait entirely. That is a reasonable choice: a CAPTCHA
    that was SOLVED mid-batch still marks the run "recently blocked", so a
    batch that then came back empty for an unrelated reason (every
    suggestion a duplicate) would otherwise pay 300s for nothing.
    """
    raw = (os.getenv("MARKET_RESEARCH_BLOCK_COOLOFF") or "").strip()
    if not raw:
        return _DEFAULT_BLOCKED_COOL_OFF_SECONDS
    try:
        return max(0.0, float(raw))
    except ValueError:
        return _DEFAULT_BLOCKED_COOL_OFF_SECONDS


def max_blocked_rounds() -> int:
    raw = (os.getenv("MARKET_RESEARCH_MAX_BLOCKED_ROUNDS") or "").strip()
    if not raw:
        return _DEFAULT_MAX_BLOCKED_ROUNDS
    try:
        return max(1, int(raw))
    except ValueError:
        return _DEFAULT_MAX_BLOCKED_ROUNDS

# How far back a CAPTCHA still counts as "Google is throttling us". Long
# enough to cover a batch's worth of queries (a batch can take minutes at
# 10-25s pacing), so a CAPTCHA early in a batch still explains an empty
# result at the end of it.
_BLOCK_LOOKBACK_SECONDS = 600.0

# A refusal ("no response available for this search") is a THIRD cause of
# an empty batch, distinct from both blocking and exhaustion. Google served
# the page and the model declined the prompt, so waiting cannot help and
# stopping is premature — retry soon with a fresh request instead.
_REFUSAL_LOOKBACK_SECONDS = 600.0
_REFUSAL_RETRY_PAUSE_SECONDS = 20.0
_MAX_REFUSAL_ROUNDS = 4


def _backend_is_blocked() -> bool:
    """True when Google has been challenging us recently.

    Distinguishes "Google is refusing us" from "the market has no more
    companies" — two causes of an empty batch that must be handled
    oppositely (cool off and retry vs. stop). Returns False if the backend
    isn't in use, so the relay path is unaffected.

    Deliberately does NOT test the consecutive-CAPTCHA circuit breaker.
    That counter resets to zero whenever a CAPTCHA is solved, so a SOLVED
    CAPTCHA left no evidence and the runner reported "no sign of blocking"
    for a batch whose own log said "CAPTCHA cleared" — stopping a run at
    467/500 as "genuinely exhausted". A solve gets one query through; it
    does not lift Google's throttle, so it must still count as blocking.
    """
    try:
        from vendor_intel.scraping import google_ai_mode as aim

        if not aim.saw_captcha_recently(_BLOCK_LOOKBACK_SECONDS):
            return False
        # A CAPTCHA happened, but if AI Mode has ALSO been declining
        # prompts more recently, the refusal is the better explanation for
        # an empty batch — and refusals are not fixed by waiting. Cooling
        # off 5 minutes for a refusal wastes the run's time and changes
        # nothing (observed live: "no response available for this search"
        # followed by a 300s cool-off).
        try:
            if aim.saw_refusal_recently(_REFUSAL_LOOKBACK_SECONDS):
                return False
        except AttributeError:
            pass
        return True
    except Exception:
        return False


def _backend_refused() -> bool:
    """True when AI Mode recently declined a prompt (not a block).

    Kept separate from _backend_is_blocked because the two need opposite
    handling: a refusal is retried promptly, a block is waited out.
    """
    try:
        from vendor_intel.scraping import google_ai_mode as aim

        return aim.saw_refusal_recently(_REFUSAL_LOOKBACK_SECONDS)
    except Exception:
        return False


def _reset_backend_refusal() -> None:
    """Clear the refusal marker so each retry is judged on its own result."""
    try:
        from vendor_intel.scraping import google_ai_mode as aim

        aim._last_refusal_at = 0.0
    except Exception:
        pass


def _captcha_evidence() -> str:
    """Human-readable CAPTCHA state, for the stop message.

    Printed alongside "no sign of blocking" so that claim is auditable
    rather than asserted: an earlier run said "no sign of blocking" on a
    batch that had logged "CAPTCHA cleared", and there was no way to see
    the contradiction from the log alone.
    """
    try:
        from vendor_intel.scraping import google_ai_mode as aim

        s = aim.captcha_stats()
        since = s["seconds_since_last"]
        if since is None:
            return "no CAPTCHA seen this run"
        return f"{s['total']} CAPTCHA(s) this run, last {since}s ago"
    except Exception:
        return "CAPTCHA state unavailable"


def _reset_backend_block() -> None:
    """Clear the backend's CAPTCHA state before a retry.

    Required after a cool-off for two reasons: the circuit breaker fails
    every query fast once tripped, so without resetting it the retry can't
    even reach Google; and the recency marker must be cleared too, or the
    next empty batch would still look "blocked" and the run would cool off
    repeatedly without ever being able to conclude the market is exhausted.
    """
    try:
        from vendor_intel.scraping import google_ai_mode as aim

        aim._last_captcha_at = 0.0

        aim._consecutive_captchas = 0
    except Exception:
        pass



def _refresh_export(market: str) -> None:
    """Rewrite the market's Excel after each completed batch.

    So the file is usable DURING a long run rather than only at the end:
    the market definition is there from the start and companies appear as
    batches land, which is what makes progress inspectable on a run that
    takes hours.

    Deliberately best-effort. The export re-reads state from disk, so a
    failure here (most likely the workbook being open in Excel, which
    locks it on Windows) must never abort discovery — the companies are
    already safely persisted in the state file, and the next batch will
    rewrite the sheet anyway.

    Controlled by MARKET_RESEARCH_LIVE_EXPORT=false if the rewrite ever
    becomes a bottleneck on a very large market.
    """
    if (os.getenv("MARKET_RESEARCH_LIVE_EXPORT") or "true").strip().lower() in (
        "0", "false", "no", "off",
    ):
        return
    try:
        from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

        export_xlsx(market)
    except Exception as exc:  # noqa: BLE001
        _log(f"could not refresh the Excel ({type(exc).__name__}) — "
             "companies are safe in the state file; will retry next batch")

def run_until_target(
    market: str, *, target: int = 500, batch_size: int = 20, max_batches: int | None = None
) -> str:
    """Keeps requesting batches of `batch_size` (deduped against everything
    already found) until `target` total verified companies are reached, or
    _MAX_CONSECUTIVE_EMPTY_BATCHES batches in a row come back with zero new
    companies (nothing more to find, not just one bad-luck timeout) — per
    the spec's "give me more N" loop repeated automatically. Returns a
    summary string; the accumulated companies are in the persisted state
    file, and callers that also want an Excel file should call
    export_xlsx() afterward with the same market.
    """
    batch_size = max(1, min(20, batch_size))
    state = MarketResearchState.load(market)
    batches_run = 0
    consecutive_empty = 0
    # Consecutive batches that returned almost nothing. See
    # _LOW_YIELD_THRESHOLD for why "almost nothing" is not "nothing".
    low_yield_batches = 0
    blocked_rounds = 0
    refused_rounds = 0

    _log(f'starting run_until_target for "{market}": target={target}, already have {len(state.companies)}')

    while len(state.companies) < target:
        if max_batches is not None and batches_run >= max_batches:
            _log(f"reached max_batches={max_batches} — stopping")
            break
        before = len(state.companies)
        _log(f"--- batch {batches_run + 1}: {before}/{target} companies so far ---")
        verified_rows = _run_one_batch(state, market, batch_size=batch_size)
        if not verified_rows:
            # An empty batch means one of two very different things, and
            # conflating them is what made a CAPTCHA-blocked run report
            # itself "exhausted" and stop at 318 of 500 companies. Ask the
            # backend whether Google is currently blocking us before
            # concluding the market has no companies left.
            if _backend_is_blocked():
                blocked_rounds += 1
                if blocked_rounds > max_blocked_rounds():
                    _log(
                        "Google is still blocking after "
                        f"{max_blocked_rounds()} cool-off attempts — stopping so "
                        "progress is saved. This is a RATE LIMIT, not exhaustion: "
                        "re-run later (and raise GOOGLE_AI_MODE_QUERY_DELAY) to "
                        "continue from where this left off."
                    )
                    break
                cool_off = block_cool_off_seconds() * blocked_rounds
                if cool_off > 0:
                    _log(
                        f"batch empty because Google is BLOCKING (CAPTCHA), not because "
                        f"the market is exhausted — cooling off {cool_off:.0f}s "
                        f"({blocked_rounds}/{max_blocked_rounds()}) then retrying"
                    )
                else:
                    _log(
                        f"batch empty after a CAPTCHA — retrying immediately "
                        f"({blocked_rounds}/{max_blocked_rounds()}, cool-off disabled)"
                    )
                _reset_backend_block()
                if cool_off > 0:
                    time.sleep(cool_off)
                continue

            if _backend_refused():
                # AI Mode served the page and declined the prompt. Not a
                # block (waiting is useless) and not exhaustion (the market
                # still has companies) — retry promptly with a smaller ask.
                refused_rounds += 1
                if refused_rounds <= _MAX_REFUSAL_ROUNDS:
                    _log(
                        "AI Mode declined the prompt rather than blocking or running "
                        f"out ({refused_rounds}/{_MAX_REFUSAL_ROUNDS}) — resetting "
                        "the profile and retrying"
                    )
                    _reset_backend_refusal()
                    # Rebuild the profile, do not merely pause.
                    #
                    # A refusal is session state: the verifier proved it
                    # live, where "It looks like there's no response
                    # available for this search" repeated through browser
                    # restarts and waiting, then cleared the moment the
                    # profile was rebuilt. Sleeping 20s and re-asking the
                    # same session mostly re-earns the same page, which is
                    # why a refusal streak used to run out the budget and
                    # end the batch. Staged extensions live outside the
                    # profile, so captcha-raptor survives.
                    from vendor_intel.scraping import google_ai_mode

                    google_ai_mode.reset_session(
                        "AI Mode declined the discovery prompt")
                    time.sleep(_REFUSAL_RETRY_PAUSE_SECONDS)
                    continue
                _log(
                    f"AI Mode declined {refused_rounds} times in a row — the prompt "
                    "is likely too long or too constrained for this market; stopping "
                    "so progress is saved"
                )
                break

            consecutive_empty += 1
            if consecutive_empty >= _MAX_CONSECUTIVE_EMPTY_BATCHES:
                _log(
                    f"{consecutive_empty} consecutive empty batches with no sign of "
                    f"blocking ({_captcha_evidence()}) — treating as genuinely "
                    "exhausted, stopping"
                )
                break
            _log(
                f"batch returned 0 new companies (possible transient failure, "
                f"{consecutive_empty}/{_MAX_CONSECUTIVE_EMPTY_BATCHES}) — "
                f"pausing {_EMPTY_BATCH_RETRY_PAUSE_SECONDS:.0f}s and retrying"
            )
            time.sleep(_EMPTY_BATCH_RETRY_PAUSE_SECONDS)
            continue
        blocked_rounds = 0
        refused_rounds = 0
        consecutive_empty = 0
        state.add_batch(verified_rows)
        state.save()
        batches_run += 1
        _log(f"batch {batches_run} done: +{len(verified_rows)} new (total now {len(state.companies)}/{target})")
        _refresh_export(market)
        if len(state.companies) == before:
            break

        # A market can dry up WITHOUT ever producing a fully empty batch,
        # which the consecutive_empty counter above cannot see.
        #
        # Observed on Polyurethane: batches of +1 and +2 while every round
        # returned one or two already-collected names (211 website
        # duplicates and 29 name duplicates rejected across the run). At
        # that rate the remaining ~95 companies would never arrive, and
        # every batch spent queries the next market needs.
        #
        # Stopping here is not giving up on a target: it is reporting the
        # number the market actually contains. Padding it would mean
        # accepting the duplicates the dedupe just rejected.
        if len(verified_rows) <= _LOW_YIELD_THRESHOLD:
            low_yield_batches += 1
            if low_yield_batches >= _MAX_LOW_YIELD_BATCHES:
                _log(
                    f"{low_yield_batches} consecutive batches yielded "
                    f"<={_LOW_YIELD_THRESHOLD} new companies — the market is "
                    f"exhausted at {len(state.companies)}. Stopping rather "
                    "than spending queries re-asking for names already held."
                )
                break
        else:
            low_yield_batches = 0

    return (
        f"Reached {len(state.companies)}/{target} verified companies for "
        f"\"{market}\" across {batches_run} batch(es)."
    )
