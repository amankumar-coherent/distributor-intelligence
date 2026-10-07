"""Batch company screening via the free ChatGPT web-UI scraper.

Kompass's own per-company profile pages carry authoritative role tags
("distributor" / "producer" / "service" per activity), but fetching one
page per company is slow and Kompass rate-limits it hard — concurrent
fetches came back HTTP 405 or redirected to account.kompass.com/login.

This module takes the opposite approach: Kompass is used only to harvest
company names (available in the listing's JSON-LD, no profile fetch), and
the ChatGPT scraper is asked about companies in batches — "is each of these
a distributor for <market>?" — plus a second pass to fill the export
columns.

TRADEOFF, stated plainly: this replaces structured evidence from Kompass
with an LLM's general knowledge. To limit invented answers the prompts
require an explicit "unknown" rather than a guess, and unknown companies
are treated as NOT verified. Nothing here should be read as Kompass having
confirmed anything.
"""
from __future__ import annotations

import json
import os
import re
import time
from dataclasses import dataclass, field
from typing import Any

_DEFAULT_BATCH_SIZE = 25
# Give up on screening after this many consecutive unusable batches — past
# that it's a backend problem (rate limit, block, closed browser), not bad
# luck, and continuing just wastes hours marking everything "unknown".
_MAX_CONSECUTIVE_BATCH_FAILURES = 5
# Pause between batches. Google throttles rapid-fire AI Mode queries; a live
# run degraded from working to 50-consecutive-failures around batch 45 with
# no delay at all.
_BATCH_PAUSE_SECONDS = 2.0


def _log(msg: str) -> None:
    print(f"  [chatgpt-screen] {msg}", flush=True)


def _detail(msg: str) -> None:
    print(f"      . {msg}", flush=True)


def screen_batch_size() -> int:
    raw = (os.getenv("CHATGPT_SCREEN_BATCH_SIZE") or str(_DEFAULT_BATCH_SIZE)).strip()
    try:
        n = int(raw)
    except ValueError:
        return _DEFAULT_BATCH_SIZE
    # Capped at 40: past that a single reply gets long enough that ChatGPT
    # starts truncating or dropping companies from the JSON array, which
    # costs more than the extra batch size saves.
    return max(1, min(40, n))


def screening_enabled() -> bool:
    """Gate for screening candidates instead of fetching each Kompass profile.

    Reads KOMPASS_GOOGLE_AI_SCREEN. The old KOMPASS_CHATGPT_SCREEN name is
    still honoured so an existing .env keeps working, but it refers to the
    same behaviour — screening is done by Google AI Overview now, not the
    removed ChatGPT scraper. Keeping the old gate name around after the
    rename is what silently sent a live run back to slow per-profile
    Kompass verification.
    """
    for key in ("KOMPASS_GOOGLE_AI_SCREEN", "KOMPASS_CHATGPT_SCREEN"):
        raw = (os.getenv(key) or "").strip().lower()
        if raw:
            return raw in ("1", "true", "yes", "on")
    return False


# Back-compat alias for any caller still importing the old name.
chatgpt_screening_enabled = screening_enabled


@dataclass
class ScreenedCompany:
    name: str
    profile_url: str = ""
    country_hint: str = ""
    is_distributor: bool = False
    verdict: str = "unknown"  # distributor | manufacturer | other | unknown
    reason: str = ""
    fields: dict[str, str] = field(default_factory=dict)


_ROLE_PROMPT = """For each company below, say whether it is a DISTRIBUTOR \
(reseller/wholesaler/channel partner that buys and resells other brands' \
products) in the {market} market.
{guidance}
Companies:
{companies}

Rules:
- "distributor" only if the company genuinely distributes/resells products \
in this market.
- "manufacturer" if it primarily makes/produces its own products (even if it \
also sells them directly).
- "other" if it operates in this market but is neither (e.g. pure service \
provider, consultancy, hospital).
- "unknown" if you do not actually recognise the company or cannot tell. \
Do NOT guess — an honest "unknown" is required rather than a plausible \
sounding answer.

Reply with ONLY a JSON array, no markdown fences, no commentary, one compact \
object per company in the SAME order as listed. Keep it short — no reasoning \
text, just the verdicts:
[{{"n": <number>, "verdict": "distributor|manufacturer|other|unknown"}}]"""


_FIELDS_PROMPT = """For each company below, give the requested facts.
{guidance}
Companies:
{companies}

Fields to fill for each: {fields}

Rules:
- Use "" (empty string) for anything you do not actually know. Do NOT guess \
or invent websites, dates, names, or numbers.
- Website must be the company's real official domain, not a directory or \
social page.
- Keep each value short and plain (no markdown, no sentences unless the \
field clearly wants prose).
- "Contact Person", "Role" and "LinkedIn" are ONLY for sales or procurement \
leadership (e.g. VP of Sales, Head of Sales, Sales Director, Sales Manager, \
Chief Sales Officer, Head of Procurement, Procurement/Purchasing Manager or \
Director, Chief Procurement Officer). If the person you know of holds any \
other role (CEO, founder, marketing, engineering, general manager, ...), \
leave all three empty rather than listing them.
- "Headquarters" is where the company is actually based. "Continent / \
Geography" is the region(s) of its offices/physical presence. "Operational \
Presence" is where it actually sells or distributes — these last two are \
often different, so do not just repeat the headquarters.

Reply with ONLY a JSON array, no markdown fences, no commentary, one object \
per company in the SAME order as listed:
[{{"n": <number>, "fields": {{"<field name>": "<value>"}}}}]"""


def prompt_tuning_enabled() -> bool:
    """Let DeepSeek tailor the ChatGPT prompts to the market, and reword
    them when a batch comes back unusable. Off => fixed templates only."""
    return (os.getenv("CHATGPT_SCREEN_TUNE_PROMPTS", "true") or "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )


def _deepseek_json(system: str, user: str, *, label: str, max_tokens: int = 1200) -> dict[str, Any]:
    """Call DeepSeek through the same dispatch the rest of the repo uses.
    Returns {} on any failure so callers fall back to fixed templates."""
    if not (os.getenv("OPENAI_API_KEY") or "").strip():
        return {}
    try:
        from vendor_intel.config import Settings
        from vendor_intel.pipeline.chatgpt_expand import _chat_json, _client, _model

        settings = Settings()
        data = _chat_json(
            _client(settings),
            _model(settings),
            system,
            user,
            label=label,
            max_tokens=max_tokens,
            retries=2,
        )
        return data or {}
    except Exception as exc:  # noqa: BLE001
        _detail(f"prompt tuning unavailable ({type(exc).__name__}: {exc}) — using fixed template")
        return {}


_MARKET_GUIDANCE_SYSTEM = (
    "You write short, precise guidance that helps another assistant classify "
    "companies in a specific B2B market. You never invent company names and "
    "never pad with filler."
)


def market_guidance(market: str) -> dict[str, str]:
    """Ask DeepSeek for market-specific guidance to splice into the ChatGPT
    prompts: what a distributor looks like in THIS market, and what the
    ambiguous columns mean here. Returns {} when unavailable, in which case
    the fixed templates are used unchanged."""
    if not prompt_tuning_enabled() or not market.strip():
        return {}

    cached = _GUIDANCE_CACHE.get(market.strip().lower())
    if cached is not None:
        return cached

    data = _deepseek_json(
        _MARKET_GUIDANCE_SYSTEM,
        (
            f'Market: "{market}"\n\n'
            "Reply with a single JSON object and nothing else. Keep every "
            "value under 40 words:\n"
            '{"distributor_definition": "<what counts as a DISTRIBUTOR (vs a '
            'manufacturer) specifically in this market>",\n'
            ' "category_hint": "<what \\"Core Categories\\" and '
            '\\"Specialty Focus\\" mean for companies in this market>",\n'
            ' "brands_hint": "<what kind of brands a distributor in this '
            'market would carry>"}'
        ),
        label="chatgpt-screen-guidance",
    )

    guidance = {
        k: str(data.get(k) or "").strip()
        for k in ("distributor_definition", "category_hint", "brands_hint")
    }
    guidance = {k: v for k, v in guidance.items() if v}
    _GUIDANCE_CACHE[market.strip().lower()] = guidance
    if guidance:
        _log(f"market guidance for '{market}': {guidance.get('distributor_definition', '')[:110]}")
    return guidance


_GUIDANCE_CACHE: dict[str, dict[str, str]] = {}


def reword_prompt(prompt: str, *, problem: str) -> str:
    """Retry a failed batch with a stricter reminder appended.

    An earlier version asked DeepSeek to rewrite the whole prompt and
    return it inside a JSON string. That reliably failed: the model burned
    its entire token budget (finish=length) and returned nothing, because
    re-emitting a long prompt escaped inside JSON is a bad fit for a
    JSON-mode call. Appending a short, deterministic corrective note gets
    the same benefit — a differently-worded second attempt — without
    depending on an LLM round-trip that doesn't work here."""
    if not prompt_tuning_enabled() or not prompt.strip():
        return ""
    return (
        f"{prompt}\n\n"
        f"IMPORTANT — the previous attempt failed because {problem}. "
        "Reply with ONLY the JSON array described above: no explanation, no "
        "markdown fences, no text before or after it. Include exactly one "
        'object per company, each with its "n" index, in the same order as '
        "listed."
    )


def _ask(prompt: str, *, timeout: int = 180) -> str:
    """Ask Google AI Overview one question and return the answer text.

    The DeepSeek-API, Google-AI-Mode and ChatGPT-web-UI backends were all
    removed in favour of the AI Overview scraper, which runs in the user's
    own signed-in Chrome (no login wall, no CAPTCHA, no API cost).

    IMPORTANT: Google is a SEARCH engine, so this can only answer a natural
    question about ONE subject. A batched "classify these 25 companies"
    prompt matches no web page and returns "About 0 results" — which is why
    screening asks one company per query and gets its speed from
    concurrency instead (see screen_roles_via_google_ai).
    """
    import asyncio

    from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_enabled

    if not google_ai_enabled():
        raise RuntimeError("GOOGLE_AI_SCRAPER_ENABLED is not set")
    try:
        data = asyncio.run(google_ai_ask(prompt))
        return (data.get("markdown") or "").strip()
    except Exception as exc:  # noqa: BLE001
        _detail(f"google ai overview call failed: {type(exc).__name__}: {exc}")
        return ""


def _parse_json_array(reply: str) -> list[dict[str, Any]]:
    """Extract the answer's JSON array from a raw web-LLM reply.

    Scans ALL bracketed spans and returns the last one that parses into
    non-empty dicts, rather than the first. Google AI Mode echoes the whole
    prompt above its answer, so the first "[...]" in the text is the
    *template* we asked it to follow (`[{"n": <number>, ...}]`) — which is
    not valid JSON and previously caused every batch to be discarded even
    though the real array was right below it.
    """
    text = (reply or "").strip()
    if text.startswith("```"):
        text = text.strip("`").removeprefix("json").strip()

    best: list[dict[str, Any]] = []
    for start, ch in enumerate(text):
        if ch != "[":
            continue
        # Try the widest span first, then shrink: the answer array may be
        # followed by trailing page chrome ("AI Mode response is ready").
        end = text.rfind("]")
        while end > start:
            try:
                data = json.loads(text[start : end + 1])
            except (json.JSONDecodeError, ValueError):
                end = text.rfind("]", start, end)
                continue
            if isinstance(data, list):
                dicts = [d for d in data if isinstance(d, dict)]
                if dicts:
                    best = dicts  # keep scanning; later arrays win
            break
    return best


def _ask_for_batch(prompt: str, *, expected: int, label: str) -> dict[int, dict[str, Any]]:
    """Send one batch prompt and parse the JSON array back into {n: item}.

    If the reply is unusable (error, unparseable, or clearly short), ask
    DeepSeek to reword the prompt and try once more. Returns {} if both
    attempts fail — callers leave those companies untouched rather than
    guessing."""
    attempt_prompt = prompt
    for attempt in (1, 2):
        problem = ""
        try:
            reply = _ask(attempt_prompt)
        except Exception as exc:  # noqa: BLE001
            problem = f"the request failed ({type(exc).__name__}: {exc})"
            reply = ""

        if not problem:
            items = _parse_json_array(reply)
            if not items:
                problem = "the reply was not a parseable JSON array"
            elif len(items) < expected:
                problem = (
                    f"only {len(items)} of {expected} companies were answered; "
                    "every company must get its own object"
                )
            else:
                answers: dict[int, dict[str, Any]] = {}
                for item in items:
                    try:
                        answers[int(item.get("n"))] = item
                    except (TypeError, ValueError):
                        continue
                if answers:
                    return answers
                problem = 'objects were missing the "n" index field'

        if attempt == 2:
            _detail(f"{label} failed after reworded retry: {problem}")
            return {}

        _detail(f"{label}: {problem} — asking DeepSeek to reword and retrying")
        reworded = reword_prompt(prompt, problem=problem)
        if not reworded:
            return {}
        attempt_prompt = reworded

    return {}


def _format_companies(batch: list[ScreenedCompany]) -> str:
    lines = []
    for i, c in enumerate(batch, start=1):
        bits = [f"{i}. {c.name}"]
        if c.country_hint:
            bits.append(f"({c.country_hint})")
        lines.append(" ".join(bits))
    return "\n".join(lines)


def google_ai_screening_enabled() -> bool:
    """Screen companies via the Google AI Overview scraper instead of the
    ChatGPT web UI.

    Why this exists: anonymous chatgpt.com starts demanding "Sign in to
    continue" partway through a long run, which stalls screening; driving a
    real logged-in Chrome profile turned out to be blocked too (enterprise
    policy on the profile, and Windows DPAPI prevents copying its cookies
    elsewhere). The Google AI scraper has no login wall.

    Important difference: Google is a SEARCH engine, not a chat model, so it
    cannot answer a batched "classify these 25 companies" prompt — that
    returns literally zero results. Companies are therefore asked ONE per
    query, but several queries run concurrently (see
    google_ai_concurrency), which measured ~4x faster than sequential."""
    return (os.getenv("KOMPASS_GOOGLE_AI_SCREEN", "false") or "false").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )


def google_ai_concurrency() -> int:
    """How many Google AI queries to run at once (each is a background tab)."""
    raw = (os.getenv("GOOGLE_AI_SCREEN_CONCURRENCY") or "5").strip()
    try:
        n = int(raw)
    except ValueError:
        return 5
    return max(1, min(10, n))


_MANUFACTURER_PHRASES = (
    "is a manufacturer", "is primarily a manufacturer", "manufactures",
    "designs and manufactures", "develops and manufactures",
    "manufacturer and developer", "manufacturer and brand owner",
    "produces its own", "original equipment manufacturer", "oem",
    # Live answers often phrase it as "is an original manufacturer, NOT a
    # distributor" or "operates as both a manufacturer and a distributor".
    # Without these the negated "distributor" blocked the distributor rule
    # while nothing matched the manufacturer rule, so a clearly-stated
    # manufacturer came back "unknown" (observed: APROKOREA, METAPHARM).
    "original manufacturer", "is a manufacturer and", "manufacturer, not a",
    "both a manufacturer", "develops, manufactures", "designs and produces",
    "produces proprietary", "manufactures proprietary", "its own proprietary",
)
_DISTRIBUTOR_PHRASES = (
    "is a distributor", "is primarily a distributor", "distributor of",
    "resells", "reseller", "wholesaler", "authorised distributor",
    "authorized distributor", "sales channel partner",
)
# Negations that INVERT a phrase's meaning when they appear just before it,
# e.g. "is an OEM, not a third-party reseller" must not count as distributor
# evidence (observed live: that exact sentence was misread as "distributor").
_NEGATIONS = ("not a", "not an", "rather than a", "rather than an", "instead of a", "isn't a", "is not a")
_NEGATION_WINDOW = 40  # chars before the phrase to inspect


def _negation_window(low: str, idx: int) -> str:
    """Text just before `idx` that could negate the phrase there.

    Stops at a sentence boundary: a negation in the PREVIOUS sentence says
    nothing about this one. Without this, "Epsilon is not a reseller.
    Epsilon is a distributor..." had the second clause wrongly suppressed by
    the first sentence's "not a"."""
    window = low[max(0, idx - _NEGATION_WINDOW) : idx]
    for boundary in (". ", "! ", "? ", "\n"):
        cut = window.rfind(boundary)
        if cut >= 0:
            window = window[cut + len(boundary) :]
    return window


def _mentions(low: str, phrases: tuple[str, ...]) -> str | None:
    """Return the earliest phrase occurrence that is NOT negated, else None.

    Scans every phrase and every occurrence: a negated mention of one phrase
    ("is not a reseller") must not stop a later genuine mention of another
    ("is a distributor of imported brands") from being found. Picks the
    earliest position so the quoted reason matches what the text leads with.
    """
    best_phrase: str | None = None
    best_idx: int | None = None
    for phrase in phrases:
        start = 0
        while (idx := low.find(phrase, start)) >= 0:
            window = _negation_window(low, idx)
            if not any(neg in window for neg in _NEGATIONS):
                if best_idx is None or idx < best_idx:
                    best_idx, best_phrase = idx, phrase
                break  # earliest non-negated hit for THIS phrase is enough
            start = idx + len(phrase)  # negated here; keep looking
    return best_phrase


def _sentence_around(body: str, low: str, phrase: str) -> str:
    """Quote the sentence containing the first NON-negated use of `phrase`,
    so the recorded reason is the one that actually drove the verdict rather
    than an earlier, negated mention of the same words."""
    start_at = 0
    idx = low.find(phrase)
    while idx >= 0:
        if not any(neg in _negation_window(low, idx) for neg in _NEGATIONS):
            break
        start_at = idx + len(phrase)
        idx = low.find(phrase, start_at)
    if idx < 0:
        idx = low.find(phrase)  # fall back rather than crash
    start = body.rfind(".", 0, idx) + 1
    end = body.find(".", idx)
    return body[start : end + 1 if end > 0 else len(body)].strip()


# Text that only ever appears on Google's bot-block / CAPTCHA interstitial.
_BLOCK_PAGE_MARKERS: tuple[str, ...] = (
    "unusual traffic from your computer network",
    "our systems have detected unusual traffic",
    "solving the above captcha",
    "why did this happen",
    "/sorry/index",
    "recaptcha",
)


def _market_rejection(low: str, market: str) -> str:
    """Quote Google's own words when it says the company is NOT in this market.

    The screen asks "is X a distributor or a manufacturer in the {market}
    market?", so Google often opens with a flat "No, X is not a distributor
    of wearable medical devices" and then describes what the company
    actually does. That description frequently contains distributor
    vocabulary for a DIFFERENT industry ("a B2B wholesaler for orthopedic
    shoemaking"), and because the negation lives in the previous sentence,
    the sentence-scoped negation check cannot reach it — so the company was
    accepted as a distributor of a market it does not serve.

    Being a distributor is necessary but not sufficient: it must be a
    distributor of THIS market. This gate is what enforces the second half.
    """
    if not market:
        return ""
    head = low[:400]  # the verdict is always in Google's opening lines
    market_low = market.lower().strip()
    # Match the market loosely: Google reorders and singularises it
    # ("wearable medical devices" -> "wearable medical device").
    tokens = [t for t in re.split(r"\W+", market_low) if len(t) > 3]
    for pattern in (
        r"\bis\s+not\s+a\s+(?:distributor|reseller|wholesaler|supplier|dealer)\b",
        r"\bnot\s+a\s+(?:distributor|reseller|wholesaler|supplier|dealer)\s+(?:of|for|in)\b",
        r"\bdoes\s+not\s+(?:distribute|sell|supply|deal\s+in)\b",
    ):
        m = re.search(pattern, head)
        if not m:
            continue
        # Only trust the rejection when the market is named in the SAME
        # sentence. Scanning further ahead made "No, Acme is not publicly
        # listed. Acme is a distributor of wearable medical devices." look
        # like a market rejection because the market appeared downstream.
        sentence_end = head.find(".", m.start())
        near = head[m.start() : sentence_end if sentence_end > 0 else len(head)]
        if tokens and sum(1 for t in tokens if t in near) >= max(1, len(tokens) - 1):
            end = head.find(".", m.start())
            return head[m.start() : end if end > 0 else len(head)].strip()[:160]
    return ""


def _verdict_from_overview(text: str, *, market: str = "") -> tuple[str, str]:
    """Map a Google AI Overview answer to (verdict, reason).

    Returns ("unknown", ...) when the overview doesn't clearly say — never
    guesses, so an ambiguous answer is dropped rather than accepted.

    Manufacturer evidence is checked FIRST and wins outright: a company that
    "manufactures and also distributes" is a manufacturer for our purposes,
    and the whole point of this screen is to exclude makers."""
    body = (text or "").strip()
    if not body:
        return "unknown", "no overview returned"
    low = body.lower()
    # A bot-block page is not an answer. Google serves "unusual traffic from
    # your computer network" instead of results when it rate-limits us, and
    # that page contains the echoed query string — which includes the words
    # "distributor" and "manufacturer". Parsed as prose it yielded a
    # confident "distributor" verdict for EVERY company screened during a
    # block, so this must be detected before any phrase matching.
    if any(marker in low for marker in _BLOCK_PAGE_MARKERS):
        return "unknown", "blocked by Google bot detection (CAPTCHA page)"

    if "did not match any documents" in low or "about 0 results" in low:
        return "unknown", "no search results"

    # Checked BEFORE the role phrases: an explicit "not a distributor of
    # {market}" must win even though the following sentence usually contains
    # distributor vocabulary for the company's real, unrelated industry.
    if reason := _market_rejection(low, market):
        return "manufacturer", f"not in {market}: {reason}"

    if (hit := _mentions(low, _MANUFACTURER_PHRASES)) is not None:
        return "manufacturer", _sentence_around(body, low, hit)[:160]
    if (hit := _mentions(low, _DISTRIBUTOR_PHRASES)) is not None:
        return "distributor", _sentence_around(body, low, hit)[:160]

    # Chinese/Japanese answers, for the ~554 CN/JP companies in this market.
    # Google replies in the company's own language, so the English phrase
    # lists above find nothing and every such company scored "unknown".
    # Manufacturer is checked FIRST here too, matching the English rule that
    # a company which both makes and resells counts as a manufacturer.
    if any(word in body for word in _CJK_MANUFACTURER):
        return "manufacturer", next(w for w in _CJK_MANUFACTURER if w in body)
    if any(word in body for word in _CJK_DISTRIBUTOR):
        return "distributor", next(w for w in _CJK_DISTRIBUTOR if w in body)

    return "unknown", body[:160]


def screen_roles_via_google_ai(
    companies: list[ScreenedCompany], *, market: str, concurrency: int | None = None
) -> list[ScreenedCompany]:
    """Classify each company with its own Google AI Overview query, running
    `concurrency` queries at a time. Mutates and returns the same objects."""
    import asyncio

    from vendor_intel.clients.google_ai_scraper import google_ai_ask

    workers = concurrency or google_ai_concurrency()
    total = len(companies)
    _log(f"screening {total} companies for '{market}' via Google AI, {workers} at a time")

    async def run() -> None:
        sem = asyncio.Semaphore(workers)

        async def one(company: ScreenedCompany) -> None:
            async with sem:
                query = (
                    f"Is {company.name} a distributor (reseller of other brands) "
                    f"or a manufacturer in the {market} market?"
                )
                try:
                    data = await google_ai_ask(query)
                except Exception as exc:  # noqa: BLE001
                    company.verdict, company.reason = "unknown", f"{type(exc).__name__}"
                    return
                # Still blocked after the client's wait-and-retry. This is
                # NOT evidence about the company, so mark it for a later
                # pass instead of burning its one screening chance — the
                # checkpoint records processed companies, so a verdict here
                # would mean it is never asked about again.
                if data.get("captcha_blocked"):
                    company.verdict = "blocked"
                    company.is_distributor = False
                    company.reason = "google bot detection — not screened"
                    return

                verdict, reason = _verdict_from_overview(
                    data.get("markdown") or "", market=market
                )
                company.verdict = verdict
                company.is_distributor = verdict == "distributor"
                company.reason = reason

        done = 0
        for start in range(0, total, workers * 4):
            chunk = companies[start : start + workers * 4]
            await asyncio.gather(*(one(c) for c in chunk))
            done += len(chunk)
            kept = sum(1 for c in companies[:done] if c.is_distributor)
            _detail(f"screened {done}/{total} — {kept} distributors so far")

    asyncio.run(run())
    return companies


def screen_roles(
    companies: list[ScreenedCompany], *, market: str, batch_size: int | None = None
) -> list[ScreenedCompany]:
    """Ask ChatGPT, `batch_size` companies at a time, whether each is a
    distributor for `market`. Mutates and returns the same objects.

    A company keeps verdict="unknown"/is_distributor=False whenever the
    batch fails or the model doesn't answer for it — never defaults to
    accepting."""
    size = batch_size or screen_batch_size()
    total = len(companies)
    _log(f"screening {total} companies for '{market}' role, {size} at a time")

    guidance = market_guidance(market)
    guidance_text = (
        f"\nWhat counts as a distributor here: {guidance['distributor_definition']}\n"
        if guidance.get("distributor_definition")
        else ""
    )

    consecutive_failures = 0
    for start in range(0, total, size):
        batch = companies[start : start + size]
        prompt = _ROLE_PROMPT.format(
            market=market, companies=_format_companies(batch), guidance=guidance_text
        )
        answers = _ask_for_batch(prompt, expected=len(batch), label=f"batch {start // size + 1}")

        # Bail out when the backend is clearly broken rather than grinding
        # through hundreds more batches producing nothing. Observed live: 50
        # consecutive batches failed to parse, so 1,225 companies were all
        # marked "unknown" and the run would have wasted hours before
        # anyone noticed.
        if not answers:
            consecutive_failures += 1
            if consecutive_failures >= _MAX_CONSECUTIVE_BATCH_FAILURES:
                _log(
                    f"ABORTING screening: {consecutive_failures} batches in a row "
                    f"returned nothing usable ({start + len(batch)}/{total} companies "
                    f"reached). The backend is likely rate-limited or blocked — "
                    f"remaining companies stay 'unknown' rather than being "
                    f"silently rejected."
                )
                return companies
            # Back off progressively — if this is throttling, hammering
            # straight into the next batch guarantees another failure.
            backoff = _BATCH_PAUSE_SECONDS * (2**consecutive_failures)
            _detail(f"backing off {backoff:.0f}s before the next batch")
            time.sleep(backoff)
        else:
            consecutive_failures = 0
            time.sleep(_BATCH_PAUSE_SECONDS)

        for i, company in enumerate(batch, start=1):
            item = answers.get(i)
            if not item:
                continue
            verdict = str(item.get("verdict") or "unknown").strip().lower()
            company.verdict = verdict if verdict in (
                "distributor",
                "manufacturer",
                "other",
                "unknown",
            ) else "unknown"
            company.is_distributor = company.verdict == "distributor"
            company.reason = str(item.get("reason") or "").strip()

        done = min(start + size, total)
        kept = sum(1 for c in companies[:done] if c.is_distributor)
        _detail(f"screened {done}/{total} — {kept} distributors so far")

    return companies


_FILL_QUESTIONS = {
    "Website": "official company website URL",
    "Founded": "year the company was founded",
    "Headquarters": "headquarters city and country",
    # Ask for NAMED COUNTRIES, not regions. Asking for "countries or
    # regions" invited the vague answer: live results came back as
    # "Europe" / "Global", which is useless in a spreadsheet cell meant to
    # list where a company actually operates.
    # Phrased as a plain search question, NOT as an instruction with
    # prohibitions. Measured across 3 companies × 5 phrasings: the
    # "list the countries ... do NOT answer 'Global'" wording returned 0
    # countries, while asking about the sales/distribution network returned
    # 7. Google is a search engine — negative instructions read as noise and
    # suppress results rather than shaping them.
    "Continent / Geography": (
        "where are its offices and physical locations? Which countries"
    ),
    "Operational Presence": (
        "what is its international sales and distribution network? Which countries"
    ),
    "Ownership": "ownership type (private, public, family-owned, subsidiary of X)",
    "Employees": "number of employees",
    "Core Categories": "main product categories it distributes",
    "Specialty Focus": "its specialisation or niche",
    "Key Brands Represented": "brands it distributes",
    "Online / Offline / Both": "does it sell online, offline, or both",
    "Contact Person": "name of its head of sales or procurement",
    "Role": "that person's job title",
    "Email": "that person's email address",
    "LinkedIn": "that person's LinkedIn profile URL",
    "Office No.": "main office phone number",
    "Office Email": "general office email address",
}


def fill_fields_via_google_ai(
    companies: list[ScreenedCompany],
    *,
    columns: list[str],
    market: str = "",
    concurrency: int | None = None,
) -> list[ScreenedCompany]:
    """Fill each company's still-empty columns with Google AI Overview.

    Two deliberate differences from the batch version:

    1. ONE query per company. Google is a search engine — a prompt listing
       many companies matches no web page and returns "About 0 results".
    2. Only columns that are still EMPTY for that company are asked about.
       Anything Kompass's listing already supplied (headquarters, employee
       count, phone) is left alone: it's first-party data and shouldn't be
       overwritten by a researched guess, nor re-researched at a cost.
    """
    import asyncio

    from vendor_intel.clients.google_ai_scraper import google_ai_ask

    workers = concurrency or google_ai_concurrency()
    total = len(companies)
    _log(f"filling missing columns for {total} companies via Google AI, {workers} at a time")

    async def run() -> None:
        sem = asyncio.Semaphore(workers)

        async def one(company: ScreenedCompany) -> None:
            gaps = [c for c in columns if not str(company.fields.get(c) or "").strip()]
            if not gaps:
                return
            where = f" in {company.country_hint}" if company.country_hint else ""

            # The two geography columns are asked SEPARATELY. Bundled into
            # one query their answers are indistinguishable, so both columns
            # ended up with the same value — which destroys the whole point
            # of the distinction (offices vs. where it actually sells).
            geo_cols = [c for c in gaps if c in _GEO_COLUMNS]
            other_cols = [c for c in gaps if c not in _GEO_COLUMNS]

            async def ask(question: str, *, raw: bool = False) -> str:
                # Deliberately a NATURAL question, not a "reply as JSON"
                # prompt. Google is a search engine: asked for strict JSON
                # it returns an empty template even for companies it
                # demonstrably knows. Asked plainly it gives a real answer.
                #
                # raw=True returns the UNCLEANED scrape. Both cleaners drop
                # countries: measured on one company the raw text held 9
                # countries, DeepSeek-cleaned 9 but a different set, and the
                # heuristic-only just 1. For list-extraction the raw text is
                # strictly better — regex doesn't care about readability.
                async with sem:
                    try:
                        data = await google_ai_ask(f"{company.name}{where}: {question}?")
                    except Exception:
                        return ""
                if raw:
                    return (
                        str(data.get("markdown_raw") or "").strip()
                        or str(data.get("markdown") or "").strip()
                    )
                return (data.get("markdown") or "").strip()

            if other_cols:
                answer = await ask(", ".join(_FILL_QUESTIONS.get(c, c) for c in other_cols))
                if answer:
                    for key, value in _extract_fields_from_prose(answer, other_cols).items():
                        company.fields[key] = value

            # Ask each geography column with SEVERAL phrasings and union the
            # results. Google AI Overview is non-deterministic: measured on
            # one company, the identical query returned 0-7 countries across
            # six runs (2 of 6 empty), yet the union was 7. A single ask
            # therefore samples a fraction of what Google actually knows,
            # and which fraction is luck. Different phrasings also surface
            # different facts, so varying them widens coverage further.
            for column in geo_cols:
                found: list[str] = []
                for question in _geo_questions_for(column):
                    answer = await ask(question)
                    if not answer:
                        continue
                    for name in _country_names_in(answer, ignore=company.name):
                        if name not in found:
                            found.append(name)
                if found:
                    company.fields[column] = ", ".join(found)

        for start in range(0, total, workers * 4):
            chunk = companies[start : start + workers * 4]
            await asyncio.gather(*(one(c) for c in chunk))
            _detail(f"filled {min(start + len(chunk), total)}/{total}")

    asyncio.run(run())
    return companies


_WEBSITE_RE = re.compile(r"https?://[^\s)\]>,\"']+|(?<![\w@.])(?:www\.)?[a-z0-9-]+\.[a-z]{2,}(?:\.[a-z]{2})?(?![\w@])", re.I)
_YEAR_RE = re.compile(r"\b(1[89]\d{2}|20[0-2]\d)\b")
_EMPLOYEES_RE = re.compile(r"\b(\d[\d,]*(?:\s*[-–]\s*\d[\d,]*)?)\s*(?:\+\s*)?employees\b", re.I)
# "employs approximately 16 people" / "has around 250 staff" — the number
# comes AFTER the verb, so the pattern above cannot see it.
_EMPLOYEES_PROSE_RE = re.compile(
    r"(?:employs|employing|has|with)\s+"
    r"(?:approximately|around|about|roughly|some|over|more\s+than|~)?\s*"
    r"(\d[\d,]*)\s*(?:\+\s*)?(?:people|staff|employees|workers|persons)\b",
    re.I,
)
# "employee count between 10 and 49" / "between 50 and 200 employees"
_EMPLOYEES_RANGE_RE = re.compile(
    r"(?:between|from)\s+(\d[\d,]*)\s+(?:and|to|-|–)\s+(\d[\d,]*)\s*"
    r"(?:\+\s*)?(?:people|staff|employees|workers)?",
    re.I,
)
# Ownership was ASKED for but never extracted — guaranteeing an empty column.
_OWNERSHIP_RE = re.compile(
    r"[^.;\n]{0,120}\b(?:privately\s+(?:held|owned)|private\s+company|private,|"
    r"family[-\s](?:owned|run)|publicly\s+(?:traded|listed)|public\s+company|"
    r"state[-\s]owned|government[-\s]owned|"
    r"(?:a\s+)?subsidiary\s+of|owned\s+by|part\s+of\s+the)\b[^.;\n]{0,120}",
    re.I,
)
_PHONE_RE_FILL = re.compile(r"\+\d[\d\s().-]{7,}\d")
_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
_LINKEDIN_RE = re.compile(r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/[^\s)\]>,\"']+", re.I)

# Domains that are never a company's own site.
_NOT_A_COMPANY_SITE = (
    "google.", "kompass.", "linkedin.", "facebook.", "twitter.", "x.com",
    "instagram.", "youtube.", "wikipedia.", "bloomberg.", "crunchbase.",
    "dnb.com", "europages.", "yelp.", "indeed.",
)


_CATEGORIES_RE = re.compile(
    r"(?:distributes|supplies|sells|offers|provides|"
    r"product\s+categories\s+include|portfolio\s+includes|"
    r"specialis(?:es|ing)\s+in|specializ(?:es|ing)\s+in)\s+"
    r"(?P<items>[^.;\n]{6,200})",
    re.I,
)
_SPECIALTY_RE = re.compile(
    r"(?:specialis(?:es|ing|ed)?\s+in|specializ(?:es|ing|ed)?\s+in|"
    r"focus(?:es|ed)?\s+on|niche\s+(?:is|in)|"
    r"sole\s+authoriz?sed\s+|expert(?:ise)?\s+in)\s*"
    r"(?P<focus>[^.;\n]{4,140})",
    re.I,
)
# "distributor for Kawasumi and Toray brand", "brands like Kawasumi and Toray"
_BRANDS_RE = re.compile(
    r"(?:brands?\s+(?:like|such\s+as|including|:)|"
    r"(?:authoris?zed|official)\s+(?:technical\s+service\s+and\s+)?"
    r"(?:distributor|dealer|reseller|agent)\s+(?:for|of)|"
    r"distributes)\s+"
    r"(?P<brands>(?:[A-Z][\w&.\-]*(?:\s+[A-Z][\w&.\-]*)?"
    r"(?:\s*(?:,|and|&)\s*[A-Z][\w&.\-]*(?:\s+[A-Z][\w&.\-]*)?){0,7}))",
)
# Words that follow a brand list but are not brands.
_NOT_A_BRAND = {
    "the", "and", "its", "it", "they", "this", "these", "their", "a", "an",
    "in", "for", "of", "to", "from", "brand", "brands", "products", "devices",
    "equipment", "systems", "solutions", "turkey", "germany", "italy", "spain",
    "europe", "company", "ltd", "gmbh", "srl", "sti", "inc", "sa", "bv", "nv",
}


def _clean_list_phrase(raw: str) -> str:
    """Tidy a prose list into a short comma-separated cell value."""
    text = re.sub(r"\s+", " ", raw or "").strip(" .,;:")
    # Cut trailing subordinate clauses: "..., primarily for hospitals"
    text = re.split(r",?\s+(?:primarily|mainly|mostly|especially|particularly)\b", text, maxsplit=1)[0]
    text = re.split(r"\s+(?:for|to)\s+(?:hospitals|clinics|customers)\b", text, maxsplit=1)[0]
    text = text.strip(" .,;:")
    return text[:160] if len(text) > 3 else ""


def _trim_clause(raw: str) -> str:
    text = re.sub(r"\s+", " ", raw or "").strip(" .,;:")
    text = re.split(r",?\s+(?:and\s+is|which|where|while)\b", text, maxsplit=1)[0]
    return text.strip(" .,;:")[:120]


def _split_brand_names(raw: str) -> list[str]:
    """Split a matched brand run into individual names, dropping filler."""
    out: list[str] = []
    for part in re.split(r"\s*(?:,|and|&)\s*", raw or ""):
        name = part.strip(" .,;:")
        # Trailing nouns leak in: "Kawasumi and Toray brand hemodialysis"
        while name and name.split()[-1].lower() in _NOT_A_BRAND:
            name = " ".join(name.split()[:-1])
        if not name or name.lower() in _NOT_A_BRAND:
            continue
        if not name[0].isupper() or len(name) < 2 or len(name) > 40:
            continue
        out.append(name)
    return out


def _sales_channel(text: str) -> str:
    """Classify the sales channel into the three values this column holds."""
    low = (text or "").lower()
    has_online = bool(
        re.search(r"\b(?:sells?|sold|available|operates?)\b[^.;\n]{0,40}\bonline\b", low)
        or re.search(r"\be-?commerce\s+(?:shop|store|platform|site)\b", low)
        or "online store" in low
        or "webshop" in low
    )
    # "does NOT operate a direct online e-commerce shop" is the opposite.
    if re.search(r"\b(?:not|no|without)\b[^.;\n]{0,40}\b(?:online|e-?commerce)\b", low):
        has_online = False
    has_offline = bool(
        re.search(r"\b(?:offline|in[- ]person|brick[- ]and[- ]mortar|physical\s+(?:store|shop)|"
                  r"traditional\s+b2b|direct\s+(?:sales|procurement)|showroom|"
                  r"distribution\s+channels?)\b", low)
    )
    if re.search(r"\bboth\s+online\s+and\s+offline\b", low):
        return "Both"
    if has_online and has_offline:
        return "Both"
    if has_online:
        return "Online"
    if has_offline:
        return "Offline"
    return ""


def _extract_fields_from_prose(text: str, gaps: list[str]) -> dict[str, str]:
    """Pull requested fields out of a plain-language AI Overview answer.

    Only patterns we can recognise with confidence are extracted; anything
    else is left for a later enrichment stage rather than guessed at from
    free text.
    """
    out: dict[str, str] = {}

    if "Website" in gaps:
        for m in _WEBSITE_RE.finditer(text):
            url = m.group(0).rstrip(".,);:")
            low = url.lower()
            if any(bad in low for bad in _NOT_A_COMPANY_SITE):
                continue
            if "." not in low.split("/")[0]:
                continue
            out["Website"] = url if low.startswith("http") else f"https://{url}"
            break

    if "Founded" in gaps:
        # Prefer a year stated near founding language.
        near = re.search(
            r"(?:founded|established|since|incorporated)[^.]{0,40}?\b(1[89]\d{2}|20[0-2]\d)\b",
            text,
            re.I,
        )
        if near:
            out["Founded"] = near.group(1)

    if "Employees" in gaps:
        # Try the strict "N employees" form first, then the prose forms
        # Google actually uses. Measured live: "employs approximately 16
        # people" and "employee count between 10 and 49" both parsed as
        # nothing, because the old pattern required the literal word
        # "employees" immediately after the number.
        if m := _EMPLOYEES_RE.search(text):
            out["Employees"] = m.group(1).strip()
        elif m := _EMPLOYEES_PROSE_RE.search(text):
            out["Employees"] = m.group(1).strip()
        elif m := _EMPLOYEES_RANGE_RE.search(text):
            lo, hi = m.group(1).strip(), m.group(2).strip()
            out["Employees"] = f"{lo}-{hi}"

    if "Ownership" in gaps and (m := _OWNERSHIP_RE.search(text)):
        # Normalise to the handful of values this column is meant to hold,
        # so the spreadsheet stays sortable instead of carrying a sentence.
        raw = m.group(0).lower()
        parts: list[str] = []
        for label, needles in (
            ("Family-owned", ("family-owned", "family owned", "family-run")),
            ("Private", ("privately held", "privately owned", "private company", "private,")),
            ("Public", ("publicly traded", "publicly listed", "public company")),
            ("Subsidiary", ("subsidiary", "owned by", "part of")),
            ("State-owned", ("state-owned", "government-owned")),
        ):
            if any(n in raw for n in needles) and label not in parts:
                parts.append(label)
        # "not listed on any public stock exchange" is a statement that the
        # company is private, not that it is public.
        if "Public" in parts and re.search(r"not\s+(?:listed|publicly)", text, re.I):
            parts.remove("Public")
        if parts:
            out["Ownership"] = ", ".join(parts)

    if "Office No." in gaps and (m := _PHONE_RE_FILL.search(text)):
        out["Office No."] = m.group(0).strip()

    if (m := _EMAIL_RE.search(text)):
        addr = m.group(0)
        # info@/sales@ style addresses are the office, not a named person.
        generic = addr.split("@")[0].lower() in (
            "info", "sales", "contact", "office", "hello", "enquiries", "support",
        )
        if generic and "Office Email" in gaps:
            out["Office Email"] = addr
        elif not generic and "Email" in gaps:
            out["Email"] = addr

    if "LinkedIn" in gaps and (m := _LINKEDIN_RE.search(text)):
        out["LinkedIn"] = m.group(0).rstrip(".,);")

    # --- Core Categories / Specialty Focus / Key Brands / sales channel ---
    # All four were ASKED on every company and none had any extraction, so
    # they sat at exactly 0% no matter how good the answer was. Measured on
    # a live answer, Google returns these richly:
    #   "distributes medical equipment, consumables, and water treatment
    #    systems"                                  -> Core Categories
    #   "specializing in hemodialysis equipment"   -> Specialty Focus
    #   "distributor for Kawasumi and Toray"       -> Key Brands
    #   "sells offline through traditional B2B..." -> Online / Offline / Both

    if "Core Categories" in gaps and (m := _CATEGORIES_RE.search(text)):
        cats = _clean_list_phrase(m.group("items"))
        if cats:
            out["Core Categories"] = cats

    if "Specialty Focus" in gaps and (m := _SPECIALTY_RE.search(text)):
        focus = _trim_clause(m.group("focus"))
        if 3 < len(focus) <= 120:
            out["Specialty Focus"] = focus

    if "Key Brands Represented" in gaps:
        brands: list[str] = []
        for bm in _BRANDS_RE.finditer(text):
            for name in _split_brand_names(bm.group("brands")):
                if name not in brands:
                    brands.append(name)
        if brands:
            out["Key Brands Represented"] = ", ".join(brands[:8])

    if "Online / Offline / Both" in gaps and (channel := _sales_channel(text)):
        out["Online / Offline / Both"] = channel

    # Headquarters. Also previously unhandled: it sat at 1.5%/405 not
    # because Google failed to answer but because nothing read the answer.
    if "Headquarters" in gaps and (m := _HQ_RE.search(text)):
        hq = m.group("place").strip(" .;,")
        # Keep "City, Country"; drop trailing clauses like "Italy, where it".
        hq = re.split(r"\s+(?:where|which|and)\b", hq)[0].strip(" .;,")
        if 2 < len(hq) <= 60:
            out["Headquarters"] = hq

    # Contact Person / Role. Previously unhandled entirely, which is why
    # both sat at exactly 0/405 — the answer could never be parsed, no
    # matter what Google returned. Only sales/procurement leadership counts
    # (per the standing rule), so a CEO or founder is deliberately skipped.
    if "Contact Person" in gaps or "Role" in gaps:
        from vendor_intel.pipeline.company_enrichment import is_accepted_contact_role

        for m in _PERSON_ROLE_RE.finditer(text):
            # Two alternation branches ("Name, Role" and "Role Name"), so
            # take whichever pair matched.
            person = (
                m.group("person")
                or m.group("person2")
                or m.group("person3")
                or m.group("person4")
                or ""
            ).strip()
            role = (
                m.group("role")
                or m.group("role2")
                or m.group("role3")
                or m.group("role4")
                or ""
            ).strip()
            if not person or not role or not is_accepted_contact_role(role):
                continue
            if "Contact Person" in gaps:
                out["Contact Person"] = person
            if "Role" in gaps:
                out["Role"] = role
            break

    return out


# "Jane Smith, Head of Sales" / "Jane Smith is the Sales Director" /
# "Sales Manager Jane Smith". Names are Capitalised words; the role has to
# survive is_accepted_contact_role(), so a loose match here is safe.
# "headquartered in X" / "based in X" / "Headquarters: X" — the phrasings
# Google actually returns for a headquarters question.
_HQ_RE = re.compile(
    r"(?:head\s?quarter(?:ed|s)?|headquarters|based)\s*"
    r"(?::|\s+(?:is|are|was|were)?\s*(?:located\s+)?in\s+)\s*"
    r"(?P<place>[A-Z][^.;\n]{2,60})",
    re.I,
)

# Role titles that pass the sales/procurement gate. Anchored to whole
# words so the captured Role is the title itself, not a run-on phrase
# ("Head of Sales is", "Sales Manager Anna Bianchi").
_ROLE_WORDS = (
    r"(?:Chief\s+(?:Sales|Procurement|Purchasing)\s+Officer"
    r"|(?:VP|Vice\s+President|Head|Director|Manager)\s+of\s+"
    r"(?:Sales|Procurement|Purchasing)"
    r"|(?:Sales|Procurement|Purchasing)\s+"
    r"(?:VP|Vice\s+President|Director|Manager|Head|Lead))"
)
# A person is 2-3 capitalised words. Requiring capitals keeps sentence
# fragments ("handles key accounts") out of the name slot.
_LEADIN = r"(?:Contact|Email|Call|Reach|Ask|See)\s+"
_PERSON = r"(?-i:[A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,2})"
_PERSON_ROLE_RE = re.compile(
    # Case-insensitive: live answers phrase the title in lower case
    # mid-sentence ("the overall head of sales is Corrado Damascati"),
    # which a capitals-only pattern silently missed.
    # "The Head of Sales is Marco Rossi" - checked FIRST so the role is not
    # swallowed into the name by a later branch.
    rf"(?:the\s+)?(?:\w+\s+){{0,2}}?(?P<role3>{_ROLE_WORDS})\s+is\s+(?P<person3>{_PERSON})"
    # "Marco Rossi, Sales Director" / "Marco Rossi serves as Head of Sales"
    rf"|(?:{_LEADIN})?(?P<person>{_PERSON})\s*"
    rf"(?:,\s*|\s+(?:is|serves)\s+(?:the\s+|as\s+)?|\s+[—-]\s*)"
    rf"(?P<role>{_ROLE_WORDS})"
    # "Sales Director Marco Rossi"
    rf"|(?P<role2>{_ROLE_WORDS})\s+(?P<person2>{_PERSON})"
    # "Luisa Marini is the Direttore Commerciale (Head of Sales / ...)"
    # Live answers often give the local-language title first and the
    # English one in parentheses, so the role is NOT adjacent to the name.
    # Allow intervening text up to the qualifying title on the same line.
    rf"|(?P<person4>{_PERSON})\s+(?:is|serves\s+as)\s+[^.;\n]{{0,80}}?"
    rf"(?P<role4>{_ROLE_WORDS})",
    re.I,
)


# Enough coverage for the markets this pipeline targets; unmatched names are
# simply not extracted rather than guessed at.
_COUNTRIES = (
    "Afghanistan", "Albania", "Algeria", "Argentina", "Australia", "Austria",
    "Bahrain", "Bangladesh", "Belarus", "Belgium", "Bolivia", "Bosnia",
    "Brazil", "Bulgaria", "Cambodia", "Canada", "Chile", "China", "Colombia",
    "Costa Rica", "Croatia", "Cyprus", "Czech Republic", "Czechia", "Denmark",
    "Dominican Republic", "Ecuador", "Egypt", "Estonia", "Ethiopia", "Finland",
    "France", "Georgia", "Germany", "Ghana", "Greece", "Guatemala", "Hungary",
    "Iceland", "India", "Indonesia", "Iran", "Iraq", "Ireland", "Israel",
    "Italy", "Japan", "Jordan", "Kazakhstan", "Kenya", "Kuwait", "Latvia",
    "Lebanon", "Libya", "Lithuania", "Luxembourg", "Malaysia", "Malta",
    "Mexico", "Moldova", "Morocco", "Netherlands", "New Zealand", "Nigeria",
    "North Macedonia", "Norway", "Oman", "Pakistan", "Panama", "Paraguay",
    "Peru", "Philippines", "Poland", "Portugal", "Qatar", "Romania", "Russia",
    "Saudi Arabia", "Senegal", "Serbia", "Singapore", "Slovakia", "Slovenia",
    "South Africa", "South Korea", "Spain", "Sri Lanka", "Sweden",
    "Switzerland", "Taiwan", "Tanzania", "Thailand", "Tunisia", "Turkey",
    "Uganda", "Ukraine", "United Arab Emirates", "United Kingdom",
    "United States", "Uruguay", "Uzbekistan", "Venezuela", "Vietnam",
    # Common alternates, normalised below.
    "UAE", "UK", "USA", "U.S.", "US", "Korea", "Holland", "England",
)
_COUNTRY_ALIASES = {
    "uae": "United Arab Emirates", "uk": "United Kingdom",
    "usa": "United States", "u.s.": "United States", "us": "United States",
    "korea": "South Korea", "holland": "Netherlands",
    "england": "United Kingdom", "czechia": "Czech Republic",
}


# Country names in Chinese and Japanese, mapped to the SAME canonical
# English name the Latin table produces — the export column must read
# "Indonesia" whether Google answered in English or Chinese.
#
# Needed because Kompass's country sweep includes China (326 candidates) and
# Japan (228), and Google answers those companies in their own language. The
# Latin matcher found nothing in such answers, so a live Chinese reply that
# named ten countries and "105 countries and regions" extracted zero.
#
# Matched by plain substring, NOT \b word boundaries: CJK text has no spaces,
# so \b never fires between two Han characters.
_CJK_COUNTRIES: dict[str, str] = {
    # East Asia
    "中国": "China", "中國": "China", "日本": "Japan", "韓国": "South Korea",
    "韩国": "South Korea", "韓國": "South Korea", "台湾": "Taiwan", "台灣": "Taiwan",
    "香港": "Hong Kong", "蒙古": "Mongolia", "朝鲜": "North Korea",
    # South-East / South Asia
    "印度尼西亚": "Indonesia", "インドネシア": "Indonesia",
    "马来西亚": "Malaysia", "マレーシア": "Malaysia",
    "越南": "Vietnam", "ベトナム": "Vietnam",
    "菲律宾": "Philippines", "フィリピン": "Philippines",
    "泰国": "Thailand", "タイ": "Thailand",
    "新加坡": "Singapore", "シンガポール": "Singapore",
    "孟加拉国": "Bangladesh", "巴基斯坦": "Pakistan",
    "印度": "India", "インド": "India", "斯里兰卡": "Sri Lanka",
    "缅甸": "Myanmar", "柬埔寨": "Cambodia",
    # Central Asia / Middle East
    "土库曼斯坦": "Turkmenistan", "哈萨克斯坦": "Kazakhstan",
    "乌兹别克斯坦": "Uzbekistan", "伊拉克": "Iraq", "伊朗": "Iran",
    "沙特阿拉伯": "Saudi Arabia", "阿联酋": "United Arab Emirates",
    "土耳其": "Turkey", "トルコ": "Turkey", "以色列": "Israel",
    # Europe
    "德国": "Germany", "ドイツ": "Germany", "法国": "France", "フランス": "France",
    "英国": "United Kingdom", "イギリス": "United Kingdom",
    "意大利": "Italy", "イタリア": "Italy", "西班牙": "Spain", "スペイン": "Spain",
    "荷兰": "Netherlands", "オランダ": "Netherlands",
    "比利时": "Belgium", "瑞士": "Switzerland", "スイス": "Switzerland",
    "瑞典": "Sweden", "奥地利": "Austria", "波兰": "Poland",
    "俄罗斯": "Russia", "ロシア": "Russia", "葡萄牙": "Portugal",
    "希腊": "Greece", "丹麦": "Denmark", "芬兰": "Finland", "挪威": "Norway",
    "捷克": "Czech Republic", "匈牙利": "Hungary", "罗马尼亚": "Romania",
    "乌克兰": "Ukraine",
    # Americas / Africa / Oceania
    "美国": "United States", "アメリカ": "United States", "米国": "United States",
    "加拿大": "Canada", "カナダ": "Canada", "墨西哥": "Mexico",
    "巴西": "Brazil", "ブラジル": "Brazil", "阿根廷": "Argentina",
    "智利": "Chile", "哥伦比亚": "Colombia", "秘鲁": "Peru",
    "南非": "South Africa", "埃及": "Egypt", "尼日利亚": "Nigeria",
    "肯尼亚": "Kenya", "赞比亚": "Zambia", "科特迪瓦": "Ivory Coast",
    "摩洛哥": "Morocco", "澳大利亚": "Australia", "オーストラリア": "Australia",
    "新西兰": "New Zealand",
}

# Role words in Chinese and Japanese, so a CJK answer can be classified.
_CJK_DISTRIBUTOR = (
    "经销商", "經銷商", "代理商", "分销商", "分銷商", "批发商", "批發商",
    "贸易商", "貿易商", "进出口公司", "進出口公司", "销售代理", "販売代理店",
    "代理店", "卸売業者", "ディストリビューター", "유통업체", "대리점",
)
_CJK_MANUFACTURER = (
    "制造商", "製造商", "生产商", "生產商", "厂商", "廠商", "制造企业",
    "生产企业", "メーカー", "製造業者", "製造元", "제조업체", "제조사",
)


def _cjk_countries_in(text: str) -> list[tuple[int, str]]:
    """(position, canonical English name) for CJK country names in `text`.

    Longest-first so 印度尼西亚 (Indonesia) is matched before the 印度
    (India) prefix inside it — otherwise every Indonesia mention would also
    record a spurious India.
    """
    hits: list[tuple[int, str]] = []
    claimed: list[tuple[int, int]] = []
    for cjk in sorted(_CJK_COUNTRIES, key=len, reverse=True):
        start = 0
        while (idx := text.find(cjk, start)) >= 0:
            end = idx + len(cjk)
            if not any(s <= idx < e for s, e in claimed):
                hits.append((idx, _CJK_COUNTRIES[cjk]))
                claimed.append((idx, end))
            start = end
    return hits


_GEO_COLUMNS = ("Continent / Geography", "Operational Presence")

# Several phrasings per column, unioned. Two reasons, both measured:
#   - Google AI Overview is non-deterministic: the SAME query returned 0-7
#     countries across six runs of one company, 2 of them empty.
#   - Phrasings surface different facts: across 3 companies, "international
#     sales and distribution network" found 7 countries where "list the
#     countries ... do NOT answer 'Global'" found 0. Negative instructions
#     read as noise to a search engine and suppress results.
_GEO_QUESTIONS = {
    "Continent / Geography": (
        "where are its offices and physical locations? Which countries",
    ),
    "Operational Presence": (
        # Ranked by measurement across 3 companies (total countries found):
        #   supply 24 | markets 11 | export_to 8 | name_every 5 | network 3
        # "network" was the previous default and scored near-worst — it
        # tends to draw a summary ("spans over 50 countries") rather than a
        # list, which contains no extractable country name.
        "which countries does it supply customers in",
    ),
}
# Extra phrasings, unioned with the above when GOOGLE_AI_GEO_DEEP=true.
# Each one is another AI Overview call per company: measured, using all of
# them takes a 9,460-company market from ~12h to ~23h. Worth it only when
# geography coverage matters more than turnaround.
_GEO_QUESTIONS_DEEP = {
    "Continent / Geography": (
        "which countries does it have subsidiaries or branches in",
        "list every country where it has an office, branch or subsidiary",
    ),
    "Operational Presence": (
        # In measured order after the default. No single phrasing wins
        # everywhere — "supply" found 18 countries for one company and 0
        # for another, where "export_to" found 6 — so the union across
        # phrasings is what actually produces full lists.
        "which international markets and countries does it operate in",
        "which countries does it export to or sell its products in",
        "name every country in its distribution network, list them all",
    ),
}


def geo_deep_enabled() -> bool:
    return (os.getenv("GOOGLE_AI_GEO_DEEP") or "").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )


def _geo_questions_for(column: str) -> tuple[str, ...]:
    questions = _GEO_QUESTIONS.get(column, ())
    if geo_deep_enabled():
        questions = questions + _GEO_QUESTIONS_DEEP.get(column, ())
    return questions


def _country_names_in(text: str, *, ignore: str = "") -> list[str]:
    """Country names appearing in `text`, de-duplicated, in first-seen order.

    Word-boundary matched so "Chile" doesn't fire inside another word, and
    longest-first so "United States" wins over the bare "US" inside it.

    `ignore` strips a string (normally the company name) from the text
    first: answers routinely echo the company name back, and a name like
    "JORDAN GAMMA MEDIZINTECHNIK GMBH" was being read as the country Jordan
    — recording a country the company has no connection to.
    """
    if ignore.strip():
        text = re.sub(re.escape(ignore.strip()), " ", text, flags=re.I)
    low = text.lower()
    hits: list[tuple[int, str]] = []
    for name in sorted(_COUNTRIES, key=len, reverse=True):
        for m in re.finditer(rf"\b{re.escape(name.lower())}\b", low):
            hits.append((m.start(), _COUNTRY_ALIASES.get(name.lower(), name)))
            break

    # Chinese/Japanese names resolve to the same canonical English value, so
    # a CJK answer fills the column identically to an English one.
    hits.extend(_cjk_countries_in(text))

    seen: set[str] = set()
    out: list[str] = []
    for _, name in sorted(hits):
        if name not in seen:
            seen.add(name)
            out.append(name)
    return out


def _parse_json_object(reply: str) -> dict[str, Any]:
    """Pull the last valid JSON object out of a reply, mirroring
    _parse_json_array's reasoning about echoed prompt templates."""
    text = (reply or "").strip()
    if text.startswith("```"):
        text = text.strip("`").removeprefix("json").strip()
    best: dict[str, Any] = {}
    for start, ch in enumerate(text):
        if ch != "{":
            continue
        end = text.rfind("}")
        while end > start:
            try:
                data = json.loads(text[start : end + 1])
            except (json.JSONDecodeError, ValueError):
                end = text.rfind("}", start, end)
                continue
            if isinstance(data, dict) and data:
                best = data
            break
    return best


def fill_fields(
    companies: list[ScreenedCompany],
    *,
    columns: list[str],
    market: str = "",
    batch_size: int | None = None,
) -> list[ScreenedCompany]:
    """Ask ChatGPT to fill `columns` for each company, in batches. Values
    the model doesn't know come back empty and are simply not set, so a
    later enrichment stage can still try."""
    if not companies or not columns:
        return companies

    size = batch_size or screen_batch_size()
    total = len(companies)
    _log(f"filling {len(columns)} columns for {total} companies, {size} at a time")

    guidance = market_guidance(market) if market else {}
    hints = [guidance[k] for k in ("category_hint", "brands_hint") if guidance.get(k)]
    guidance_text = ("\n" + "\n".join(hints) + "\n") if hints else ""

    consecutive_failures = 0
    for start in range(0, total, size):
        batch = companies[start : start + size]
        prompt = _FIELDS_PROMPT.format(
            companies=_format_companies(batch),
            fields=", ".join(columns),
            guidance=guidance_text,
        )
        answers = _ask_for_batch(prompt, expected=len(batch), label=f"batch {start // size + 1}")
        if not answers:
            consecutive_failures += 1
            if consecutive_failures >= _MAX_CONSECUTIVE_BATCH_FAILURES:
                _log(
                    f"ABORTING field fill: {consecutive_failures} batches in a row "
                    f"returned nothing usable. Remaining fields stay empty for the "
                    f"later enrichment stages to fill."
                )
                return companies
            time.sleep(_BATCH_PAUSE_SECONDS * (2**consecutive_failures))
        else:
            consecutive_failures = 0
            time.sleep(_BATCH_PAUSE_SECONDS)

        for i, company in enumerate(batch, start=1):
            fields = (answers.get(i) or {}).get("fields") or {}
            if not isinstance(fields, dict):
                continue
            for key, value in fields.items():
                text = str(value or "").strip()
                if text and key in columns:
                    company.fields[key] = text

        _detail(f"filled {min(start + size, total)}/{total}")

    return companies
