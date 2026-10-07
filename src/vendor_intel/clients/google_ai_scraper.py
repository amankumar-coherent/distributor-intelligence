"""Google AI Overview scraper client (HTTP relay on :15551).

Requires the google-ai-scraper FastAPI backend + browser extension:
  cd google-ai-scraper-main/google-ai-scraper-main/server
  uv run google-ai-scraper --backend --port 15551

API: GET /ask?q=... → {markdown, citations, query_id, thread_id, error?}
     GET /health

DeepSeek cleanup (GOOGLE_AI_SCRAPER_LLM_CLEAN=true, OPENAI_* → api.deepseek.com):
  Google AI Overview often returns a long noisy page. DeepSeek keeps only the
  useful sentences that answer the column query (e.g. founded year of X).

When Google skips AI Overview, the client retries with rephrased queries
(GOOGLE_AI_SCRAPER_FORCE_OVERVIEW=true, default on) instead of accepting empty SERP chrome.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import time
from dataclasses import dataclass
from urllib.parse import quote_plus

import httpx

_health_cache: tuple[float, bool] = (0.0, False)
_HEALTH_TTL_SEC = 30.0

_NAV_LINE_RE = re.compile(
    r"^(skip to|accessibility|sign in|google apps|search results|ai mode|all\b|"
    r"images\b|videos\b|news\b|shopping\b|forums\b|more\b|web results|"
    r"people also ask|short videos|show all|show more|tools\b|"
    r"filter|settings|privacy|terms)\b",
    re.I,
)

_OVERVIEW_MISSING_RE = re.compile(
    r"(ai overview is not available|can't generate an ai overview|"
    r"unable to generate an ai overview|no ai overview|"
    r"an ai overview is not available)",
    re.I,
)


@dataclass
class GoogleAiHit:
    title: str
    link: str
    snippet: str


def google_ai_base_url() -> str:
    return (
        os.getenv("GOOGLE_AI_SCRAPER_URL")
        or os.getenv("GOOGLE_AI_SCRAPER_BASE_URL")
        or "http://127.0.0.1:15551"
    ).strip().rstrip("/")


def google_ai_enabled() -> bool:
    return (os.getenv("GOOGLE_AI_SCRAPER_ENABLED") or "").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )


def google_ai_discovery_enabled() -> bool:
    if not google_ai_enabled():
        return False
    raw = (os.getenv("GOOGLE_AI_SCRAPER_DISCOVERY") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def google_ai_gap_fill_enabled() -> bool:
    if not google_ai_enabled():
        return False
    raw = (os.getenv("GOOGLE_AI_SCRAPER_GAP_FILL") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def google_ai_llm_clean_enabled() -> bool:
    """Allow DeepSeek (OPENAI_*) to turn long Overview text into useful sentences."""
    if not google_ai_enabled():
        return False
    raw = (os.getenv("GOOGLE_AI_SCRAPER_LLM_CLEAN") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def google_ai_llm_clean_only_when_missing() -> bool:
    """If true, DeepSeek clean runs only when AI Overview is missing.

    Default false: long Overview text is always DeepSeek-cleaned into useful sentences.
    """
    raw = (os.getenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_ONLY_WHEN_MISSING") or "false").strip().lower()
    return raw in ("1", "true", "yes", "on")


def _llm_clean_min_chars() -> int:
    """Below this many characters, skip the DeepSeek cleanup entirely.

    Measured on real traffic: many Overview answers are already short
    (a few hundred chars) and the "cleaned" version is barely different,
    yet each call still costs input tokens for the whole text. Since every
    downstream consumer parses with regex rather than reading prose, that
    spend buys nothing. 900 keeps genuinely long, nav-chrome-heavy pages
    going through the LLM, where it does help.
    """
    try:
        return max(0, int(os.getenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MIN_CHARS", "900")))
    except ValueError:
        return 900


def _llm_clean_max_input_chars() -> int:
    """Cap on how much text is sent to the cleanup LLM (see caller)."""
    try:
        return max(500, int(os.getenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_MAX_CHARS", "4000")))
    except ValueError:
        return 4000


def _llm_clean_model() -> str:
    """Prefer DeepSeek chat when OPENAI_* points at api.deepseek.com."""
    explicit = (os.getenv("GOOGLE_AI_SCRAPER_LLM_MODEL") or "").strip()
    if explicit:
        return explicit
    model = (os.getenv("OPENAI_MODEL") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "").strip().lower()
    if "deepseek" in base or model.startswith("deepseek"):
        return model or "deepseek-chat"
    return model or "deepseek-chat"


def google_ai_force_overview_enabled() -> bool:
    """Retry / rephrase when Google returns no AI Overview (default on)."""
    raw = (os.getenv("GOOGLE_AI_SCRAPER_FORCE_OVERVIEW") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def _overview_retries() -> int:
    try:
        return max(1, min(8, int(os.getenv("GOOGLE_AI_SCRAPER_OVERVIEW_RETRIES") or "4")))
    except ValueError:
        return 4


def _overview_retry_sleep_sec() -> float:
    try:
        return max(0.5, float(os.getenv("GOOGLE_AI_SCRAPER_OVERVIEW_RETRY_SLEEP_SEC") or "2.5"))
    except ValueError:
        return 2.5


def _captcha_wait_sec() -> float:
    """How long to wait for a CAPTCHA to clear before retrying the query.

    A block is temporary: Google lifts it once the burst of requests stops,
    or the user solves the challenge in the visible tab. Discarding the
    company instead of waiting throws away work that a short pause
    recovers.
    """
    try:
        return max(0.0, float(os.getenv("GOOGLE_AI_CAPTCHA_WAIT_SEC") or "30"))
    except ValueError:
        return 30.0


# A cheap, always-answerable query used only to test whether the block has
# lifted. Deliberately generic so it never depends on a company being known.
_CAPTCHA_PROBE_QUERY = "capital of France"


def _captcha_probe_interval() -> float:
    """Seconds between checks while waiting for a CAPTCHA to be solved.

    Each probe is a real Google query, so this should not be aggressive —
    15s is slow enough not to look like traffic while the block is active,
    and fast enough that a solved challenge is noticed promptly.
    """
    try:
        return max(0.01, float(os.getenv("GOOGLE_AI_CAPTCHA_PROBE_SEC") or "15"))
    except ValueError:
        return 15.0


def _captcha_probe_timeout() -> float:
    """Give up holding the queue after this long (default 30 min).

    Without a ceiling an unattended overnight run would stall forever on a
    challenge nobody is there to solve.
    """
    try:
        return max(60.0, float(os.getenv("GOOGLE_AI_CAPTCHA_MAX_WAIT_SEC") or "1800"))
    except ValueError:
        return 1800.0


def _captcha_wait_for_solve() -> bool:
    """Hold every tab until the CAPTCHA clears, instead of a fixed sleep."""
    raw = (os.getenv("GOOGLE_AI_CAPTCHA_WAIT_FOR_SOLVE") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def _captcha_retries() -> int:
    """How many times to wait-and-retry a blocked query before giving up."""
    try:
        return max(0, min(10, int(os.getenv("GOOGLE_AI_CAPTCHA_RETRIES") or "3")))
    except ValueError:
        return 3


# Text that only ever appears on Google's bot-block / CAPTCHA interstitial.
_CAPTCHA_MARKERS: tuple[str, ...] = (
    "unusual traffic from your computer network",
    "our systems have detected unusual traffic",
    "solving the above captcha",
    "/sorry/index",
    "recaptcha",
    "why did this happen",
)


class _CaptchaGate:
    """Makes every in-flight query wait out a block that any one of them hit.

    Queries run ~10 at a time, each in its own tab. Google's block is
    per-network, not per-tab, so when one query is blocked the other nine
    are too. Without a shared gate each would independently discover the
    block, wait, and retry — hammering Google while it is actively
    rate-limiting, which prolongs the block and burns the retry budget of
    every company in flight.

    One waiter does the sleeping; the rest simply wait for the same event.
    """

    def __init__(self) -> None:
        self._until: float = 0.0
        self._lock: "asyncio.Lock | None" = None

    def _get_lock(self) -> "asyncio.Lock":
        # Created lazily: the module is imported before any event loop runs.
        if self._lock is None:
            self._lock = asyncio.Lock()
        return self._lock

    async def wait_if_blocked(self) -> None:
        loop = asyncio.get_event_loop()
        remaining = self._until - loop.time()
        if remaining > 0:
            await asyncio.sleep(remaining)

    async def block_for(self, seconds: float) -> None:
        """Record a block and sleep it out, coalescing concurrent callers."""
        async with self._get_lock():
            loop = asyncio.get_event_loop()
            now = loop.time()
            # Extend rather than restart, so ten callers hitting the same
            # block do not serialise into ten consecutive waits.
            self._until = max(self._until, now + seconds)
            remaining = self._until - now
        await asyncio.sleep(max(0.0, remaining))

    async def hold_until_clear(self, base: str, mode: str, probe) -> None:
        """Freeze ALL querying until Google actually answers again.

        `block_for` only sleeps a fixed interval and then lets every tab
        resume regardless of whether the challenge was solved — so a
        still-blocked Google gets hit by ten fresh tabs, which prolongs the
        block. This instead polls with ONE cheap query and holds the rest of
        the queue closed until that query comes back clean, which is what
        happens when the CAPTCHA is solved in the visible tab.

        Exactly one caller does the probing; the others park on the same
        barrier, so no extra tabs are opened while the block is active.
        """
        loop = asyncio.get_event_loop()
        async with self._get_lock():
            already_probing = self._until > loop.time()
            if not already_probing:
                # Claim the prober role. _until doubles as "a hold is
                # active"; it is cleared the moment the block lifts.
                self._until = loop.time() + _captcha_probe_timeout()

        if already_probing:
            # Someone else is probing. Poll the flag in short steps rather
            # than sleeping the whole budget: the earlier version slept the
            # full 30-minute timeout even after the prober had already
            # logged "CAPTCHA cleared", so every other tab stayed parked and
            # the run went silent for half an hour.
            while True:
                async with self._get_lock():
                    if self._until <= loop.time():
                        return
                await asyncio.sleep(1.0)

        interval = _captcha_probe_interval()
        deadline = asyncio.get_event_loop().time() + _captcha_probe_timeout()
        attempt = 0
        while asyncio.get_event_loop().time() < deadline:
            attempt += 1
            print(
                f"  [google_ai] CAPTCHA — no new tabs; waiting for it to be "
                f"solved (check {attempt}, every {interval:.0f}s)",
                flush=True,
            )
            await asyncio.sleep(interval)
            try:
                data = await probe(base, _CAPTCHA_PROBE_QUERY, thread_id=None, mode=mode)
            except Exception:
                continue
            if not is_captcha_page(data):
                print("  [google_ai] CAPTCHA cleared — resuming", flush=True)
                async with self._get_lock():
                    self._until = 0.0
                return

        print(
            "  [google_ai] CAPTCHA still present after the wait budget — "
            "continuing anyway",
            flush=True,
        )
        async with self._get_lock():
            self._until = 0.0


_captcha_gate = _CaptchaGate()


def is_captcha_page(data: dict) -> bool:
    """True when the scrape returned a bot-block page instead of results.

    Detected on the RAW text: the block page is prose, so the cleaners
    happily reformat it into something that reads like a real answer.
    """
    raw = str(data.get("markdown_raw") or data.get("markdown") or "").lower()
    if not raw:
        return False
    return any(marker in raw for marker in _CAPTCHA_MARKERS)


def _timeout() -> float:
    try:
        return max(30.0, float(os.getenv("GOOGLE_AI_SCRAPER_TIMEOUT_SEC", "120")))
    except ValueError:
        return 120.0


def _ask_mode() -> str:
    mode = (os.getenv("GOOGLE_AI_SCRAPER_MODE") or "pro").strip().lower()
    return mode if mode in ("fast", "pro") else "pro"


def heuristic_clean_scrape_markdown(raw: str, *, max_chars: int = 6000) -> str:
    """Cheap local cleanup before (or instead of) LLM — drops nav chrome lines."""
    text = (raw or "").replace("\x00", " ")
    text = re.sub(r"[^\S\n]{2,}", " ", text)
    lines: list[str] = []
    for line in text.splitlines():
        s = line.strip()
        if not s:
            if lines and lines[-1] != "":
                lines.append("")
            continue
        if _NAV_LINE_RE.match(s):
            continue
        if _OVERVIEW_MISSING_RE.search(s):
            continue
        if len(s) <= 2 and not s.isalnum():
            continue
        if s.lower() in {"ai overview", "ai mode", "pro", "wikipedia", "read more"}:
            continue
        lines.append(s)
    out = "\n".join(lines).strip()
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out[:max_chars]


def _ai_overview_missing(data: dict) -> bool:
    """True when scrape has no usable Google AI Overview panel."""
    err = str(data.get("error") or "").strip().lower()
    if err.startswith("no_ai_overview") or err in {
        "empty_ai_overview_extraction",
        "no_ai_overview_after_follow_up",
        "tab_timeout",
    }:
        return True
    raw = str(data.get("markdown_raw") or data.get("markdown") or "")
    if not raw.strip():
        return True
    if _OVERVIEW_MISSING_RE.search(raw):
        return True
    # SERP chrome with almost no substance
    heur = heuristic_clean_scrape_markdown(raw)
    if len(heur) < 80:
        return True
    chrome_hits = sum(
        1
        for token in (
            "advanced search",
            "about ",
            "results",
            "past hour",
            "verbatim",
            "press",
            "jump to the search box",
        )
        if token in raw.lower()
    )
    if chrome_hits >= 3 and len(heur) < 400:
        return True
    return False


def _overview_query_variants(query: str) -> list[str]:
    """Phrasings that more reliably trigger Google AI Overview panels.

    Field-specific asks (``founded year of {company}``) stay as-is first; only
    light fallbacks are added so we do not dilute the fact query.
    """
    q = (query or "").strip()
    if not q:
        return []
    field_markers = (
        "founded year",
        "headquarters",
        "number of employees",
        "official website",
        "operational presence",
        "continent and geography",
        "core categories",
        "specialty focus",
        "key brands represented",
        "distribution type",
        "sales or procurement",
        "office phone",
        "company summary",
        "ownership",
        " of ",
        "ceo founder",
        "brands products",
        "ownership parent",
        "offices locations",
        "revenue turnover",
    )
    low = q.lower()
    if any(m in low for m in field_markers):
        variants = [
            q,
            f"{q} company",
            f"{q} wikipedia",
        ]
    else:
        variants = [
            q,
            f"{q} key facts overview",
            f"Summarize: {q}",
            f"{q}? Give founding year, headquarters, and what the company does.",
            f"What are the key facts about {q}?",
        ]
    # de-dupe preserve order
    out: list[str] = []
    seen: set[str] = set()
    for v in variants:
        key = v.strip().lower()
        if key and key not in seen:
            seen.add(key)
            out.append(v.strip())
    return out


async def llm_clean_scrape_markdown(query: str, raw_markdown: str) -> str:
    """DeepSeek: turn a long Google AI Overview into useful sentences only.

    Returns empty string on skip/failure (caller keeps heuristic/raw).
    Never invents facts — only keeps sentences already in the Overview.
    """
    if not google_ai_llm_clean_enabled():
        return ""
    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        return ""
    # 12000 chars was ~3000 input tokens per call. The Overview's actual
    # answer sits at the top; the tail is citations, related searches and
    # nav chrome, which the cleaner discards anyway — so paying to send it
    # is waste. Configurable if a market needs more context.
    trimmed = heuristic_clean_scrape_markdown(
        raw_markdown, max_chars=_llm_clean_max_input_chars()
    )
    if len(trimmed) < 40:
        return trimmed

    base_url = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").strip().rstrip("/")
    model = _llm_clean_model()

    system = (
        "You clean Google AI Overview text for vendor research. "
        "The Overview is long and noisy. Keep ONLY useful factual sentences "
        "that answer the search query. "
        "Drop navigation, ads, People also ask, related searches, cookie banners, "
        "duplicate lines, and UI chrome. "
        "Copy or lightly tighten sentences from the Overview — do not invent "
        "years, cities, emails, phones, names, or URLs. "
        "Keep the company name in at least one sentence. "
        "If the Overview does not answer the query, return an empty string."
    )
    user = (
        f"Search query: {query.strip()}\n\n"
        "Google AI Overview (long / noisy):\n"
        f"```\n{trimmed}\n```\n\n"
        "Return 1–8 useful factual sentences (or short bullets) that answer the query. "
        "Plain text only. No JSON. No preamble."
    )
    payload = {
        "model": model,
        "temperature": 0,
        "max_tokens": 700,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    }
    try:
        async with httpx.AsyncClient(timeout=60.0) as client:
            r = await client.post(
                f"{base_url}/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                },
                json=payload,
            )
            if r.status_code != 200:
                return ""
            data = r.json()
            content = (
                ((data.get("choices") or [{}])[0].get("message") or {}).get("content") or ""
            ).strip()
            if content.startswith("```"):
                content = re.sub(r"^```\w*\n?", "", content)
                content = re.sub(r"\n?```$", "", content).strip()
            return content[:4000]
    except Exception:
        return ""


async def google_ai_ping(base_url: str = "") -> bool:
    """True when backend responds (extension may still be idle)."""
    global _health_cache
    now = time.monotonic()
    if now - _health_cache[0] < _HEALTH_TTL_SEC:
        return _health_cache[1]
    url = (base_url or google_ai_base_url()).rstrip("/") + "/health"
    ok = False
    try:
        async with httpx.AsyncClient(timeout=5.0) as client:
            r = await client.get(url)
            ok = r.status_code == 200
    except Exception:
        ok = False
    _health_cache = (now, ok)
    return ok


async def _ask_once(
    base: str,
    query: str,
    *,
    thread_id: str | None = None,
    mode: str = "pro",
) -> dict:
    params: dict[str, str] = {"q": query, "mode": mode}
    if thread_id:
        params["thread_id"] = thread_id
    try:
        async with httpx.AsyncClient(timeout=_timeout()) as client:
            r = await client.get(f"{base}/ask", params=params)
            if r.status_code != 200:
                return {
                    "markdown": "",
                    "citations": [],
                    "error": f"http_{r.status_code}",
                    "detail": (r.text or "")[:300],
                    "query_used": query,
                }
            data = r.json()
            if not isinstance(data, dict):
                return {"markdown": "", "citations": [], "error": "bad_json", "query_used": query}
            data["query_used"] = query
            return data
    except httpx.TimeoutException:
        return {"markdown": "", "citations": [], "error": "timeout", "query_used": query}
    except Exception as exc:
        return {
            "markdown": "",
            "citations": [],
            "error": type(exc).__name__,
            "detail": str(exc)[:200],
            "query_used": query,
        }


def _heuristic_only(data: dict) -> dict:
    raw_md = str(data.get("markdown") or data.get("markdown_raw") or "").strip()
    if not raw_md:
        data["ai_overview_missing"] = True
        data["llm_cleaned"] = False
        return data
    data["markdown_raw"] = data.get("markdown_raw") or raw_md
    heur = heuristic_clean_scrape_markdown(str(data["markdown_raw"]))
    data["markdown"] = heur or str(data["markdown_raw"])
    data["ai_overview_missing"] = _ai_overview_missing(data)
    data["llm_cleaned"] = False
    if not data["ai_overview_missing"]:
        data["source"] = "google_ai_overview"
    return data


async def _apply_deepseek_sentence_clean(
    query: str,
    data: dict,
    *,
    llm_clean: bool | None,
) -> dict:
    """After a Google AI Overview is chosen: DeepSeek keeps useful sentences only."""
    data = _heuristic_only(data)
    raw_for_llm = str(data.get("markdown") or data.get("markdown_raw") or "")
    if not raw_for_llm.strip():
        return data

    want_llm = google_ai_llm_clean_enabled() if llm_clean is None else bool(llm_clean)
    only_missing = google_ai_llm_clean_only_when_missing()
    missing = bool(data.get("ai_overview_missing"))
    # Default: clean long Overview text. Optional: only when Overview is missing.
    should_clean = want_llm and (missing if only_missing else True)

    # Skip the LLM when the heuristic clean already left something short.
    # Downstream consumers parse this with regex, not prose comprehension,
    # so paying DeepSeek to shorten an already-short answer buys nothing —
    # and measured across a full market it was the bulk of the LLM spend.
    # Long, chrome-heavy responses still get cleaned, because there the LLM
    # measurably improves what the extractors can find.
    if should_clean and len(raw_for_llm) <= _llm_clean_min_chars():
        data["llm_cleaned"] = False
        data["llm_clean_skipped"] = "short_enough"
        return data

    if not should_clean:
        return data

    before = len(raw_for_llm)
    cleaned = await llm_clean_scrape_markdown(query, raw_for_llm)
    if cleaned and len(cleaned) >= 20:
        data["markdown"] = cleaned
        data["llm_cleaned"] = True
        data["llm_clean_model"] = _llm_clean_model()
        data["source"] = "google_ai_overview+deepseek"
        print(
            f"  [google_ai] deepseek useful sentences {before}->{len(cleaned)} chars | {query[:90]}",
            flush=True,
        )
    else:
        data["llm_cleaned"] = False
    return data


async def _apply_clean(query: str, data: dict, *, llm_clean: bool | None) -> dict:
    """Heuristic strip, then DeepSeek useful-sentence clean (once per chosen scrape)."""
    return await _apply_deepseek_sentence_clean(query, data, llm_clean=llm_clean)


async def google_ai_ask(
    query: str,
    *,
    base_url: str = "",
    thread_id: str | None = None,
    llm_clean: bool | None = None,
) -> dict:
    """Call /ask. Retries when Google skips AI Overview (FORCE_OVERVIEW).

    When LLM clean is enabled, sets:
      markdown_raw  — original scrape
      markdown      — cleaned text (heuristic + optional LLM)
      llm_cleaned   — True when LLM rewrite was applied
      ai_overview_missing — True if no real overview after retries
      overview_attempts — how many /ask tries were made
    """
    q = (query or "").strip()
    if not q:
        return {"markdown": "", "citations": [], "error": "empty_query"}

    base = (base_url or google_ai_base_url()).rstrip("/")
    mode = _ask_mode()
    force = google_ai_force_overview_enabled()
    variants = _overview_query_variants(q) if force else [q]
    max_tries = _overview_retries() if force else 1
    sleep_sec = _overview_retry_sleep_sec()

    best: dict = {"markdown": "", "citations": [], "error": "no_result"}
    attempts = 0

    for i, variant in enumerate(variants):
        if attempts >= max_tries:
            break
        # First attempt may reuse caller thread_id; later retries open a fresh thread
        tid = thread_id if attempts == 0 else None

        # Before issuing the request, wait out any block another query is
        # already sitting on. Concurrency means ~10 tabs are in flight; if
        # one is blocked they all are, so continuing to fire requests both
        # wastes them and prolongs the block.
        await _captcha_gate.wait_if_blocked()

        data = await _ask_once(base, variant, thread_id=tid, mode=mode)

        # A CAPTCHA is temporary. Pause (letting the user solve it in the
        # visible tab if they wish) and re-ask the SAME query rather than
        # discarding the company.
        blocked_tries = 0
        while is_captcha_page(data) and blocked_tries < _captcha_retries():
            blocked_tries += 1
            wait = _captcha_wait_sec()
            print(
                f"  [google_ai] CAPTCHA detected — pausing all tabs "
                f"({blocked_tries}/{_captcha_retries()}) | {variant[:60]}",
                flush=True,
            )
            if _captcha_wait_for_solve():
                # Hold the whole queue until Google answers again. Opening
                # more tabs into an active block just prolongs it.
                await _captcha_gate.hold_until_clear(base, mode, _ask_once)
            else:
                await _captcha_gate.block_for(wait)
            data = await _ask_once(base, variant, thread_id=None, mode=mode)

        if is_captcha_page(data):
            # Still blocked after every retry. Report it as such so callers
            # can skip the company instead of parsing the block page — the
            # page echoes the query, so it reads as a confident answer.
            return {
                "markdown": "",
                "markdown_raw": str(data.get("markdown_raw") or data.get("markdown") or ""),
                "citations": [],
                "error": "captcha_blocked",
                "captcha_blocked": True,
                "overview_attempts": attempts + 1,
            }

        attempts += 1
        data = _heuristic_only(data)
        data["overview_attempts"] = attempts
        data["overview_query"] = variant

        if not _ai_overview_missing(data):
            data["ai_overview_missing"] = False
            # One DeepSeek pass: long Overview → useful sentences for this query
            return await _apply_deepseek_sentence_clean(q, data, llm_clean=llm_clean)

        # Keep the richest attempt so far for final DeepSeek clean
        best_len = len(str(best.get("markdown") or ""))
        cur_len = len(str(data.get("markdown") or ""))
        if cur_len >= best_len:
            best = data

        if attempts < max_tries and i + 1 < len(variants):
            await asyncio.sleep(sleep_sec)

    best["overview_attempts"] = attempts
    best["ai_overview_missing"] = True
    if not best.get("error") or str(best.get("error")).startswith("no_ai"):
        best["error"] = best.get("error") or "no_ai_overview_after_retries"
    return await _apply_deepseek_sentence_clean(q, best, llm_clean=llm_clean)


def google_ai_response_to_hits(query: str, data: dict) -> list[GoogleAiHit]:
    """Map /ask payload into title/link/snippet hits for the search router."""
    hits: list[GoogleAiHit] = []
    md = str(data.get("markdown") or "").strip()
    qid = str(data.get("query_id") or "overview")
    search_link = f"https://www.google.com/search?q={quote_plus(query)}"

    if md:
        title = "Google Search (cleaned)" if data.get("llm_cleaned") else "Google Search scrape"
        hits.append(
            GoogleAiHit(
                title=f"{title} — {query[:100]}",
                link=search_link,
                snippet=md[:2500],
            )
        )

    citations = data.get("citations") or []
    if isinstance(citations, list):
        for i, c in enumerate(citations):
            if isinstance(c, str) and c.startswith("http"):
                hits.append(
                    GoogleAiHit(
                        title=f"Citation {i + 1}",
                        link=c,
                        snippet=(md[:400] if md else ""),
                    )
                )
            elif isinstance(c, dict):
                link = str(c.get("url") or c.get("link") or "").strip()
                if not link.startswith("http"):
                    continue
                hits.append(
                    GoogleAiHit(
                        title=str(c.get("title") or f"Citation {i + 1}")[:200],
                        link=link,
                        snippet=str(c.get("snippet") or c.get("content") or md[:400])[:500],
                    )
                )
    return hits


async def google_ai_search(query: str, *, max_results: int = 12) -> list[GoogleAiHit]:
    if not google_ai_enabled():
        return []
    if not await google_ai_ping():
        return []
    data = await google_ai_ask(query)
    if data.get("error") and not data.get("markdown"):
        return []
    return google_ai_response_to_hits(query, data)[:max_results]


async def google_ai_ask_facts(query: str, *, company: str = "") -> dict:
    """Ask + LLM-clean, then extract founded/HQ JSON for gap-fill helpers."""
    data = await google_ai_ask(query, llm_clean=True)
    md = str(data.get("markdown") or "").strip()
    facts: dict = {
        "markdown": md,
        "markdown_raw": data.get("markdown_raw") or "",
        "llm_cleaned": bool(data.get("llm_cleaned")),
        "error": data.get("error"),
        "founded_year": "",
        "headquarters": "",
    }
    if not md or not (os.getenv("OPENAI_API_KEY") or "").strip():
        return facts

    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base_url = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").strip().rstrip("/")
    model = _llm_clean_model()
    label = (company or query).strip()
    prompt = (
        f"Company/topic: {label}\n\n"
        "From the cleaned text below, extract JSON keys: "
        "founded_year (YYYY or empty), headquarters (city/country or empty), "
        "summary (1-2 sentence factual summary).\n"
        "Do not invent. Use empty strings when unknown.\n\n"
        f"{md[:5000]}"
    )
    try:
        async with httpx.AsyncClient(timeout=45.0) as client:
            r = await client.post(
                f"{base_url}/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json",
                },
                json={
                    "model": model,
                    "temperature": 0,
                    "response_format": {"type": "json_object"},
                    "messages": [{"role": "user", "content": prompt}],
                },
            )
            if r.status_code != 200:
                return facts
            content = (
                ((r.json().get("choices") or [{}])[0].get("message") or {}).get("content") or ""
            ).strip()
            parsed = json.loads(content) if content else {}
            if isinstance(parsed, dict):
                facts["founded_year"] = str(parsed.get("founded_year") or "").strip()
                facts["headquarters"] = str(parsed.get("headquarters") or "").strip()
                summary = str(parsed.get("summary") or "").strip()
                if summary:
                    facts["markdown"] = summary
                    extra = []
                    if facts["founded_year"]:
                        extra.append(f"Founded: {facts['founded_year']}")
                    if facts["headquarters"]:
                        extra.append(f"HQ: {facts['headquarters']}")
                    if extra:
                        facts["markdown"] = summary + "\n" + " | ".join(extra)
    except Exception:
        pass
    return facts
