"""Google AI Overview scraper client (HTTP relay on :15551).

Requires the google-ai-scraper FastAPI backend + browser extension:
  cd google-ai-scraper-main/google-ai-scraper-main/server
  uv run google-ai-scraper --backend --port 15551

API: GET /ask?q=... → {markdown, citations, query_id, thread_id, error?}
     GET /health

Optional LLM cleanup (GOOGLE_AI_SCRAPER_LLM_CLEAN=true, uses OPENAI_* keys):
  Strips SERP chrome / nav noise and returns a concise factual answer for gap-fill
  and discovery snippets.

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
    return (os.getenv("GOOGLE_AI_SCRAPER_BASE_URL") or "http://127.0.0.1:15551").strip().rstrip("/")


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
    """Allow OPENAI_* cleanup of scrape text (default on).

    When a real AI Overview is present, LLM clean is skipped by default
    (see google_ai_llm_clean_only_when_missing). Overview text is the source of truth.
    """
    if not google_ai_enabled():
        return False
    raw = (os.getenv("GOOGLE_AI_SCRAPER_LLM_CLEAN") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def google_ai_llm_clean_only_when_missing() -> bool:
    """If true (default), LLM clean runs only when AI Overview is missing/chrome."""
    raw = (os.getenv("GOOGLE_AI_SCRAPER_LLM_CLEAN_ONLY_WHEN_MISSING") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


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
    """Phrasings that more reliably trigger Google AI Overview panels."""
    q = (query or "").strip()
    if not q:
        return []
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
    """Use OpenAI-compatible chat to return only the factual answer text.

    Returns empty string on skip/failure (caller keeps heuristic/raw).
    """
    if not google_ai_llm_clean_enabled():
        return ""
    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        return ""
    trimmed = heuristic_clean_scrape_markdown(raw_markdown, max_chars=8000)
    if len(trimmed) < 40:
        return trimmed

    base_url = (os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").strip().rstrip("/")
    model = (
        os.getenv("GOOGLE_AI_SCRAPER_LLM_MODEL") or os.getenv("OPENAI_MODEL") or "gpt-4o-mini"
    ).strip()

    system = (
        "You clean noisy Google Search / AI Overview scrape text. "
        "Remove navigation, ads, 'People also ask', cookie banners, unrelated links, "
        "and UI chrome. Keep only facts that answer the user query. "
        "Prefer a short clear paragraph plus optional bullet key facts. "
        "Do not invent facts. If nothing useful remains, return an empty string."
    )
    user = (
        f"Query: {query.strip()}\n\n"
        "Raw scrape (may include SERP chrome):\n"
        f"```\n{trimmed}\n```\n\n"
        "Return ONLY the cleaned factual answer text (plain text or light markdown). "
        "No JSON wrapper, no preamble."
    )
    payload = {
        "model": model,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    }
    try:
        async with httpx.AsyncClient(timeout=45.0) as client:
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
            return content[:6000]
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


async def _apply_clean(query: str, data: dict, *, llm_clean: bool | None) -> dict:
    raw_md = str(data.get("markdown") or "").strip()
    if not raw_md:
        data["ai_overview_missing"] = True
        data["llm_cleaned"] = False
        return data

    data["markdown_raw"] = raw_md
    heur = heuristic_clean_scrape_markdown(raw_md)
    data["markdown"] = heur or raw_md
    missing = _ai_overview_missing(data)
    data["ai_overview_missing"] = missing

    # Real AI Overview → trust Google text (heuristic only). Do NOT call LLM.
    # LLM clean only when Overview is missing / SERP chrome (fallback).
    want_llm = google_ai_llm_clean_enabled() if llm_clean is None else bool(llm_clean)
    only_missing = google_ai_llm_clean_only_when_missing()
    if want_llm and (missing or not only_missing) and (heur or raw_md):
        cleaned = await llm_clean_scrape_markdown(query, heur or raw_md)
        if cleaned and len(cleaned) >= 20:
            data["markdown"] = cleaned
            data["llm_cleaned"] = True
        else:
            data["llm_cleaned"] = False
    else:
        data["llm_cleaned"] = False
        if not missing:
            data["source"] = "google_ai_overview"
    return data


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
        data = await _ask_once(base, variant, thread_id=tid, mode=mode)
        attempts += 1
        data = await _apply_clean(q, data, llm_clean=llm_clean)
        data["overview_attempts"] = attempts
        data["overview_query"] = variant

        if not _ai_overview_missing(data):
            data["ai_overview_missing"] = False
            return data

        # Keep the richest attempt so far for final LLM fallback
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
    return best


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
    base_url = (os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").strip().rstrip("/")
    model = (
        os.getenv("GOOGLE_AI_SCRAPER_LLM_MODEL") or os.getenv("OPENAI_MODEL") or "gpt-4o-mini"
    ).strip()
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
