"""Google AI Mode backend for the market-research pipeline.

Drop-in replacement for chatgpt_oauth_client: same chat()/chat_json()
signatures, so discovery.py needs only a one-line import swap.

WHY THIS REPLACED THE CHATGPT BACKEND — the ChatGPT-account route
(openai-oauth / Codex backend-api) works, but on a free plan it hits
"The usage limit has been reached" after a few dozen calls, which stalled
--target 500 runs at ~110 companies. Codex usage is capped separately from
the (now uncapped) regular ChatGPT text chat, so "unlimited free text chat"
does not apply to it. Every alternative route to a real logged-in ChatGPT
session was tested and blocked: a copied Chrome profile can't decrypt its
own cookies in a new browser instance, a fresh Playwright profile gets
Google's "This browser or app may not be secure" OAuth refusal and then a
failing Cloudflare Turnstile, and driving the real Chrome profile directly
hangs at launch (organization-managed Chrome).

Google AI Mode has none of those problems: no account, no login, no
per-account quota. Verified live — a batched "reply with ONLY a JSON array"
prompt returned clean valid JSON in ~10s (vs ~166s for the same shape of
request through the ChatGPT proxy).

It is NOT a metered API, so there is no usage-limit error to retry around.
The failure mode is a CAPTCHA instead, and vendor_intel.scraping.
google_ai_mode already owns that handling (a consecutive-CAPTCHA circuit
breaker that fails fast once Google has flagged the IP/profile, so a bad
run can't spin for hours). This module deliberately does not re-implement
any of it — it only adapts the text-in/text-out `ask()` to the
chat()/chat_json() shape the pipeline expects.
"""
from __future__ import annotations

import json
import os
from typing import Any


def _log(msg: str) -> None:
    print(f"  [ai-mode-client] {msg}", flush=True)


class AIModeUnavailable(RuntimeError):
    """AI Mode returned nothing (CAPTCHA, timeout, or empty answer).

    Named/raised to match chatgpt_oauth_client.ChatGPTOAuthUnavailable so
    callers can treat both backends identically.
    """


def _ensure_enabled() -> None:
    """AI Mode is opt-in via env, but this pipeline's whole discovery step
    depends on it, so default it on rather than silently doing nothing.

    Set GOOGLE_AI_MODE_ENABLED=false to genuinely disable it.
    """
    if not (os.getenv("GOOGLE_AI_MODE_ENABLED") or "").strip():
        os.environ["GOOGLE_AI_MODE_ENABLED"] = "true"


def use_relay() -> bool:
    """Whether to route through the google-ai-scraper relay (:15551)
    instead of the Patchright AI Mode browser.

    DEFAULT IS OFF, and that is deliberate: the relay does NOT actually
    deliver AI Mode. Its extension aims at udm=50, but modern Chrome
    hijacks that into an internal `chrome://contextual-tasks` page that
    extensions cannot read, so the relay silently falls back to a normal
    Google search (AI Overview). Its replies are still LABELLED
    "AI Mode reply for ...", which makes the fallback easy to miss — that
    label is the relay's own text, not evidence of udm=50. The user
    requires genuine AI Mode, so the Patchright path (which navigates
    udm=50 directly, the whole point of that module) stays the default.

    The relay's real advantage, if you ever need it: it drives the user's
    own signed-in Chrome, so Google treats it as a human session and never
    CAPTCHAs it. Patchright uses a separate automation profile, which
    Google can flag — set MARKET_RESEARCH_USE_RELAY=true to trade AI Mode
    for that reliability.
    """
    raw = (os.getenv("MARKET_RESEARCH_USE_RELAY") or "").strip().lower()
    if raw not in ("1", "true", "yes", "on"):
        return False
    try:
        from vendor_intel.clients.google_ai_scraper import google_ai_enabled

        return bool(google_ai_enabled())
    except Exception:
        return False


def _ask_via_relay(prompt: str) -> str:
    """Ask through the google-ai-scraper relay in the user's real Chrome.

    llm_clean is forced OFF: the relay's DeepSeek cleanup flattens a JSON
    reply into bare alternating lines ('"company"' then '"Ingram Micro"'),
    which destroys the structure discovery.py has to parse. The RAW reply
    is already clean, valid JSON in a ```json fence (verified live).
    """
    import asyncio

    from vendor_intel.clients.google_ai_scraper import google_ai_ask

    data = asyncio.run(google_ai_ask(prompt, llm_clean=False))
    text = str(data.get("markdown_raw") or data.get("markdown") or "").strip()
    if not text:
        raise AIModeUnavailable(f"relay returned no answer (error={data.get('error')})")
    return text


def chat(system: str, user: str, *, max_tokens: int = 2000, **_ignored: Any) -> str:
    """Plain-text completion via Google AI Mode.

    `max_tokens` and any other chatgpt_oauth_client-specific kwargs (model,
    port) are accepted and ignored: AI Mode is a web UI, not a metered API,
    so there is nothing to pass them to. Keeping them in the signature is
    what makes this a true drop-in swap.

    AI Mode takes one prompt string, so `system` is prepended to `user`
    rather than sent as a separate role.
    """
    prompt = f"{system.strip()}\n\n{user.strip()}" if system.strip() else user.strip()

    if use_relay():
        return _ask_via_relay(prompt)

    _ensure_enabled()
    from vendor_intel.scraping import google_ai_mode

    text = (google_ai_mode.ask(prompt) or "").strip()
    if not text:
        raise AIModeUnavailable("AI Mode returned no answer (CAPTCHA, timeout, or empty)")
    return text


def chat_json(system: str, user: str, *, max_tokens: int = 4096, **_ignored: Any) -> Any:
    """JSON completion via Google AI Mode.

    AI Mode has no enforced JSON mode, but it does reliably follow an
    explicit "reply with ONLY a JSON array" instruction (verified live).
    Its answer arrives as rendered page text, which routinely carries
    markdown fences and a stray leading "json" language tag, plus Google's
    own "Use code with caution." footer — all stripped here before parsing.
    """
    raw = chat(system, user, max_tokens=max_tokens)
    return _parse_json_answer(raw)


def _parse_json_answer(raw: str) -> Any:
    """Best-effort JSON extraction from AI Mode's rendered answer text."""
    text = (raw or "").strip()

    # Strip the ```json fence / bare "json" language tag AI Mode's rendered
    # code blocks come through as, and Google's trailing UI footer.
    text = text.replace("Use code with caution.", "")
    if text.startswith("```"):
        text = text.strip("`")
    if text.lstrip().lower().startswith("json"):
        text = text.lstrip()[4:]

    # Prefer the widest array, else the widest object — the answer usually
    # has prose around the payload, so slicing to the brackets is what makes
    # this parse at all.
    for opener, closer in (("[", "]"), ("{", "}")):
        start, end = text.find(opener), text.rfind(closer)
        if start >= 0 and end > start:
            try:
                return json.loads(text[start : end + 1])
            except (json.JSONDecodeError, ValueError):
                continue
    raise AIModeUnavailable(f"could not parse JSON from AI Mode answer: {text[:200]}")
