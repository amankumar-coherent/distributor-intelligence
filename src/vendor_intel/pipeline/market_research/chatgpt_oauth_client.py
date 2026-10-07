"""Client for the free ChatGPT-account-backed LLM engine (EvanZhouDev/openai-oauth,
run via `npx openai-oauth@latest` from the published npm package — NOT the
unbuilt reference clone at vendor/openai-oauth/, which has no dist/ output
and whose package name would otherwise collide with npx's resolution).

openai-oauth turns a real, logged-in ChatGPT account into a local
OpenAI-compatible HTTP endpoint (default http://127.0.0.1:10531/v1) — no API
key, no Selenium, no anonymous-only/Google-SSO-blocked limitation that
vendor_intel.scraping.chatgpt_scraper has.

Per explicit user instruction, this client is ChatGPT-ONLY — it does NOT
fall back to any other LLM_PROVIDER (Anthropic/DeepSeek/etc.) on a ChatGPT
failure. Every company must come exclusively from the user's own ChatGPT
account. Instead of switching providers, a ChatGPT failure (proxy down,
usage limit reached, timeout) is retried against ChatGPT itself with a
wait in between — see chat()/chat_json()'s retry loop below. This was a
deliberate reversal of an earlier fallback-to-DeepSeek design: DeepSeek
turned out to independently hit its own usage limits at the same time as
ChatGPT, which meant a run could silently start pulling companies from a
different provider than the user asked for.

This module does NOT reimplement an LLM call path. It only:
  1. Manages the `npx openai-oauth@latest` local proxy process (start/
     health-check). Always passes --yes so npx's "install this package?"
     confirmation prompt can't hang forever with stdin devnulled (this bit
     us during integration — the process looked "started" but never became
     healthy because it was silently blocked on that prompt).
  2. Points vendor_intel.placeholders.llm's existing OpenAI-compatible
     dispatch (_openai_complete, reached via LLM_PROVIDER="openai") at that
     proxy for the duration of one call, then restores prior state.
  3. Retries the SAME proxy call (with a pause) on any failure — proxy not
     running, not logged in, a request error, or a usage-limit response —
     up to _MAX_RETRIES times before giving up and raising, rather than
     ever routing to a different provider.

First-run login is interactive and NOT done automatically here: run
`npx openai-oauth@latest login` once (from the repo root, or anywhere) to
establish a session before using this client — credentials are saved to
~/.codex/auth.json regardless of cwd. That is a deliberate choice —
auto-driving a real OAuth browser login from unattended code is out of scope
and the user must authorize their own ChatGPT account once, interactively.
Known upstream issue (EvanZhouDev/openai-oauth#27, open as of integration
time): the CLI's auto-opened browser tab can fail with an OpenAI-side
"Authentication Error: A required parameter is missing" even though the
printed login URL is well-formed — the reliable workaround is copying that
URL from the terminal and pasting it into a fresh browser tab manually
rather than letting the CLI auto-open it.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import time
from pathlib import Path

import httpx

REPO_ROOT = Path(__file__).resolve().parents[4]
DEFAULT_PORT = 10531
DEFAULT_BASE_URL = f"http://127.0.0.1:{DEFAULT_PORT}/v1"

_HEALTH_TIMEOUT = 3.0
_START_POLL_INTERVAL = 1.0


def _log(msg: str) -> None:
    print(f"  [chatgpt-oauth] {msg}", flush=True)


def _npx_path() -> str | None:
    return shutil.which("npx")


def is_healthy(base_url: str = DEFAULT_BASE_URL) -> bool:
    """True if a local openai-oauth proxy is already up and serving models."""
    try:
        resp = httpx.get(f"{base_url.rstrip('/')}/models", timeout=_HEALTH_TIMEOUT)
        return resp.status_code == 200
    except Exception:
        return False


def ensure_running(*, port: int = DEFAULT_PORT, timeout: float = 30.0) -> bool:
    """Start the openai-oauth local proxy (detached) if it isn't already up.

    Returns True if the proxy is confirmed healthy by the time this returns,
    False otherwise (npx missing, spawn failed, or health check never
    succeeded within `timeout`) — callers must treat False as "use the
    configured LLM fallback instead", never raise.
    """
    base_url = f"http://127.0.0.1:{port}/v1"
    if is_healthy(base_url):
        return True

    npx = _npx_path()
    if not npx:
        _log("npx not found on PATH — cannot start openai-oauth proxy")
        return False

    try:
        subprocess.Popen(
            [npx, "--yes", "openai-oauth@latest", "--port", str(port), "--detach"],
            # Run from repo root, NOT vendor/openai-oauth/ — that's an unbuilt
            # monorepo clone (kept for reference only) whose package name
            # collides with the real one; running npx from inside it can
            # make npx try to resolve the local (unbuilt, dist/-less)
            # workspace package instead of fetching the published npm one.
            cwd=str(REPO_ROOT),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL,
        )
    except Exception as exc:  # noqa: BLE001
        _log(f"failed to start openai-oauth: {type(exc).__name__}: {exc}")
        return False

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if is_healthy(base_url):
            _log(f"proxy healthy at {base_url}")
            return True
        time.sleep(_START_POLL_INTERVAL)
    _log(f"proxy did not become healthy within {timeout}s")
    return False


class ChatGPTOAuthUnavailable(RuntimeError):
    """Raised only internally; chat()/chat_json() catch this and fall back."""


# The proxy only serves whatever Codex/ChatGPT models the logged-in account
# actually has access to (discovered via its own /v1/models) — NOT arbitrary
# OpenAI API model ids. placeholders.llm's default model for LLM_PROVIDER=
# "openai" is "gpt-4o-mini", which this proxy rejects with a bare HTTP 500
# ("Bad Request") since it isn't one of the account's models. Must always
# pass an explicit model this proxy actually serves.
DEFAULT_MODEL = "gpt-5.4-mini"


def _looks_like_llm_failure(text: str) -> bool:
    """placeholders.llm._dispatch() swallows HTTP errors into a JSON string
    payload (status: "llm_failed") instead of raising — so a real failure
    would otherwise look like a normal successful reply and never trigger
    the fallback. Detect that shape so chat()/chat_json() can still fall
    back on it."""
    s = (text or "").strip()
    if not s.startswith("{") or '"status"' not in s:
        return False
    try:
        import json as _json

        return _json.loads(s).get("status") == "llm_failed"
    except (ValueError, TypeError):
        return False


# A full 20-company draft response (reasoning + ~15 fields x 20 companies)
# measured at ~267s against the real openai-oauth proxy during integration
# — well past placeholders.llm's hardcoded 120s _TIMEOUT, which made every
# such call look like a failure and silently fall back. placeholders.llm
# has no per-call timeout parameter, so this is patched the same way
# LLM_PROVIDER already is: temporarily, only for the duration of one call
# through this module, restored immediately after — never a permanent
# change to the shared timeout every other pipeline in this repo relies on.
PROXY_TIMEOUT = 300.0


def _via_proxy(
    system: str, user: str, *, port: int, max_tokens: int, json_mode: bool, model: str
):
    """Run one completion through the openai-oauth proxy by temporarily
    pointing placeholders.llm's OpenAI dispatch at it, then restoring prior
    provider/env state. Raises ChatGPTOAuthUnavailable on any failure —
    proxy not running, not logged in, or the request itself erroring."""
    from vendor_intel import placeholders

    llm = placeholders.llm  # module, so we can temporarily patch LLM_PROVIDER

    if not ensure_running(port=port):
        raise ChatGPTOAuthUnavailable("openai-oauth proxy is not available")

    base_url = f"http://127.0.0.1:{port}/v1"
    prior_provider = llm.LLM_PROVIDER
    prior_base_url = os.environ.get("OPENAI_BASE_URL")
    prior_api_key = os.environ.get("OPENAI_API_KEY")
    prior_timeout = llm._TIMEOUT
    try:
        llm.LLM_PROVIDER = "openai"
        llm._TIMEOUT = PROXY_TIMEOUT
        os.environ["OPENAI_BASE_URL"] = base_url
        os.environ["OPENAI_API_KEY"] = os.environ.get("OPENAI_API_KEY") or "unused"
        if json_mode:
            result = llm.llm_complete_json(system, user, model=model, max_tokens=max_tokens)
        else:
            result = llm.llm_complete(system, user, model=model, max_tokens=max_tokens)
        if isinstance(result, str) and _looks_like_llm_failure(result):
            raise RuntimeError(f"proxy request failed: {result[:200]}")
        return result
    except Exception as exc:  # noqa: BLE001
        raise ChatGPTOAuthUnavailable(str(exc)) from exc
    finally:
        llm.LLM_PROVIDER = prior_provider
        llm._TIMEOUT = prior_timeout
        if prior_base_url is None:
            os.environ.pop("OPENAI_BASE_URL", None)
        else:
            os.environ["OPENAI_BASE_URL"] = prior_base_url
        if prior_api_key is None:
            os.environ.pop("OPENAI_API_KEY", None)
        else:
            os.environ["OPENAI_API_KEY"] = prior_api_key


# ChatGPT-only: on any failure, retry the SAME proxy call this many times
# (with a pause between attempts) rather than ever switching providers.
#
# Two very different failure shapes need two very different pauses:
#   - A genuine account USAGE LIMIT (free-tier message cap) does not clear
#     in seconds or even a minute — it's an hours-scale reset. Retrying
#     every 60s just wastes 4 attempts and gives up in ~4 minutes, long
#     before the limit could plausibly have cleared. This case gets a much
#     longer pause and more attempts, since the whole point of "wait it
#     out" is to actually span the reset window.
#   - Any OTHER failure (proxy not started yet, a one-off network blip)
#     really can clear in seconds, so a short pause is appropriate there —
#     no reason to wait minutes for something that's likely already fixed.
_MAX_RETRIES = 4
_RETRY_PAUSE_SECONDS = 60.0

_MAX_USAGE_LIMIT_RETRIES = 12
_USAGE_LIMIT_RETRY_PAUSE_SECONDS = 900.0  # 15 min between checks, up to 3h total

_USAGE_LIMIT_MARKERS = ("usage limit", "rate limit", "quota")


def _looks_like_usage_limit(text: str) -> bool:
    t = (text or "").lower()
    return any(marker in t for marker in _USAGE_LIMIT_MARKERS)


def _chat_with_retry(system: str, user: str, *, port: int, max_tokens: int, json_mode: bool, model: str):
    """Calls _via_proxy, retrying on ChatGPTOAuthUnavailable. Raises the
    last ChatGPTOAuthUnavailable if every attempt fails — callers must not
    silently substitute a different provider's answer.

    A usage-limit-shaped failure gets its own longer retry budget
    (_MAX_USAGE_LIMIT_RETRIES x _USAGE_LIMIT_RETRY_PAUSE_SECONDS, spanning
    hours) since that's the only failure shape that genuinely needs a long
    wait to clear; any other failure uses the short _MAX_RETRIES budget.
    """
    last_exc: ChatGPTOAuthUnavailable | None = None
    attempt = 0
    usage_limit_attempt = 0
    while True:
        attempt += 1
        try:
            return _via_proxy(system, user, port=port, max_tokens=max_tokens, json_mode=json_mode, model=model)
        except ChatGPTOAuthUnavailable as exc:
            last_exc = exc
            if _looks_like_usage_limit(str(exc)):
                usage_limit_attempt += 1
                if usage_limit_attempt >= _MAX_USAGE_LIMIT_RETRIES:
                    break
                _log(
                    f"ChatGPT usage limit reached — waiting "
                    f"{_USAGE_LIMIT_RETRY_PAUSE_SECONDS / 60:.0f} min before checking again "
                    f"(attempt {usage_limit_attempt}/{_MAX_USAGE_LIMIT_RETRIES})"
                )
                time.sleep(_USAGE_LIMIT_RETRY_PAUSE_SECONDS)
                continue
            if attempt >= _MAX_RETRIES:
                break
            _log(
                f"ChatGPT call failed ({exc}) — retrying in {_RETRY_PAUSE_SECONDS:.0f}s "
                f"(attempt {attempt}/{_MAX_RETRIES})"
            )
            time.sleep(_RETRY_PAUSE_SECONDS)
    assert last_exc is not None
    raise last_exc


def chat(
    system: str,
    user: str,
    *,
    port: int = DEFAULT_PORT,
    max_tokens: int = 2000,
    model: str = DEFAULT_MODEL,
) -> str:
    """Plain-text completion via the user's own ChatGPT account only —
    retries on failure, never falls back to a different LLM_PROVIDER."""
    return _chat_with_retry(system, user, port=port, max_tokens=max_tokens, json_mode=False, model=model)


def chat_json(
    system: str,
    user: str,
    *,
    port: int = DEFAULT_PORT,
    max_tokens: int = 4096,
    model: str = DEFAULT_MODEL,
):
    """JSON completion via the user's own ChatGPT account only — retries on
    failure, never falls back to a different LLM_PROVIDER."""
    return _chat_with_retry(system, user, port=port, max_tokens=max_tokens, json_mode=True, model=model)
