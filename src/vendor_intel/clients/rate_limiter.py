"""Shared token-bucket rate limiter + cooldown-based circuit breaker for the
free/scraped search engines (ddgs backends, Bing HTML).

Why this exists: under Phase 3's concurrent batch mode, several markets run
"at once", each independently hammering the same free search engines. Without
a limiter shared across all of them, the aggregate request rate can spike to
what looks like a bot swarm — exactly what gets an engine blocked. This module
is the shared, in-process throttle every concurrent market's searches pass
through.

Thread-safe: acquire()/note_outcome() are plain blocking (time.sleep-based)
functions, meant to be called from a worker thread (duckduckgo.py's and
bing_html.py's HTTP calls already run via asyncio.to_thread), not from the
event-loop thread directly — blocking there would stall every concurrently
running market, not just the one waiting for a token.

Circuit breaker state (dead_until) is checkpointed to the search_rate_state
SQLite table so a killed/resumed batch doesn't immediately re-hammer an
engine it just got blocked by — see batch/db.py's checkpoint_search_rate_state
/ load_search_rate_state. Token levels themselves are NOT persisted (a fresh,
full bucket on process restart is the safe-by-default behavior, not a bug).
"""
from __future__ import annotations

import threading
import time

# (capacity, refill_per_sec) per engine — conservative defaults for free/
# scraped engines. Override per-engine at runtime via configure() if needed.
_DEFAULT_RATES: dict[str, tuple[float, float]] = {
    "ddgs": (3.0, 0.35),
    "ddgs:bing": (4.0, 0.5),
    "ddgs:duckduckgo": (3.0, 0.4),
    "ddgs:brave": (3.0, 0.4),
    "ddgs:yandex": (2.0, 0.3),
    "ddgs:yahoo": (2.0, 0.3),
    "ddgs:mojeek": (1.0, 0.15),
    "ddgs:google": (1.0, 0.08),  # slowest — Google rate-limits aggressively
    "bing_html": (2.0, 0.15),
}
_DEFAULT_RATE = (5.0, 0.5)

# Escalating cooldown on repeated blocks: 1min -> 5min -> 30min, then stays at 30min.
_COOLDOWNS_SECONDS = [60.0, 300.0, 1800.0]

_CHECKPOINT_EVERY = 10  # persist to SQLite every N acquires per engine, not every one

_lock = threading.Lock()
_tokens: dict[str, float] = {}
_last_refill: dict[str, float] = {}
_dead_until: dict[str, float] = {}
_fail_streak: dict[str, int] = {}
_request_count: dict[str, int] = {}
_rates: dict[str, tuple[float, float]] = {}


def configure(engine: str, *, capacity: float, refill_per_sec: float) -> None:
    """Override the default (capacity, refill_per_sec) for one engine."""
    with _lock:
        _rates[engine] = (capacity, refill_per_sec)


def _rate_for(engine: str) -> tuple[float, float]:
    return _rates.get(engine) or _DEFAULT_RATES.get(engine) or _DEFAULT_RATE


def _refill_locked(engine: str) -> None:
    capacity, refill_per_sec = _rate_for(engine)
    now = time.time()
    last = _last_refill.get(engine, now)
    elapsed = max(0.0, now - last)
    current = _tokens.get(engine, capacity)
    _tokens[engine] = min(capacity, current + elapsed * refill_per_sec)
    _last_refill[engine] = now


def is_dead(engine: str) -> bool:
    with _lock:
        until = _dead_until.get(engine)
        return until is not None and time.time() < until


def dead_until(engine: str) -> float | None:
    with _lock:
        return _dead_until.get(engine)


def acquire(engine: str, *, max_wait: float = 30.0) -> bool:
    """Block until a token is available for `engine`, up to `max_wait`
    seconds. Returns False immediately if the engine is circuit-broken, or
    after `max_wait` if no token became available in time — callers should
    skip this engine for the current request rather than proceed.
    """
    deadline = time.time() + max_wait
    while True:
        with _lock:
            if engine in _dead_until and time.time() < _dead_until[engine]:
                try:
                    from vendor_intel.clients import search_stats

                    search_stats.record_rate_limit_skip(engine)
                except Exception:
                    pass
                return False
            _refill_locked(engine)
            if _tokens.get(engine, 0.0) >= 1.0:
                _tokens[engine] -= 1.0
                _request_count[engine] = _request_count.get(engine, 0) + 1
                should_checkpoint = _request_count[engine] % _CHECKPOINT_EVERY == 0
                tokens_snapshot = _tokens[engine]
                if should_checkpoint:
                    _checkpoint(engine, tokens_snapshot)
                return True
            _capacity, refill_per_sec = _rate_for(engine)
            wait_for = (1.0 - _tokens.get(engine, 0.0)) / max(refill_per_sec, 0.01)
        now = time.time()
        if now >= deadline:
            try:
                from vendor_intel.clients import search_stats

                search_stats.record_rate_limit_skip(engine)
            except Exception:
                pass
            return False
        time.sleep(max(0.01, min(wait_for, deadline - now)))


def note_outcome(engine: str, *, blocked: bool) -> None:
    """Record whether the last request to `engine` looked like a block
    (429/403/captcha/no-results-after-retries). Any success clears the fail
    streak; a block escalates the cooldown and persists it."""
    with _lock:
        if not blocked:
            _fail_streak.pop(engine, None)
            return
        streak = _fail_streak.get(engine, 0) + 1
        _fail_streak[engine] = streak
        cooldown = _COOLDOWNS_SECONDS[min(streak - 1, len(_COOLDOWNS_SECONDS) - 1)]
        until = time.time() + cooldown
        _dead_until[engine] = until
        _checkpoint(engine, _tokens.get(engine, 0.0), dead_until_override=until)


def _checkpoint(engine: str, tokens: float, *, dead_until_override: float | None = None) -> None:
    try:
        from vendor_intel.batch.db import checkpoint_search_rate_state

        du = dead_until_override if dead_until_override is not None else _dead_until.get(engine)
        checkpoint_search_rate_state(engine, tokens=tokens, dead_until=du)
    except Exception:
        pass  # rate-limit persistence is best-effort, never fail a search over it


def hydrate_from_db() -> None:
    """Call once at batch startup to restore circuit-breaker cooldowns (not
    token levels) from a prior run, so a resumed batch doesn't immediately
    re-hammer an engine it was just blocked by."""
    try:
        from vendor_intel.batch.db import load_search_rate_state

        state = load_search_rate_state()
    except Exception:
        return
    now = time.time()
    with _lock:
        for engine, until in state.items():
            if until is not None and until > now:
                _dead_until[engine] = until


def reset_for_tests() -> None:
    with _lock:
        _tokens.clear()
        _last_refill.clear()
        _dead_until.clear()
        _fail_streak.clear()
        _request_count.clear()
        _rates.clear()
