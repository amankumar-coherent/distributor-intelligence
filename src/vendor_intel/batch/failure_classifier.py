"""Classify a market-run exception as 'transient' (worth retrying) or
'permanent' (retrying won't help) — decides whether batch/db.py's
mark_market_failed requeues the market or parks it as failed.

Extends the same string-matching heuristic already used for individual
search-engine errors in clients/ddg_worker_pool.py (_is_transient_engine_error),
generalized to whole-market pipeline failures (network/LLM/search errors vs.
programming/config errors).
"""
from __future__ import annotations

_TRANSIENT_MARKERS = (
    "connecttimeout",
    "timeout",
    "connection attempt failed",
    "connectionerror",
    "connection reset",
    "connection refused",
    "read timed out",
    "server disconnected",
    "winerror 10060",
    "getaddrinfo failed",
    "network is unreachable",
    "temporarily unavailable",
    "captcha",
    "ratelimit",
    "rate limit",
    "too many requests",
    "429",
    "500",
    "502",
    "503",
    "504",
    "no results found",
)

# Exception types that are almost always a bug or bad input, not a flaky
# network blip — retrying without a code/config/data fix would just waste
# budget re-running the same failure.
_PERMANENT_EXCEPTION_TYPES = (
    "ValueError",
    "TypeError",
    "KeyError",
    "AttributeError",
    "FileNotFoundError",
    "NotADirectoryError",
    "PermissionError",
    "ImportError",
    "ModuleNotFoundError",
)


def classify_failure(exc: BaseException) -> str:
    """Return 'transient' or 'permanent' for a market-run exception.

    Defaults to 'transient' when uncertain — a wasted retry (bounded by
    markets.max_attempts) is cheaper than silently dropping a market that
    would have succeeded on a second try.
    """
    message = str(exc).lower()
    if any(marker in message for marker in _TRANSIENT_MARKERS):
        return "transient"

    if type(exc).__name__ in _PERMANENT_EXCEPTION_TYPES:
        return "permanent"

    return "transient"
