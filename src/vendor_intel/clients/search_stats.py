"""In-process search telemetry — cache hits, rate limits, retries, engine usage."""
from __future__ import annotations

import json
import threading
import time
from pathlib import Path
from typing import Any

_lock = threading.Lock()
_stats: dict[str, Any] = {
    "total_searches": 0,
    "cache_hits": 0,
    "rate_limit_skips": 0,
    "failed_searches": 0,
    "retries": 0,
    "fallback_used": 0,
    "successful_searches": 0,
    "by_engine": {},
    "started_at": time.time(),
}


def _inc(key: str, amount: int = 1) -> None:
    with _lock:
        _stats[key] = int(_stats.get(key) or 0) + amount


def record_search_start() -> None:
    _inc("total_searches")


def record_cache_hit() -> None:
    _inc("cache_hits")


def record_rate_limit_skip(engine: str = "") -> None:
    _inc("rate_limit_skips")
    if engine:
        record_engine(engine, "skipped")


def record_retry() -> None:
    _inc("retries")


def record_failure() -> None:
    _inc("failed_searches")


def record_success(engine: str = "") -> None:
    _inc("successful_searches")
    if engine:
        record_engine(engine, "ok")


def record_fallback() -> None:
    _inc("fallback_used")


def record_engine(engine: str, outcome: str) -> None:
    with _lock:
        bucket = _stats.setdefault("by_engine", {})
        eng = bucket.setdefault(engine, {"ok": 0, "skipped": 0, "failed": 0})
        if outcome in eng:
            eng[outcome] += 1


def snapshot() -> dict[str, Any]:
    with _lock:
        elapsed = max(1.0, time.time() - float(_stats.get("started_at") or time.time()))
        total = int(_stats.get("total_searches") or 0)
        return {
            **dict(_stats),
            "elapsed_sec": round(elapsed, 1),
            "avg_sec_per_search": round(elapsed / max(total, 1), 2),
        }


def reset_for_tests() -> None:
    with _lock:
        _stats.clear()
        _stats.update(
            {
                "total_searches": 0,
                "cache_hits": 0,
                "rate_limit_skips": 0,
                "failed_searches": 0,
                "retries": 0,
                "fallback_used": 0,
                "successful_searches": 0,
                "by_engine": {},
                "started_at": time.time(),
            }
        )


def log_summary(*, prefix: str = "search_stats") -> None:
    if not _env_true("SEARCH_STATS_LOG", default=True):
        return
    snap = snapshot()
    print(
        f"  [{prefix}] searches={snap['total_searches']} ok={snap['successful_searches']} "
        f"cache_hits={snap['cache_hits']} rate_limit_skips={snap['rate_limit_skips']} "
        f"retries={snap['retries']} failed={snap['failed_searches']} "
        f"fallback={snap['fallback_used']} avg={snap['avg_sec_per_search']}s/search",
        flush=True,
    )


def persist_summary(path: Path | None = None) -> Path | None:
    """Write rolling stats JSON (best-effort checkpoint)."""
    try:
        from vendor_intel.config import _project_root

        out = path or (_project_root() / "data" / "search_stats.json")
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(snapshot(), indent=2), encoding="utf-8")
        return out
    except Exception:
        return None


def _env_true(name: str, *, default: bool = False) -> bool:
    import os

    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in ("1", "true", "yes", "on")
