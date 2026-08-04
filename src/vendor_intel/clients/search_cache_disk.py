"""Disk-backed search cache for single-market / Streamlit runs (complements batch SQLite)."""
from __future__ import annotations

import hashlib
import json
import os
import time
from pathlib import Path
from typing import Any

_DEFAULT_TTL_SEC = 7 * 24 * 3600


def _cache_dir() -> Path:
    from vendor_intel.config import _project_root

    raw = (os.getenv("SEARCH_DISK_CACHE_DIR") or "data/search_cache").strip()
    base = Path(raw)
    if not base.is_absolute():
        base = _project_root() / base
    base.mkdir(parents=True, exist_ok=True)
    return base


def _cache_key(query: str, geo: str, namespace: str) -> str:
    raw = f"{namespace}|{(geo or '').strip().lower()}|{query.strip().lower()}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:40]


def _ttl_sec() -> int:
    try:
        return max(3600, int(os.getenv("SEARCH_DISK_CACHE_TTL_SEC", str(_DEFAULT_TTL_SEC))))
    except ValueError:
        return _DEFAULT_TTL_SEC


def disk_cache_enabled() -> bool:
    return os.getenv("SEARCH_DISK_CACHE", "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )


def get_disk_cached_search(query: str, geo: str, namespace: str) -> list[dict[str, Any]] | None:
    if not disk_cache_enabled():
        return None
    path = _cache_dir() / f"{_cache_key(query, geo, namespace)}.json"
    if not path.is_file():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        saved_at = float(payload.get("saved_at") or 0)
        if time.time() - saved_at > _ttl_sec():
            path.unlink(missing_ok=True)
            return None
        rows = payload.get("results")
        if isinstance(rows, list):
            return rows
    except Exception:
        return None
    return None


def put_disk_cached_search(
    query: str,
    geo: str,
    namespace: str,
    results: list[dict[str, Any]],
) -> None:
    if not disk_cache_enabled() or not results:
        return
    path = _cache_dir() / f"{_cache_key(query, geo, namespace)}.json"
    try:
        path.write_text(
            json.dumps(
                {
                    "query": query,
                    "geo": geo,
                    "namespace": namespace,
                    "saved_at": time.time(),
                    "results": results,
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    except Exception:
        pass
