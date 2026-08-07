"""SQLite registry + cross-market caches for batch market runs.

Single file (`data/batch.db`, WAL mode) shared by every market run in a batch:
market queue/state (populated starting Phase 2), a cross-market company crawl
cache (wired in now, via `enrichment/smart_enrichment.py`), and — added in
later phases — search-result cache and search rate-limit state. Plain
`sqlite3` + small repository functions, no ORM, matching the rest of this
codebase.
"""
from __future__ import annotations

import hashlib
import json
import sqlite3
import threading
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from vendor_intel.config import _project_root

_DB_PATH = _project_root() / "data" / "batch.db"
_CACHE_DIR = _project_root() / "data" / "company_cache"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS batches (
    batch_id                TEXT PRIMARY KEY,
    input_file              TEXT,
    total_markets           INTEGER NOT NULL DEFAULT 0,
    max_concurrent_markets  INTEGER NOT NULL DEFAULT 1,
    created_at              TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);

CREATE TABLE IF NOT EXISTS markets (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id            TEXT NOT NULL,
    query               TEXT NOT NULL,
    country             TEXT NOT NULL DEFAULT 'global',
    market_slug         TEXT NOT NULL UNIQUE,
    market_type         TEXT NOT NULL DEFAULT 'general',
    sections_json       TEXT,
    seeds_file          TEXT,
    profile             TEXT NOT NULL DEFAULT 'quality',
    cap                 TEXT NOT NULL DEFAULT 'standard',
    priority            INTEGER NOT NULL DEFAULT 100,
    status              TEXT NOT NULL DEFAULT 'pending',
    worker_id           TEXT,
    attempt_count       INTEGER NOT NULL DEFAULT 0,
    max_attempts        INTEGER NOT NULL DEFAULT 3,
    last_error          TEXT,
    last_error_kind     TEXT,
    companies_exported  INTEGER,
    csv_path            TEXT,
    xlsx_path           TEXT,
    docx_path           TEXT,
    json_path           TEXT,
    queued_at           TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    started_at          TEXT,
    finished_at         TEXT,
    elapsed_seconds      REAL
);
CREATE INDEX IF NOT EXISTS idx_markets_status ON markets(status, priority);
CREATE INDEX IF NOT EXISTS idx_markets_batch ON markets(batch_id);

CREATE TABLE IF NOT EXISTS companies (
    domain                  TEXT PRIMARY KEY,
    canonical_name          TEXT,
    first_seen_market_id    INTEGER REFERENCES markets(id),
    last_crawled_at         TEXT,
    crawl_source            TEXT,
    crawl_payload_path      TEXT,
    crawl_success           INTEGER NOT NULL DEFAULT 0,
    times_seen              INTEGER NOT NULL DEFAULT 1,
    updated_at              TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
);

CREATE TABLE IF NOT EXISTS company_market_verdicts (
    market_id       INTEGER NOT NULL REFERENCES markets(id),
    domain          TEXT NOT NULL REFERENCES companies(domain),
    is_relevant     INTEGER NOT NULL,
    role            TEXT,
    confidence      REAL,
    quality_score   REAL,
    classified_at   TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    PRIMARY KEY (market_id, domain)
);

CREATE TABLE IF NOT EXISTS search_cache (
    query_norm      TEXT NOT NULL,
    geo             TEXT NOT NULL DEFAULT '',
    engine          TEXT NOT NULL,
    results_json    TEXT NOT NULL,
    fetched_at      TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    hit_count       INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (query_norm, geo, engine)
);
CREATE INDEX IF NOT EXISTS idx_search_cache_fetched ON search_cache(fetched_at);

CREATE TABLE IF NOT EXISTS search_rate_state (
    engine          TEXT PRIMARY KEY,
    tokens          REAL NOT NULL DEFAULT 0,
    last_refill_at  REAL NOT NULL DEFAULT 0,
    dead_until      REAL
);

"""

_local = threading.local()
_init_lock = threading.Lock()
_schema_ready = False


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def get_connection() -> sqlite3.Connection:
    """One connection per OS thread (sqlite3 connections aren't thread-safe).

    Safe to call from the asyncio event-loop thread and from asyncio.to_thread
    worker threads alike — each gets its own connection to the same WAL-mode
    database file.
    """
    global _schema_ready
    conn = getattr(_local, "conn", None)
    if conn is not None:
        return conn

    _DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(_DB_PATH), timeout=10.0, isolation_level=None)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA busy_timeout=5000")
    conn.execute("PRAGMA foreign_keys=ON")

    with _init_lock:
        if not _schema_ready:
            conn.executescript(_SCHEMA)
            _schema_ready = True

    _local.conn = conn
    return conn


def reset_for_tests() -> None:
    """Force the next get_connection() to re-run schema setup (test helper)."""
    global _schema_ready
    _schema_ready = False
    conn = getattr(_local, "conn", None)
    if conn is not None:
        conn.close()
        _local.conn = None


# --------------------------------------------------------------------------
# Market registry
# --------------------------------------------------------------------------


def create_batch(
    batch_id: str,
    *,
    input_file: str = "",
    total_markets: int = 0,
) -> None:
    """Create (or refresh the bookkeeping fields of) a batch row."""
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO batches (batch_id, input_file, total_markets)
        VALUES (?, ?, ?)
        ON CONFLICT(batch_id) DO UPDATE SET
            input_file = excluded.input_file,
            total_markets = excluded.total_markets
        """,
        (batch_id, input_file, total_markets),
    )


def upsert_market(
    batch_id: str,
    *,
    query: str,
    country: str,
    market_slug: str,
    market_type: str = "general",
    sections: list[str] | None = None,
    seeds_file: str | None = None,
    profile: str = "quality",
    cap: str = "standard",
    priority: int = 100,
) -> tuple[int, bool]:
    """Insert a market row, or update its definition in place if `market_slug`
    already exists. Never touches status/attempt_count on an update, so
    re-loading the same input file resumes a batch instead of resetting it.

    Returns (market_id, was_newly_inserted).
    """
    conn = get_connection()
    sections_json = json.dumps(list(sections or []))
    existing = conn.execute(
        "SELECT id FROM markets WHERE market_slug = ?", (market_slug,)
    ).fetchone()
    if existing is not None:
        conn.execute(
            """
            UPDATE markets SET
                batch_id = ?, query = ?, country = ?, market_type = ?, sections_json = ?,
                seeds_file = ?, profile = ?, cap = ?, priority = ?
            WHERE market_slug = ?
            """,
            (
                batch_id,
                query,
                country,
                market_type,
                sections_json,
                seeds_file,
                profile,
                cap,
                priority,
                market_slug,
            ),
        )
        return int(existing["id"]), False

    cur = conn.execute(
        """
        INSERT INTO markets (
            batch_id, query, country, market_slug, market_type, sections_json,
            seeds_file, profile, cap, priority
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (batch_id, query, country, market_slug, market_type, sections_json,
         seeds_file, profile, cap, priority),
    )
    return int(cur.lastrowid), True


def list_markets(batch_id: str) -> list[sqlite3.Row]:
    conn = get_connection()
    return conn.execute(
        "SELECT * FROM markets WHERE batch_id = ? ORDER BY priority, id", (batch_id,)
    ).fetchall()


def get_market_by_slug(market_slug: str) -> sqlite3.Row | None:
    conn = get_connection()
    return conn.execute(
        "SELECT * FROM markets WHERE market_slug = ?", (market_slug,)
    ).fetchone()


def batch_status_counts(batch_id: str) -> dict[str, int]:
    conn = get_connection()
    rows = conn.execute(
        "SELECT status, COUNT(*) AS n FROM markets WHERE batch_id = ? GROUP BY status",
        (batch_id,),
    ).fetchall()
    return {r["status"]: int(r["n"]) for r in rows}


def set_batch_concurrency(batch_id: str, max_concurrent_markets: int) -> None:
    conn = get_connection()
    conn.execute(
        "UPDATE batches SET max_concurrent_markets = ? WHERE batch_id = ?",
        (max_concurrent_markets, batch_id),
    )


def batch_progress_summary(batch_id: str) -> dict[str, Any]:
    """Aggregate numbers for status_batch.py: totals, average per-market
    wall-clock time (from markets already 'done' in this batch), and the
    concurrency the batch was last run with (for the ETA estimate).
    """
    conn = get_connection()
    row = conn.execute(
        """
        SELECT
            COALESCE(SUM(companies_exported), 0) AS companies_total,
            AVG(CASE WHEN status = 'done' THEN elapsed_seconds END) AS avg_elapsed
        FROM markets WHERE batch_id = ?
        """,
        (batch_id,),
    ).fetchone()
    batch_row = conn.execute(
        "SELECT max_concurrent_markets FROM batches WHERE batch_id = ?", (batch_id,)
    ).fetchone()
    concurrency = int(batch_row["max_concurrent_markets"]) if batch_row else 1
    avg_elapsed = row["avg_elapsed"]
    return {
        "companies_exported_total": int(row["companies_total"]),
        "avg_elapsed_seconds": float(avg_elapsed) if avg_elapsed is not None else None,
        "concurrency": max(1, concurrency),
    }


def list_recent_failures(batch_id: str, *, limit: int = 5) -> list[sqlite3.Row]:
    conn = get_connection()
    return conn.execute(
        """
        SELECT query, country, market_slug, last_error, last_error_kind, finished_at
        FROM markets
        WHERE batch_id = ? AND status = 'failed'
        ORDER BY finished_at DESC
        LIMIT ?
        """,
        (batch_id, limit),
    ).fetchall()


def claim_next_market(batch_id: str, worker_id: str) -> sqlite3.Row | None:
    """Atomically claim the highest-priority pending market for `worker_id`.

    Single UPDATE...WHERE id=(subquery)...RETURNING statement — safe under
    concurrent callers (multiple asyncio tasks or OS processes) sharing this
    WAL-mode database: SQLite serializes the writes, so two callers can never
    claim the same row (see the plan's "Worker pool: how markets get divided").
    """
    conn = get_connection()
    row = conn.execute(
        """
        UPDATE markets
        SET status = 'running', worker_id = ?, started_at = ?,
            attempt_count = attempt_count + 1
        WHERE id = (
            SELECT id FROM markets
            WHERE batch_id = ? AND status = 'pending'
            ORDER BY priority, id
            LIMIT 1
        )
        RETURNING *
        """,
        (worker_id, _now_iso(), batch_id),
    ).fetchone()
    return row


def mark_market_done(
    market_id: int,
    *,
    companies_exported: int,
    csv_path: str = "",
    xlsx_path: str = "",
    docx_path: str = "",
    json_path: str = "",
) -> None:
    conn = get_connection()
    now = _now_iso()
    conn.execute(
        """
        UPDATE markets SET
            status = 'done', finished_at = ?, companies_exported = ?,
            csv_path = ?, xlsx_path = ?, docx_path = ?, json_path = ?,
            elapsed_seconds = CASE WHEN started_at IS NOT NULL
                THEN (julianday(?) - julianday(started_at)) * 86400.0
                ELSE elapsed_seconds END
        WHERE id = ?
        """,
        (now, companies_exported, csv_path, xlsx_path, docx_path, json_path, now, market_id),
    )


def mark_market_failed(market_id: int, *, error: str, kind: str = "transient") -> None:
    """Requeue transient failures (up to `max_attempts`); park permanent ones as 'failed'.

    NOTE: callers pass kind='transient' by default until Phase 3's
    failure_classifier.py lands — every error is retried up to max_attempts
    rather than immediately parked, which is the safer default.
    """
    conn = get_connection()
    now = _now_iso()
    row = conn.execute(
        "SELECT attempt_count, max_attempts FROM markets WHERE id = ?", (market_id,)
    ).fetchone()
    attempt_count = int(row["attempt_count"]) if row else 1
    max_attempts = int(row["max_attempts"]) if row and row["max_attempts"] else 3

    if kind == "transient" and attempt_count < max_attempts:
        conn.execute(
            """
            UPDATE markets SET
                status = 'pending', worker_id = NULL,
                last_error = ?, last_error_kind = ?
            WHERE id = ?
            """,
            (error[:2000], kind, market_id),
        )
    else:
        conn.execute(
            """
            UPDATE markets SET
                status = 'failed', finished_at = ?,
                last_error = ?, last_error_kind = ?
            WHERE id = ?
            """,
            (now, error[:2000], kind, market_id),
        )


def reap_stale_running(batch_id: str, *, older_than_seconds: float = 0.0) -> int:
    """Requeue markets stuck in 'running' from a prior crashed process.

    Default ``older_than_seconds=0`` reaps every running row at startup so a
    killed ``run_batch`` (Cursor crash / OOM) can resume immediately. Callers
    that need a grace window can pass a positive age.
    """
    conn = get_connection()
    if older_than_seconds and older_than_seconds > 0:
        cutoff = datetime.now(timezone.utc) - timedelta(seconds=older_than_seconds)
        cur = conn.execute(
            """
            UPDATE markets
            SET status = 'pending', worker_id = NULL,
                last_error = 'reaped: stale running market at startup',
                last_error_kind = 'transient'
            WHERE batch_id = ? AND status = 'running' AND started_at < ?
            """,
            (batch_id, cutoff.isoformat()),
        )
    else:
        cur = conn.execute(
            """
            UPDATE markets
            SET status = 'pending', worker_id = NULL,
                last_error = 'reaped: stale running market at startup',
                last_error_kind = 'transient'
            WHERE batch_id = ? AND status = 'running'
            """,
            (batch_id,),
        )
    return cur.rowcount


def requeue_failed_markets(batch_id: str) -> int:
    """Move parked ``failed`` markets back to ``pending`` for another attempt."""
    conn = get_connection()
    cur = conn.execute(
        """
        UPDATE markets
        SET status = 'pending', worker_id = NULL, finished_at = NULL,
            last_error = COALESCE(last_error, 'requeued failed market'),
            last_error_kind = 'transient'
        WHERE batch_id = ? AND status = 'failed'
        """,
        (batch_id,),
    )
    return cur.rowcount


# --------------------------------------------------------------------------
# Cross-market company crawl cache
# --------------------------------------------------------------------------


def _cache_payload_path(domain: str) -> Path:
    digest = hashlib.sha1(domain.encode("utf-8")).hexdigest()[:24]
    return _CACHE_DIR / f"{digest}.json"


def get_cached_company(
    domain: str,
    *,
    max_age_days: float | None = 30.0,
) -> dict[str, Any] | None:
    """Return a previously cached crawl payload for `domain`, or None on a
    cache miss / stale entry / unreadable payload file."""
    dom = (domain or "").strip().lower()
    if not dom:
        return None

    conn = get_connection()
    row = conn.execute(
        "SELECT crawl_payload_path, crawl_success, updated_at "
        "FROM companies WHERE domain = ?",
        (dom,),
    ).fetchone()
    if row is None or not row["crawl_success"] or not row["crawl_payload_path"]:
        return None

    if max_age_days is not None:
        try:
            updated = datetime.fromisoformat(str(row["updated_at"]))
        except ValueError:
            return None
        if updated.tzinfo is None:
            updated = updated.replace(tzinfo=timezone.utc)
        if datetime.now(timezone.utc) - updated > timedelta(days=max_age_days):
            return None

    path = Path(row["crawl_payload_path"])
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def upsert_company_cache(
    domain: str,
    payload: dict[str, Any],
    *,
    canonical_name: str = "",
    crawl_source: str = "",
    success: bool = True,
    market_id: int | None = None,
) -> None:
    """Persist a crawl payload for `domain` so later markets can reuse it."""
    dom = (domain or "").strip().lower()
    if not dom:
        return

    # get_connection() ensures _CACHE_DIR exists — must run before the file write below.
    conn = get_connection()
    path = _cache_payload_path(dom)
    try:
        path.write_text(json.dumps(payload), encoding="utf-8")
    except (OSError, TypeError, ValueError):
        return

    now = _now_iso()
    conn.execute(
        """
        INSERT INTO companies (
            domain, canonical_name, first_seen_market_id, last_crawled_at,
            crawl_source, crawl_payload_path, crawl_success, times_seen, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)
        ON CONFLICT(domain) DO UPDATE SET
            canonical_name = CASE WHEN excluded.canonical_name != ''
                                   THEN excluded.canonical_name ELSE companies.canonical_name END,
            last_crawled_at = excluded.last_crawled_at,
            crawl_source = excluded.crawl_source,
            crawl_payload_path = excluded.crawl_payload_path,
            crawl_success = excluded.crawl_success,
            times_seen = companies.times_seen + 1,
            updated_at = excluded.updated_at
        """,
        (dom, canonical_name, market_id, now, crawl_source, str(path), int(success), now),
    )


def company_cache_stats() -> dict[str, int]:
    """Cheap counters for observability (Phase 4)."""
    conn = get_connection()
    row = conn.execute(
        "SELECT COUNT(*) AS n, COALESCE(SUM(times_seen), 0) AS total_hits, "
        "COALESCE(SUM(times_seen) - COUNT(*), 0) AS reused "
        "FROM companies WHERE crawl_success = 1"
    ).fetchone()
    return {
        "cached_companies": int(row["n"]),
        "total_sightings": int(row["total_hits"]),
        "cross_market_reuses": int(row["reused"]),
    }


# --------------------------------------------------------------------------
# Search-result cache + search rate-limit/circuit-breaker state (Phase 6)
# --------------------------------------------------------------------------


def _normalize_query(text: str) -> str:
    return " ".join((text or "").strip().lower().split())


def get_cached_search(
    query: str,
    geo: str,
    engine: str,
    *,
    max_age_days: float = 14.0,
) -> list[dict] | None:
    """Return cached search results for (query, geo, engine), or None on a
    cache miss / stale entry. Bumps hit_count on every hit for observability."""
    conn = get_connection()
    q_norm = _normalize_query(query)
    geo_norm = (geo or "").strip().lower()
    row = conn.execute(
        "SELECT results_json, fetched_at FROM search_cache "
        "WHERE query_norm = ? AND geo = ? AND engine = ?",
        (q_norm, geo_norm, engine),
    ).fetchone()
    if row is None:
        return None
    try:
        fetched = datetime.fromisoformat(str(row["fetched_at"]))
    except ValueError:
        return None
    if fetched.tzinfo is None:
        fetched = fetched.replace(tzinfo=timezone.utc)
    if datetime.now(timezone.utc) - fetched > timedelta(days=max_age_days):
        return None
    try:
        results = json.loads(row["results_json"])
    except (TypeError, json.JSONDecodeError):
        return None
    conn.execute(
        "UPDATE search_cache SET hit_count = hit_count + 1 "
        "WHERE query_norm = ? AND geo = ? AND engine = ?",
        (q_norm, geo_norm, engine),
    )
    return results


def put_cached_search(query: str, geo: str, engine: str, results: list[dict]) -> None:
    conn = get_connection()
    q_norm = _normalize_query(query)
    geo_norm = (geo or "").strip().lower()
    now = _now_iso()
    conn.execute(
        """
        INSERT INTO search_cache (query_norm, geo, engine, results_json, fetched_at, hit_count)
        VALUES (?, ?, ?, ?, ?, 0)
        ON CONFLICT(query_norm, geo, engine) DO UPDATE SET
            results_json = excluded.results_json,
            fetched_at = excluded.fetched_at
        """,
        (q_norm, geo_norm, engine, json.dumps(results), now),
    )


def search_cache_stats() -> dict[str, int]:
    conn = get_connection()
    row = conn.execute(
        "SELECT COUNT(*) AS n, COALESCE(SUM(hit_count), 0) AS hits FROM search_cache"
    ).fetchone()
    return {"cached_queries": int(row["n"]), "cache_hits": int(row["hits"])}


def checkpoint_search_rate_state(
    engine: str, *, tokens: float, dead_until: float | None
) -> None:
    """Persist one engine's rate-limiter state (rate_limiter.py calls this
    periodically, not on every request) so a killed/restarted batch doesn't
    forget a recent block."""
    conn = get_connection()
    conn.execute(
        """
        INSERT INTO search_rate_state (engine, tokens, last_refill_at, dead_until)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(engine) DO UPDATE SET
            tokens = excluded.tokens,
            last_refill_at = excluded.last_refill_at,
            dead_until = excluded.dead_until
        """,
        (engine, tokens, time.time(), dead_until),
    )


def load_search_rate_state() -> dict[str, float | None]:
    """engine -> dead_until (unix epoch seconds, or None if not/no-longer dead).
    Called once at batch startup to hydrate rate_limiter.py's in-memory state."""
    conn = get_connection()
    rows = conn.execute("SELECT engine, dead_until FROM search_rate_state").fetchall()
    return {r["engine"]: (float(r["dead_until"]) if r["dead_until"] is not None else None) for r in rows}
