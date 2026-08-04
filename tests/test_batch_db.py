"""Batch registry + cross-market company cache (src/vendor_intel/batch/db.py)."""
from __future__ import annotations

import asyncio

import pytest

from vendor_intel.batch import db


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    """Point the module at a throwaway DB/cache dir for every test."""
    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    yield
    db.reset_for_tests()


def test_schema_creates_wal_mode_db():
    conn = db.get_connection()
    mode = conn.execute("PRAGMA journal_mode").fetchone()[0]
    assert mode.lower() == "wal"
    tables = {
        r[0]
        for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()
    }
    assert {"markets", "companies", "company_market_verdicts", "search_cache",
            "search_rate_state", "batches"} <= tables


def test_cache_miss_on_empty_db():
    assert db.get_cached_company("example.com") is None


def test_cache_roundtrip():
    payload = {"domain": "example.com", "data": {"company": {"name": "Example Inc"}}}
    db.upsert_company_cache(
        "example.com", payload, canonical_name="Example Inc", crawl_source="smart_crawl"
    )
    assert db.get_cached_company("example.com") == payload


def test_cross_market_reuse_counter():
    payload = {"domain": "acme.com", "data": {}}
    db.upsert_company_cache("acme.com", payload, canonical_name="Acme")
    db.upsert_company_cache("acme.com", payload, canonical_name="Acme")  # 2nd market "sees" it
    stats = db.company_cache_stats()
    assert stats["cached_companies"] == 1
    assert stats["cross_market_reuses"] == 1


def test_stale_entry_is_a_miss():
    payload = {"domain": "old.com", "data": {}}
    db.upsert_company_cache("old.com", payload)
    assert db.get_cached_company("old.com", max_age_days=0) is None
    assert db.get_cached_company("old.com", max_age_days=365) == payload


def test_domain_normalization_is_case_insensitive():
    payload = {"domain": "mixedcase.com", "data": {}}
    db.upsert_company_cache("MixedCase.com", payload)
    assert db.get_cached_company("mixedcase.COM") == payload


def test_missing_payload_file_is_a_miss(tmp_path):
    payload = {"domain": "vanished.com", "data": {}}
    db.upsert_company_cache("vanished.com", payload)
    path = db._cache_payload_path("vanished.com")
    path.unlink()
    assert db.get_cached_company("vanished.com") is None


def _seed_market(batch_id="b1", query="Widget Market", country="Germany", **kw):
    db.create_batch(batch_id, input_file="test.csv", total_markets=1)
    market_id, was_new = db.upsert_market(
        batch_id,
        query=query,
        country=country,
        market_slug=kw.pop("market_slug", "widget_market_germany"),
        **kw,
    )
    return market_id, was_new


def test_upsert_market_inserts_then_updates_in_place():
    market_id1, was_new1 = _seed_market(priority=100)
    assert was_new1 is True

    market_id2, was_new2 = _seed_market(priority=5, market_type="industrial")
    assert was_new2 is False
    assert market_id2 == market_id1

    row = db.get_market_by_slug("widget_market_germany")
    assert row["priority"] == 5
    assert row["market_type"] == "industrial"
    assert row["status"] == "pending"  # unaffected by the update


def test_upsert_market_preserves_progress_on_reload():
    market_id, _ = _seed_market()
    db.claim_next_market("b1", "worker-1")
    db.mark_market_done(market_id, companies_exported=12, csv_path="x.csv")

    # Re-loading the same input (e.g. a resumed batch) must not reset status.
    _seed_market(priority=1)
    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "done"
    assert row["companies_exported"] == 12
    assert row["priority"] == 1


def test_claim_next_market_picks_highest_priority_first():
    db.create_batch("b1")
    db.upsert_market("b1", query="Low prio", country="global", market_slug="low", priority=200)
    db.upsert_market("b1", query="High prio", country="global", market_slug="high", priority=1)

    claimed = db.claim_next_market("b1", "worker-1")
    assert claimed["market_slug"] == "high"
    assert claimed["status"] == "running"
    assert claimed["worker_id"] == "worker-1"
    assert claimed["attempt_count"] == 1


def test_claim_next_market_never_double_claims():
    db.create_batch("b1")
    db.upsert_market("b1", query="Only one", country="global", market_slug="only", priority=1)

    first = db.claim_next_market("b1", "worker-1")
    second = db.claim_next_market("b1", "worker-2")
    assert first is not None
    assert second is None  # already claimed, nothing left pending


def test_claim_next_market_empty_queue_returns_none():
    db.create_batch("b1")
    assert db.claim_next_market("b1", "worker-1") is None


def test_mark_market_done_records_outputs_and_elapsed():
    market_id, _ = _seed_market()
    db.claim_next_market("b1", "worker-1")
    db.mark_market_done(market_id, companies_exported=7, csv_path="a.csv", xlsx_path="a.xlsx")

    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "done"
    assert row["companies_exported"] == 7
    assert row["csv_path"] == "a.csv"
    assert row["xlsx_path"] == "a.xlsx"
    assert row["finished_at"] is not None
    assert row["elapsed_seconds"] is not None
    assert row["elapsed_seconds"] >= 0


def test_mark_market_failed_requeues_transient_up_to_max_attempts():
    market_id, _ = _seed_market()
    for _ in range(3):
        claimed = db.claim_next_market("b1", "worker-1")
        assert claimed is not None, "expected the market to be requeued as pending"
        db.mark_market_failed(market_id, error="boom", kind="transient")

    row = db.get_market_by_slug("widget_market_germany")
    # attempt_count hit max_attempts (default 3) on the 3rd claim -> parked as failed.
    assert row["status"] == "failed"
    assert row["attempt_count"] == 3
    assert row["last_error"] == "boom"

    # No further claims — the queue is empty, not stuck retrying forever.
    assert db.claim_next_market("b1", "worker-1") is None


def test_mark_market_failed_permanent_parks_immediately():
    market_id, _ = _seed_market()
    db.claim_next_market("b1", "worker-1")
    db.mark_market_failed(market_id, error="bad config", kind="permanent")

    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "failed"
    assert row["attempt_count"] == 1


def test_batch_status_counts():
    db.create_batch("b1")
    db.upsert_market("b1", query="A", country="global", market_slug="a")
    db.upsert_market("b1", query="B", country="global", market_slug="b")
    claimed = db.claim_next_market("b1", "worker-1")
    counts = db.batch_status_counts("b1")
    assert counts == {"pending": 1, "running": 1}


def test_reap_stale_running_requeues_old_markets():
    market_id, _ = _seed_market()
    db.claim_next_market("b1", "worker-1")
    conn = db.get_connection()
    conn.execute(
        "UPDATE markets SET started_at = '2000-01-01T00:00:00+00:00' WHERE id = ?",
        (market_id,),
    )
    reaped = db.reap_stale_running("b1", older_than_seconds=60)
    assert reaped == 1
    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "pending"
    assert row["worker_id"] is None


def test_set_batch_concurrency_and_progress_summary_default():
    db.create_batch("b1")
    summary = db.batch_progress_summary("b1")
    assert summary == {
        "companies_exported_total": 0,
        "avg_elapsed_seconds": None,
        "concurrency": 1,
    }

    db.set_batch_concurrency("b1", 4)
    summary = db.batch_progress_summary("b1")
    assert summary["concurrency"] == 4


def test_batch_progress_summary_aggregates_done_markets():
    db.create_batch("b1")
    m1, _ = db.upsert_market("b1", query="A", country="global", market_slug="a")
    m2, _ = db.upsert_market("b1", query="B", country="global", market_slug="b")
    db.claim_next_market("b1", "w1")
    db.claim_next_market("b1", "w1")
    db.mark_market_done(m1, companies_exported=10, csv_path="a.csv")
    db.mark_market_done(m2, companies_exported=20, csv_path="b.csv")

    summary = db.batch_progress_summary("b1")
    assert summary["companies_exported_total"] == 30
    assert summary["avg_elapsed_seconds"] is not None
    assert summary["avg_elapsed_seconds"] >= 0


def test_list_recent_failures_only_returns_failed():
    db.create_batch("b1")
    m1, _ = db.upsert_market("b1", query="Good", country="global", market_slug="good")
    m2, _ = db.upsert_market("b1", query="Bad", country="global", market_slug="bad")
    db.claim_next_market("b1", "w1")
    db.claim_next_market("b1", "w1")
    db.mark_market_done(m1, companies_exported=5)
    db.mark_market_failed(m2, error="explosion", kind="permanent")

    failures = db.list_recent_failures("b1")
    assert len(failures) == 1
    assert failures[0]["query"] == "Bad"
    assert failures[0]["last_error"] == "explosion"
    assert failures[0]["last_error_kind"] == "permanent"


def test_list_recent_failures_respects_limit():
    db.create_batch("b1")
    for i in range(7):
        market_id, _ = db.upsert_market("b1", query=f"M{i}", country="global", market_slug=f"m{i}")
        db.claim_next_market("b1", "w1")
        db.mark_market_failed(market_id, error="x", kind="permanent")

    assert len(db.list_recent_failures("b1", limit=3)) == 3
    assert len(db.list_recent_failures("b1", limit=100)) == 7


def test_search_cache_miss_then_roundtrip():
    assert db.get_cached_search("rupture disc manufacturers", "Brazil", "ddgs") is None

    results = [{"title": "Acme", "link": "https://acme.com", "snippet": "x", "backend": "ddgs"}]
    db.put_cached_search("Rupture Disc Manufacturers", "Brazil", "ddgs", results)

    # Case/whitespace-insensitive query match.
    got = db.get_cached_search("  rupture   disc MANUFACTURERS ", "Brazil", "ddgs")
    assert got == results


def test_search_cache_scoped_by_geo_and_engine():
    results = [{"title": "A"}]
    db.put_cached_search("widget market", "Brazil", "ddgs", results)
    assert db.get_cached_search("widget market", "Germany", "ddgs") is None
    assert db.get_cached_search("widget market", "Brazil", "bing_html") is None
    assert db.get_cached_search("widget market", "Brazil", "ddgs") == results


def test_search_cache_stale_entry_is_a_miss():
    db.put_cached_search("widget market", "Brazil", "ddgs", [{"title": "A"}])
    assert db.get_cached_search("widget market", "Brazil", "ddgs", max_age_days=0) is None
    assert db.get_cached_search("widget market", "Brazil", "ddgs", max_age_days=365) is not None


def test_search_cache_hit_count_increments():
    db.put_cached_search("widget market", "Brazil", "ddgs", [{"title": "A"}])
    db.get_cached_search("widget market", "Brazil", "ddgs")
    db.get_cached_search("widget market", "Brazil", "ddgs")
    stats = db.search_cache_stats()
    assert stats["cached_queries"] == 1
    assert stats["cache_hits"] == 2


def test_checkpoint_and_load_search_rate_state():
    assert db.load_search_rate_state() == {}
    db.checkpoint_search_rate_state("ddgs", tokens=3.5, dead_until=None)
    db.checkpoint_search_rate_state("bing_html", tokens=0.0, dead_until=9999999999.0)

    state = db.load_search_rate_state()
    assert state["ddgs"] is None
    assert state["bing_html"] == 9999999999.0


def test_checkpoint_search_rate_state_upserts():
    db.checkpoint_search_rate_state("ddgs", tokens=1.0, dead_until=None)
    db.checkpoint_search_rate_state("ddgs", tokens=2.0, dead_until=123.0)
    state = db.load_search_rate_state()
    assert state == {"ddgs": 123.0}


def test_enrichment_skips_recrawl_across_markets(monkeypatch):
    """The actual value proposition: a company seen in market A is not
    re-crawled when it shows up again in market B."""
    from vendor_intel.enrichment import smart_enrichment as se

    se.clear_enrichment_cache()
    calls = {"n": 0}

    async def fake_ssc_only(dom):
        calls["n"] += 1
        return {"domain": dom, "source": "ssc", "pages": [{"url": f"https://{dom}", "text": "x" * 200}]}

    monkeypatch.setattr(se, "_ssc_only", fake_ssc_only)

    async def run():
        name1, data1, err1 = await se._crawl_one("Acme Corp", "acme.com", use_ssc=True)
        assert err1 is None
        assert calls["n"] == 1

        # New "market run": fresh in-memory cache, same domain.
        se.clear_enrichment_cache()
        name2, data2, err2 = await se._crawl_one("Acme Corp (seen again)", "acme.com", use_ssc=True)
        assert err2 is None
        assert calls["n"] == 1, "expected cross-market cache hit, crawler was called again"
        assert data2 == data1

    asyncio.run(run())
