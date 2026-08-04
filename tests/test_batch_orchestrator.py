"""Concurrent worker pool (src/vendor_intel/batch/orchestrator.py) — proves
markets actually overlap in time under max_concurrent_markets > 1, that
concurrency scales down per-market enrich/classify concurrency sanely, and
that every market still ends up processed exactly once (no double-claims,
no drops) regardless of concurrency level. All pipeline calls are mocked —
zero real network/LLM calls.
"""
from __future__ import annotations

import asyncio

import pytest

from vendor_intel.batch import db, orchestrator, registry_loader
from vendor_intel.config import Settings


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    yield
    db.reset_for_tests()


def _settings() -> Settings:
    return Settings.load().model_copy(
        update={"pipeline_profile": "quality", "pipeline_enrich_concurrent": 12, "pipeline_classify_concurrent": 12}
    )


def test_scale_settings_for_concurrency_divides_down():
    s = _settings()
    scaled = orchestrator.scale_settings_for_concurrency(s, 4)
    assert scaled.pipeline_enrich_concurrent == 3   # 12 // 4
    assert scaled.pipeline_classify_concurrent == 3


def test_scale_settings_for_concurrency_floors_at_two():
    s = _settings()
    scaled = orchestrator.scale_settings_for_concurrency(s, 20)  # 12 // 20 == 0
    assert scaled.pipeline_enrich_concurrent == 2
    assert scaled.pipeline_classify_concurrent == 2


def test_scale_settings_for_concurrency_noop_at_one():
    s = _settings()
    scaled = orchestrator.scale_settings_for_concurrency(s, 1)
    assert scaled is s  # unchanged object, not just equal values


def test_run_batch_async_processes_every_market_exactly_once(tmp_path, monkeypatch):
    p = tmp_path / "markets.csv"
    p.write_text(
        "query,country,priority\n"
        + "".join(f"Market {i},Global,{i}\n" for i in range(10)),
        encoding="utf-8",
    )
    registry_loader.populate_registry(p, "b1")

    processed = []

    async def fake_execute_and_export_async(query, country, settings, output_dir, **kw):
        processed.append(query)
        await asyncio.sleep(0)
        return {"relevant_companies": [], "_csv_path": "", "_json_path": ""}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake_execute_and_export_async)

    counts = asyncio.run(
        orchestrator.run_batch_async("b1", _settings(), max_concurrent_markets=4)
    )
    assert counts == {"done": 10}
    assert sorted(processed) == sorted(f"Market {i}" for i in range(10))
    assert len(processed) == 10  # no duplicates, nothing dropped


def test_markets_actually_overlap_under_concurrency(tmp_path, monkeypatch):
    """The real point of Phase 3: prove more than one market runs AT THE SAME
    TIME, not just that the queue eventually drains."""
    p = tmp_path / "markets.csv"
    p.write_text(
        "query,country,priority\n" + "".join(f"Market {i},Global,{i}\n" for i in range(6)),
        encoding="utf-8",
    )
    registry_loader.populate_registry(p, "b1")

    in_flight = {"current": 0, "max_seen": 0}

    async def fake_execute_and_export_async(query, country, settings, output_dir, **kw):
        in_flight["current"] += 1
        in_flight["max_seen"] = max(in_flight["max_seen"], in_flight["current"])
        await asyncio.sleep(0.05)  # hold the "slot" long enough for others to join in
        in_flight["current"] -= 1
        return {"relevant_companies": [], "_csv_path": "", "_json_path": ""}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake_execute_and_export_async)

    counts = asyncio.run(
        orchestrator.run_batch_async("b1", _settings(), max_concurrent_markets=3)
    )
    assert counts == {"done": 6}
    assert in_flight["max_seen"] >= 2, (
        f"expected multiple markets in flight simultaneously, only saw {in_flight['max_seen']}"
    )


def test_run_batch_async_respects_limit(tmp_path, monkeypatch):
    p = tmp_path / "markets.csv"
    p.write_text(
        "query,country,priority\n" + "".join(f"Market {i},Global,{i}\n" for i in range(5)),
        encoding="utf-8",
    )
    registry_loader.populate_registry(p, "b1")

    async def fake_execute_and_export_async(query, country, settings, output_dir, **kw):
        return {"relevant_companies": [], "_csv_path": "", "_json_path": ""}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake_execute_and_export_async)

    counts = asyncio.run(
        orchestrator.run_batch_async("b1", _settings(), max_concurrent_markets=2, limit=2)
    )
    assert counts == {"done": 2, "pending": 3}


def test_run_batch_async_handles_mixed_success_and_failure(tmp_path, monkeypatch):
    p = tmp_path / "markets.csv"
    p.write_text(
        "query,country,priority\n"
        "Good Market,Global,1\n"
        "Bad Market,Global,2\n",
        encoding="utf-8",
    )
    registry_loader.populate_registry(p, "b1")

    async def fake_execute_and_export_async(query, country, settings, output_dir, **kw):
        if query == "Bad Market":
            raise ValueError("bad config for this market")
        return {"relevant_companies": [{"company": "X"}], "_csv_path": "c.csv", "_json_path": "c.json"}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake_execute_and_export_async)

    counts = asyncio.run(
        orchestrator.run_batch_async("b1", _settings(), max_concurrent_markets=2)
    )
    assert counts == {"done": 1, "failed": 1}  # ValueError -> permanent -> parked, not retried

    bad = db.get_market_by_slug("bad_market_global")
    assert bad["last_error"] == "ValueError: bad config for this market"
    assert bad["last_error_kind"] == "permanent"


def test_run_batch_async_reaps_stale_running_on_start(tmp_path, monkeypatch):
    p = tmp_path / "markets.csv"
    p.write_text("query,country\nStale Market,Global\n", encoding="utf-8")
    registry_loader.populate_registry(p, "b1")

    # Simulate a market left 'running' by a crashed prior process.
    market = db.claim_next_market("b1", "dead-worker")
    conn = db.get_connection()
    conn.execute(
        "UPDATE markets SET started_at = '2000-01-01T00:00:00+00:00' WHERE id = ?",
        (market["id"],),
    )

    async def fake_execute_and_export_async(query, country, settings, output_dir, **kw):
        return {"relevant_companies": [], "_csv_path": "", "_json_path": ""}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake_execute_and_export_async)

    counts = asyncio.run(orchestrator.run_batch_async("b1", _settings(), max_concurrent_markets=2))
    assert counts == {"done": 1}
