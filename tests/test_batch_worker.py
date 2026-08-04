"""Direct tests of batch/worker.py's process_market_async — the single
"run one claimed market" function shared by sequential and concurrent paths.
"""
from __future__ import annotations

import asyncio

import pytest

from vendor_intel.batch import db, registry_loader, worker
from vendor_intel.config import Settings


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    yield
    db.reset_for_tests()


def _claimed_market(tmp_path, query="Widget Market", country="Germany"):
    p = tmp_path / "markets.txt"
    p.write_text(f"{query} | {country}\n", encoding="utf-8")
    registry_loader.populate_registry(p, "b1")
    return db.claim_next_market("b1", "worker-1")


def _settings() -> Settings:
    return Settings.load().model_copy(update={"pipeline_profile": "quality"})


def test_process_market_async_success(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)

    async def fake(query, country, settings, output_dir, **kw):
        assert query == "Widget Market"
        assert country == "Germany"
        assert str(output_dir).endswith("widget_market_germany")
        return {
            "relevant_companies": [{"company": "A"}, {"company": "B"}, {"company": "C"}],
            "_csv_path": str(output_dir / "widget_market_germany.csv"),
            "_xlsx_path": str(output_dir / "widget_market_germany.xlsx"),
        }

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=True))

    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "done"
    assert row["companies_exported"] == 3
    assert row["xlsx_path"].endswith(".xlsx")


def test_process_market_async_transient_failure_requeues(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)

    async def fake(*a, **kw):
        raise TimeoutError("Read timed out")

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=False))

    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "pending"
    assert row["last_error_kind"] == "transient"


def test_process_market_async_permanent_failure_parks(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)

    async def fake(*a, **kw):
        raise ValueError("Provide a query")

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=False))

    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "failed"
    assert row["last_error_kind"] == "permanent"


def test_process_market_async_passes_market_type_through_to_execute_and_export(tmp_path, monkeypatch):
    """market_type (from the registry row) must reach execute_and_export_async
    so the Excel export uses the right column schema for this market's vertical."""
    db.create_batch("b1")
    db.upsert_market(
        "b1", query="Laser Market", country="Global", market_slug="laser",
        market_type="security_tech",
    )
    market = db.claim_next_market("b1", "worker-1")
    seen = {}

    async def fake(query, country, settings, output_dir, *, market_type="fmcg", **kw):
        seen["market_type"] = market_type
        return {"relevant_companies": []}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=False))
    assert seen["market_type"] == "security_tech"


def test_process_market_async_defaults_market_type_general_when_unset(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)  # .txt-loaded markets default to "general"
    seen = {}

    async def fake(query, country, settings, output_dir, *, market_type="fmcg", **kw):
        seen["market_type"] = market_type
        return {"relevant_companies": []}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=False))
    assert seen["market_type"] == "general"


def test_process_market_async_uses_market_specific_geo_limits(tmp_path, monkeypatch):
    """Distinct markets with different countries should each compute their own
    pipeline_limits(country=...), not one fixed value for the whole batch."""
    market = _claimed_market(tmp_path, query="Global Market", country="global")

    seen_limits = {}

    async def fake(query, country, settings, output_dir, *, enrich_limit, classify_limit, **kw):
        seen_limits["enrich_limit"] = enrich_limit
        seen_limits["classify_limit"] = classify_limit
        return {"relevant_companies": []}

    import vendor_intel.pipeline.run_and_export as rae
    monkeypatch.setattr(rae, "execute_and_export_async", fake)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=False))
    assert seen_limits["enrich_limit"] > 0
    assert seen_limits["classify_limit"] > 0


def test_process_market_async_full_enrich_calls_post_export(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)
    post_calls = []

    async def fake(query, country, settings, output_dir, **kw):
        output_dir.mkdir(parents=True, exist_ok=True)
        out = output_dir / "widget_market_germany.json"
        out.write_text('{"relevant_companies": []}', encoding="utf-8")
        return {
            "relevant_companies": [{"company": "Acme", "domain": "acme.com"}],
            "_json_path": str(out),
            "_csv_path": str(output_dir / "widget_market_germany.csv"),
            "_xlsx_path": str(output_dir / "widget_market_germany.xlsx"),
        }

    async def fake_post(result, **kw):
        post_calls.append(kw)
        result["post_enrich_ok"] = 1
        return result

    import vendor_intel.enrichment.post_export_enrich as pee
    import vendor_intel.pipeline.run_and_export as rae

    monkeypatch.setattr(rae, "execute_and_export_async", fake)
    monkeypatch.setattr(pee, "post_export_enrich_market", fake_post)

    asyncio.run(
        worker.process_market_async(
            market,
            _settings(),
            fmcg_xlsx=True,
            full_enrich=True,
            enrich_concurrent=2,
            search_backfill=True,
        )
    )

    assert len(post_calls) == 1
    assert post_calls[0]["concurrent"] == 2
    assert post_calls[0]["search_backfill"] is True
    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "done"


def test_process_market_async_skips_post_export_when_disabled(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)
    post_calls = []

    async def fake(*a, **kw):
        return {"relevant_companies": [], "_json_path": ""}

    async def fake_post(*a, **kw):
        post_calls.append(1)
        return {}

    import vendor_intel.enrichment.post_export_enrich as pee
    import vendor_intel.pipeline.run_and_export as rae

    monkeypatch.setattr(rae, "execute_and_export_async", fake)
    monkeypatch.setattr(pee, "post_export_enrich_market", fake_post)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=False, full_enrich=False))
    assert post_calls == []


def test_process_market_async_with_audit_calls_market_audit(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)
    audit_calls = []

    async def fake(query, country, settings, output_dir, **kw):
        output_dir.mkdir(parents=True, exist_ok=True)
        out = output_dir / "widget_market_germany.json"
        csv = output_dir / "widget_market_germany.csv"
        out.write_text('{"relevant_companies": [{"company": "Acme"}]}', encoding="utf-8")
        csv.write_text("Company,Website\nAcme,acme.com\n", encoding="utf-8")
        return {
            "relevant_companies": [{"company": "Acme", "domain": "acme.com"}],
            "_json_path": str(out),
            "_csv_path": str(csv),
        }

    async def fake_audit(**kw):
        audit_calls.append(kw)
        return {"keep": 1, "review": 0, "remove": 0, "total": 1}

    import vendor_intel.enrichment.market_audit as ma
    import vendor_intel.pipeline.run_and_export as rae

    monkeypatch.setattr(rae, "execute_and_export_async", fake)
    monkeypatch.setattr(ma, "audit_market_exports", fake_audit)

    asyncio.run(
        worker.process_market_async(
            market,
            _settings(),
            fmcg_xlsx=True,
            with_audit=True,
            audit_concurrent=3,
        )
    )

    assert len(audit_calls) == 1
    assert audit_calls[0]["market_query"] == "Widget Market"
    assert audit_calls[0]["concurrent"] == 3
    row = db.get_market_by_slug("widget_market_germany")
    assert row["status"] == "done"


def test_process_market_async_skips_audit_when_disabled(tmp_path, monkeypatch):
    market = _claimed_market(tmp_path)
    audit_calls = []

    async def fake(*a, **kw):
        return {"relevant_companies": [], "_json_path": ""}

    async def fake_audit(**kw):
        audit_calls.append(1)
        return {}

    import vendor_intel.enrichment.market_audit as ma
    import vendor_intel.pipeline.run_and_export as rae

    monkeypatch.setattr(rae, "execute_and_export_async", fake)
    monkeypatch.setattr(ma, "audit_market_exports", fake_audit)

    asyncio.run(worker.process_market_async(market, _settings(), fmcg_xlsx=False, with_audit=False))
    assert audit_calls == []
