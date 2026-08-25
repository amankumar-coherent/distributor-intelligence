"""Tests for automated gap-fill detection and batch wiring."""
from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    from vendor_intel.batch import db

    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    yield
    db.reset_for_tests()


@pytest.fixture
def run_batch_mod(monkeypatch):
    monkeypatch.syspath_prepend(str(ROOT / "src"))
    import importlib.util

    spec = importlib.util.spec_from_file_location("run_batch", ROOT / "run_batch.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _write_markets_csv(tmp_path) -> Path:
    p = tmp_path / "markets.csv"
    p.write_text("query,country\nMarket A,Global\nMarket B,Global\n", encoding="utf-8")
    return p


def _row(**crawl: str) -> dict:
    return {"company": "Acme Corp", "_crawl": dict(crawl)}


def test_row_has_gaps_detects_missing_fields():
    from vendor_intel.enrichment.gap_fill.gaps import missing_fields, row_has_gaps

    empty = _row()
    assert row_has_gaps(empty) is True
    gaps = missing_fields(empty)
    assert "headquarters" in gaps
    assert "linkedin" in gaps

    full = _row(
        headquarters="Boston, MA, USA",
        founded_year="2010",
        employee_count="500",
        ownership="Private",
        revenue="$10M",
        offices="Boston, London",
        regions_served="United States; Canada",
        website="https://acme.com",
        linkedin="https://linkedin.com/company/acme",
        core_categories="Widgets; Parts",
        specialty_focus="Industrial distribution",
        key_brands_represented="BrandA; BrandB",
        distribution_type="Distributor",
        retail_ecommerce="Both",
        contact_person="Jane Doe",
        contact_role="CEO",
        email="jane@acme.com",
        phone="+1-555-0100",
        summary="Acme Corp is an industrial distributor serving North America.",
    )
    assert row_has_gaps(full) is False
    assert missing_fields(full) == []


def test_merge_extracted_gap_fill_only():
    from vendor_intel.enrichment.gap_fill.gaps import has_headquarters, merge_extracted

    row = _row(headquarters="Existing HQ, USA")
    extracted = {"headquarters": "New HQ, USA", "founded_year": "2015", "linkedin_url": "linkedin.com/company/acme"}
    updates = merge_extracted(row, extracted, source="test")
    assert "headquarters" not in updates
    assert "founded_year" in updates
    assert "linkedin" in updates
    assert has_headquarters(row)
    assert row["_crawl"]["headquarters"] == "Existing HQ, USA"


def test_plausible_founded_year_rejects_garbage():
    from vendor_intel.enrichment.gap_fill.gaps import merge_extracted, plausible_founded_year

    assert plausible_founded_year("2015") == 2015
    assert plausible_founded_year("founded in 1988") == 1988
    assert plausible_founded_year("1200") is None
    assert plausible_founded_year("9999") is None
    assert plausible_founded_year("n/a") is None

    row = _row()
    updates = merge_extracted(row, {"founded_year": "1492"}, source="test")
    assert "founded_year" not in updates


def test_identity_text_matches_domain_and_name():
    from vendor_intel.enrichment.gap_fill.gaps import identity_text_matches

    assert identity_text_matches(
        "Acme Distribution", "acme-dist.com", "See https://acme-dist.com/about founded 2001"
    )
    assert identity_text_matches(
        "Acme Distribution", "", "Acme Distribution is a wholesaler in Texas"
    )
    assert not identity_text_matches(
        "Acme Distribution",
        "acme-dist.com",
        "Apex Manufacturing builds turbines in Germany since 1975",
    )


def test_consensus_buffers_then_promotes(monkeypatch):
    from vendor_intel.enrichment.gap_fill import gaps

    monkeypatch.setenv("GAP_FILL_CONSENSUS", "true")
    row = {
        "company": "Acme Distribution",
        "website": "https://acme-dist.com",
        "_crawl": {},
    }
    # Low-trust, no evidence → buffer only
    u1 = gaps.merge_extracted(row, {"founded_year": "2010"}, source="wikipedia")
    assert "founded_year" not in u1
    assert gaps.has_founded(row) is False

    # Second agreeing source → commit
    u2 = gaps.merge_extracted(row, {"founded_year": "2010"}, source="owler")
    assert "founded_year" in u2
    assert row["_crawl"]["founded_year"] == 2010


def test_merge_rejects_wrong_company_evidence():
    from vendor_intel.enrichment.gap_fill.gaps import has_founded, merge_extracted

    row = {
        "company": "Acme Distribution",
        "website": "https://acme-dist.com",
        "_crawl": {},
    }
    updates = merge_extracted(
        row,
        {"founded_year": "1975"},
        source="wikipedia",
        evidence_text="Apex Manufacturing builds turbines in Germany since 1975",
    )
    assert "founded_year" not in updates
    assert not has_founded(row)


def test_merge_extracted_skips_placeholders():
    from vendor_intel.enrichment.gap_fill.gaps import has_employees, merge_extracted

    row = _row()
    updates = merge_extracted(row, {"employee_count": "Not publicly disclosed"}, source="test")
    assert updates == []
    assert not has_employees(row)


def test_parse_sources():
    from vendor_intel.enrichment.gap_fill.sources import DEFAULT_SOURCES, parse_sources

    assert parse_sources(None) == DEFAULT_SOURCES
    assert parse_sources("all") == DEFAULT_SOURCES
    assert parse_sources("linkedin,wikipedia") == ("linkedin", "wikipedia")
    assert parse_sources("bogus,linkedin") == ("linkedin",)


def test_gap_fill_status_empty_batch(tmp_path, monkeypatch):
    from vendor_intel.enrichment.gap_fill.batch import gap_fill_status

    monkeypatch.setattr("vendor_intel.enrichment.gap_fill.batch._done_markets", lambda _bid: [])
    status = gap_fill_status("empty-batch")
    assert status["done_markets"] == 0
    assert status["gap_fill_pending"] == 0


def test_process_market_async_with_gap_fill(tmp_path, monkeypatch):
    from vendor_intel.batch import db, registry_loader, worker
    from vendor_intel.config import Settings

    p = tmp_path / "markets.txt"
    p.write_text("Widget Market | Germany\n", encoding="utf-8")
    registry_loader.populate_registry(p, "gap-test-batch")
    market = db.claim_next_market("gap-test-batch", "test-worker")
    assert market is not None

    gap_calls = []

    async def fake(query, country, settings, output_dir, **kw):
        output_dir.mkdir(parents=True, exist_ok=True)
        out = output_dir / "widget_market_germany.json"
        out.write_text('{"relevant_companies": [{"company": "Acme"}]}', encoding="utf-8")
        return {
            "relevant_companies": [{"company": "Acme", "domain": "acme.com"}],
            "_json_path": str(out),
        }

    async def fake_gap_fill(*args, **kw):
        gap_calls.append(kw)
        return {"updated": 1, "still_missing_gaps": 0}

    import vendor_intel.pipeline.run_and_export as rae

    monkeypatch.setattr(rae, "execute_and_export_async", fake)
    monkeypatch.setattr(
        "vendor_intel.enrichment.gap_fill.orchestrator.gap_fill_market",
        fake_gap_fill,
    )

    asyncio.run(
        worker.process_market_async(
            market,
            Settings.load(),
            fmcg_xlsx=False,
            gap_fill=True,
            gap_fill_concurrent=2,
            gap_fill_sources="linkedin,wikipedia",
        )
    )

    assert len(gap_calls) == 1
    assert gap_calls[0]["concurrent"] == 2
    assert gap_calls[0]["sources"] == "linkedin,wikipedia"
    slug = market["market_slug"]
    row = db.get_market_by_slug(slug)
    assert row["status"] == "done"


def test_main_gap_fill_flag_reaches_orchestrator(tmp_path, monkeypatch, run_batch_mod):
    p = _write_markets_csv(tmp_path)
    calls = []
    monkeypatch.setattr(
        "vendor_intel.batch.orchestrator.run_batch",
        lambda batch_id, settings, **kw: calls.append(kw) or {"done": 2},
    )
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "run_batch.py",
            "--input",
            str(p),
            "--batch-id",
            "b1",
            "--yes",
            "--gap-fill",
            "--gap-fill-concurrent",
            "3",
            "--gap-fill-sources",
            "linkedin,owler",
        ],
    )
    run_batch_mod.main()
    assert calls[0]["gap_fill"] is True
    assert calls[0]["gap_fill_concurrent"] == 3
    assert calls[0]["gap_fill_sources"] == "linkedin,owler"
