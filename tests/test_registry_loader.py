"""Market-list file -> SQLite registry loader (src/vendor_intel/batch/registry_loader.py)."""
from __future__ import annotations

import pytest

from vendor_intel.batch import db, registry_loader


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    yield
    db.reset_for_tests()


def test_parse_txt_file(tmp_path):
    p = tmp_path / "markets.txt"
    p.write_text(
        "# comment line, skip me\n"
        "\n"
        "Avocado Oil Market | Global\n"
        "Rupture Disc Market | Brazil | Manufacturers; Distributors\n",
        encoding="utf-8",
    )
    rows = registry_loader.parse_input_file(p)
    assert len(rows) == 2
    assert rows[0] == {
        "query": "Avocado Oil Market",
        "country": "Global",
        "market_type": "general",
        "sections": [],
        "seeds_file": None,
        "priority": 100,
    }
    assert rows[1]["sections"] == ["Manufacturers", "Distributors"]


def test_parse_csv_file(tmp_path):
    p = tmp_path / "markets.csv"
    p.write_text(
        "query,country,market_type,sections,seeds_file,priority\n"
        "Gel Ice Pack Market,Global,fmcg,Manufacturers;Distributors,,10\n"
        "High Energy Laser Market,USA,security_tech,,queries/seeds/laser.txt,\n",
        encoding="utf-8",
    )
    rows = registry_loader.parse_input_file(p)
    assert len(rows) == 2
    assert rows[0] == {
        "query": "Gel Ice Pack Market",
        "country": "Global",
        "market_type": "fmcg",
        "sections": ["Manufacturers", "Distributors"],
        "seeds_file": None,
        "priority": 10,
    }
    assert rows[1]["market_type"] == "security_tech"
    assert rows[1]["seeds_file"] == "queries/seeds/laser.txt"
    assert rows[1]["priority"] == 100  # blank -> default


def test_csv_missing_query_is_skipped(tmp_path):
    p = tmp_path / "markets.csv"
    p.write_text("query,country\n,Global\nReal Market,Global\n", encoding="utf-8")
    rows = registry_loader.parse_input_file(p)
    assert len(rows) == 1
    assert rows[0]["query"] == "Real Market"


def test_populate_registry_inserts_into_db(tmp_path):
    p = tmp_path / "markets.txt"
    p.write_text("Avocado Oil Market | Global\nRupture Disc Market | Brazil\n", encoding="utf-8")
    summary = registry_loader.populate_registry(p, "batch-1")
    assert summary == {"total": 2, "inserted": 2, "updated": 0, "skipped_duplicate_in_file": 0}

    markets = db.list_markets("batch-1")
    assert len(markets) == 2
    assert {m["status"] for m in markets} == {"pending"}


def test_populate_registry_is_idempotent_and_resumable(tmp_path):
    p = tmp_path / "markets.txt"
    p.write_text("Avocado Oil Market | Global\n", encoding="utf-8")
    registry_loader.populate_registry(p, "batch-1")

    market = db.claim_next_market("batch-1", "worker-1")
    db.mark_market_done(market["id"], companies_exported=42, csv_path="x.csv")

    # Re-loading the same file (e.g. resuming a killed batch process) must not
    # reset progress on the market that already finished.
    summary = registry_loader.populate_registry(p, "batch-1")
    assert summary["inserted"] == 0
    assert summary["updated"] == 1

    row = db.get_market_by_slug("avocado_oil_market_global")
    assert row["status"] == "done"
    assert row["companies_exported"] == 42


def test_populate_registry_dedupes_within_file(tmp_path):
    p = tmp_path / "markets.txt"
    p.write_text(
        "Avocado Oil Market | Global\n"
        "Avocado Oil Market | Global | Extra Section\n",  # same slug, appears twice
        encoding="utf-8",
    )
    summary = registry_loader.populate_registry(p, "batch-1")
    assert summary["total"] == 2
    assert summary["inserted"] == 1
    assert summary["skipped_duplicate_in_file"] == 1
    assert len(db.list_markets("batch-1")) == 1


def test_populate_registry_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        registry_loader.populate_registry(tmp_path / "nope.txt", "batch-1")


def test_populate_registry_empty_file_raises(tmp_path):
    p = tmp_path / "empty.txt"
    p.write_text("# only comments\n\n", encoding="utf-8")
    with pytest.raises(ValueError):
        registry_loader.populate_registry(p, "batch-1")
