"""Tests for Owler gap-fill merge behavior."""
from __future__ import annotations

from vendor_intel.enrichment.owler_enrich import apply_owler_to_row


def test_apply_owler_does_not_overwrite_existing_headquarters():
    row = {
        "_crawl": {"headquarters": "Boston, MA, USA", "employee_count": "500"},
        "_crawl_raw": {"data": {"location": {"headquarters": "Boston, MA, USA"}}},
        "enrichment_source": "smart_crawl",
    }
    updates = apply_owler_to_row(
        row,
        {
            "headquarters": "New York, NY, USA",
            "employee_count": "1000",
            "revenue": "$50M",
        },
    )
    assert row["_crawl"]["headquarters"] == "Boston, MA, USA"
    assert row["_crawl"]["employee_count"] == "500"
    assert row["_crawl"]["revenue"] == "$50M"
    assert "HQ=" not in " ".join(updates)
    assert "Turnover=$50M" in " ".join(updates)


def test_apply_owler_fills_empty_fields():
    row = {"_crawl": {}, "_crawl_raw": {"data": {}}}
    updates = apply_owler_to_row(
        row,
        {
            "headquarters": "Paris, France",
            "founded_year": 1999,
            "employee_count": "200",
        },
    )
    assert row["_crawl"]["headquarters"] == "Paris, France"
    assert row["_crawl"]["founded_year"] == 1999
    assert row["_crawl"]["employee_count"] == "200"
    assert len(updates) == 3
    assert "owler" in row["enrichment_source"]
