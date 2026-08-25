"""Tests for canonical founded-year overrides (Excel Founded accuracy)."""
from __future__ import annotations

from vendor_intel.enrichment.known_founded import (
    apply_known_founded_to_row,
    looks_cross_contaminated,
    resolve_known_founded,
    sanitize_founded_value,
)
from vendor_intel.enrichment.gap_fill.gaps import merge_extracted
from vendor_intel.export.landscape_slim import finalize_slim_row


def test_resolve_major_brands():
    assert resolve_known_founded("Rexel India") == 1967
    assert resolve_known_founded("Sonepar USA") == 1969
    assert resolve_known_founded("WESCO International") == 1922
    assert resolve_known_founded("Graybar Electric") == 1869
    assert resolve_known_founded("Graybar Canada") == 1920
    assert resolve_known_founded("Edmundson Electrical") == 1801
    assert resolve_known_founded("Blackwoods") == 1878
    assert resolve_known_founded("Al-Futtaim") == 1939
    assert resolve_known_founded("Guillevin (Sonepar Canada)") == 1906
    assert resolve_known_founded("Westburne (Rexel Canada)") == 1926
    assert resolve_known_founded("Nedco (Sonepar Canada)") == 1911
    assert resolve_known_founded("Random Local Wholesaler XYZ") is None


def test_sanitize_overrides_incorporation_and_contamination():
    # Incorporation mistaken for founding
    assert sanitize_founded_value("Edmundson Electrical", "1991") == "1801"
    # Cross-brand contamination (Rexel got WESCO year)
    assert sanitize_founded_value("Rexel India", "1922") == "1967"
    # Off-by-small
    assert sanitize_founded_value("Blackwoods", "1876") == "1878"
    # Unknown stays plausible
    assert sanitize_founded_value("Acme Electrical Supply Co", "1984") == "1984"


def test_looks_cross_contaminated():
    assert looks_cross_contaminated("Rexel India", 1922) is True
    assert looks_cross_contaminated("WESCO International", 1922) is False


def test_merge_extracted_forces_known_year():
    row = {
        "company": "Rexel India",
        "_crawl": {"founded_year": 1922},
    }
    updates = merge_extracted(
        row,
        {"founded_year": 1922},
        source="test",
        force_commit=True,
    )
    assert "founded_year" in updates
    assert row["_crawl"]["founded_year"] == 1967


def test_finalize_slim_row_writes_canonical_founded():
    row = {
        "Company": "Graybar Electric",
        "Founded": "1926",
        "Headquarters": "St. Louis, United States",
    }
    out = finalize_slim_row(row)
    assert out["Founded"] == "1869"


def test_apply_known_founded_to_row():
    out = apply_known_founded_to_row(
        {"Company": "Sonepar", "Founded": "Not publicly disclosed"}
    )
    assert out["Founded"] == "1969"
