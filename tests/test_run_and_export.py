"""Shared market-run/export glue (src/vendor_intel/pipeline/run_and_export.py)."""
from __future__ import annotations

from vendor_intel.pipeline.run_and_export import (
    auto_discovery_queries,
    parse_market_line,
    parse_sections,
    slugify,
)


def test_parse_market_line_market_only():
    assert parse_market_line("Atomic Clock Market") == ("Atomic Clock Market", "global", [])


def test_parse_market_line_with_country():
    assert parse_market_line("Atomic Clock Market | Germany") == (
        "Atomic Clock Market", "Germany", [],
    )


def test_parse_market_line_with_country_and_sections():
    market, country, sections = parse_market_line(
        "Drone Battery Market | global | Battery Pack Manufacturers; "
        "OEM Battery Manufacturers; BMS Manufacturers"
    )
    assert market == "Drone Battery Market"
    assert country == "global"
    assert sections == [
        "Battery Pack Manufacturers",
        "OEM Battery Manufacturers",
        "BMS Manufacturers",
    ]


def test_parse_market_line_functionality_marker():
    market, country, sections = parse_market_line(
        "Drone Battery Market ; functionality: Battery Pack Manufacturers, "
        "BMS Manufacturers, Contract Manufacturer"
    )
    assert market == "Drone Battery Market"
    assert country == "global"
    assert sections == [
        "Battery Pack Manufacturers",
        "BMS Manufacturers",
        "Contract Manufacturer",
    ]


def test_slugify():
    assert slugify("Atomic Clock Market", "Germany") == "atomic_clock_market_germany"


def test_parse_sections_semicolon_and_comma():
    assert parse_sections("A; B; C") == ["A", "B", "C"]
    assert parse_sections("A, B, C") == ["A", "B", "C"]
    assert parse_sections(None) == []
    assert parse_sections("  ") == []


def test_auto_discovery_queries_avocado_does_not_inherit_firewall():
    qs = auto_discovery_queries("Global Avocado Oil Market", "global")
    joined = " | ".join(qs).lower()
    assert "firewall" not in joined
    assert "mssp" not in joined
    assert "cybersecurity var" not in joined


def test_auto_discovery_queries_firewall_still_matches():
    qs = auto_discovery_queries("Global Firewall Market", "global")
    assert qs, "firewall market should load curated discovery queries"
    joined = " | ".join(qs).lower()
    assert "firewall" in joined


def test_auto_discovery_queries_smartwatch_no_firewall_bleed():
    qs = auto_discovery_queries("Global Smartwatch Market", "global")
    joined = " | ".join(qs).lower()
    assert "firewall" not in joined
    assert "mssp" not in joined
