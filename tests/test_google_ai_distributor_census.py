"""Tests for Google AI distributor census prompts + name extraction."""
from __future__ import annotations

from vendor_intel.discovery.google_ai_distributor_census import (
    build_distributor_census_queries,
    extract_company_names_heuristic,
    names_to_discovery_hits,
)


def test_queries_include_market_and_distributor_language():
    qs = build_distributor_census_queries(
        "Global Avocado Oil Market", target=100, max_queries=10
    )
    assert qs
    joined = " | ".join(qs).lower()
    assert "avocado oil" in joined
    assert "distributor" in joined
    assert "manufacturer" in joined  # exclusion instruction
    assert "firewall" not in joined


def test_queries_change_with_market_name():
    a = build_distributor_census_queries("Global Firewall Market", max_queries=5)
    b = build_distributor_census_queries("Global Smartwatch Market", max_queries=5)
    assert "firewall" in a[0].lower()
    assert "smartwatch" in b[0].lower()
    assert a[0] != b[0]


def test_heuristic_extracts_bullet_companies():
    md = """
AI Overview
Here are distributors:
- **Brenntag** - bulk ingredient distributor
- **Gustav Heess** - plant oil wholesaler
1. Olivado
2. Avocado Real
Show more
"""
    names = extract_company_names_heuristic(md)
    joined = " | ".join(names)
    assert "Brenntag" in joined
    assert "Gustav Heess" in joined
    assert "Olivado" in joined


def test_names_to_hits_unique_domains():
    hits = names_to_discovery_hits(
        ["Brenntag", "Gustav Heess"], market="avocado oil", prompt_id="T"
    )
    assert len(hits) == 2
    assert hits[0].backend == "google_ai_census"
    assert hits[0].source_domain != hits[1].source_domain
    assert "distributor" in hits[0].snippet.lower()
