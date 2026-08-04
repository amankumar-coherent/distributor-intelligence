"""Unit tests for free Wikidata firmographic lookup (no API key)."""
from __future__ import annotations

import pytest

from vendor_intel.placeholders import wikidata as wd


def test_domain_match_score_prefers_exact_host():
    assert wd._domain_match_score("https://www.microsoft.com/", "microsoft.com") == 100
    assert wd._domain_match_score("https://store.apple.com/us", "apple.com") == 40
    assert wd._domain_match_score("https://evil.com/?q=microsoft.com", "microsoft.com") == 0


def test_parse_bindings_ranks_top_level_over_subsidiary():
    bindings = [
        {
            "item": {"value": "https://www.wikidata.org/entity/Q999"},
            "itemLabel": {"value": "Microsoft Japan"},
            "founded": {"value": "1986-01-01"},
            "hqLabel": {"value": "Tokyo"},
            "website": {"value": "https://www.microsoft.com/ja-jp"},
            "parentLabel": {"value": "Microsoft"},
        },
        {
            "item": {"value": "https://www.wikidata.org/entity/Q2283"},
            "itemLabel": {"value": "Microsoft"},
            "founded": {"value": "1975-04-04"},
            "hqLabel": {"value": "Redmond"},
            "website": {"value": "https://www.microsoft.com/"},
            "employees": {"value": "228000"},
            "industryLabel": {"value": "software development"},
        },
    ]
    parsed = wd._parse_company_bindings(
        bindings,
        preferred_domain="microsoft.com",
        preferred_name="Microsoft",
    )
    assert parsed is not None
    assert parsed["wikidata_qid"] == "Q2283"
    assert parsed["founded_year"] == "1975"
    assert parsed["headquarters"] == "Redmond"


def test_fill_from_wikidata_merges_into_crawl(monkeypatch):
    import asyncio

    from vendor_intel.enrichment.gap_fill.sources import fill_from_wikidata

    async def fake_lookup(name: str, *, domain: str = ""):
        return {
            "wikidata_qid": "Q312",
            "wikidata_label": "Apple",
            "founded_year": "1976",
            "headquarters": "Cupertino",
            "website": "https://www.apple.com/",
            "employee_count": "164000",
            "industry": "electronics",
            "parent_org": "",
            "source_url": "https://www.wikidata.org/wiki/Q312",
        }

    monkeypatch.setattr(
        "vendor_intel.placeholders.wikidata.lookup_company",
        fake_lookup,
    )
    monkeypatch.setattr(
        "vendor_intel.placeholders.wikidata.is_enabled",
        lambda: True,
    )

    row = {"company": "Apple Inc.", "domain": "apple.com", "_crawl": {}}
    applied = asyncio.run(fill_from_wikidata(row))
    assert applied == ["wikidata"]
    crawl = row["_crawl"]
    assert crawl["founded_year"] == 1976
    assert crawl["headquarters"] == "Cupertino"
    assert crawl["wikidata_qid"] == "Q312"
    assert crawl.get("firmographic_source_wikidata") is True
