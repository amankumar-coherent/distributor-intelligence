"""Mid-enrich checkpoint: skip DNS list + already-enriched companies on resume."""
from __future__ import annotations

import asyncio
import json

from vendor_intel.pipeline.checkpoint_resume import (
    company_already_enriched,
    load_phase3_checkpoint,
    phase3_enrich_path,
    save_phase3_checkpoint,
)


def test_save_load_phase3_checkpoint_merge(tmp_path, monkeypatch):
    monkeypatch.setattr(
        "vendor_intel.pipeline.checkpoint_resume._project_root",
        lambda: tmp_path,
    )
    path = phase3_enrich_path(query="Test Market", market_slug="test_market_global")
    assert path.parent.name == "phase3"

    save_phase3_checkpoint(
        path,
        query="Test Market",
        market_slug="test_market_global",
        companies=[{"name": "Acme", "domain": "acme.com"}],
        enriched={"Acme": {"ok": 1}, "acme.com": {"ok": 1}},
        merge=False,
    )
    save_phase3_checkpoint(
        path,
        query="Test Market",
        market_slug="test_market_global",
        companies=None,
        enriched={"Beta": {"ok": 2}, "beta.com": {"ok": 2}},
        merge=True,
    )
    data = load_phase3_checkpoint(path)
    assert data is not None
    assert len(data["companies"]) == 1
    assert "Acme" in data["enriched"] and "Beta" in data["enriched"]
    assert data["enriched_count"] == 4


def test_company_already_enriched_by_domain_or_name():
    enriched = {"acme.com": {"pages": []}}
    assert company_already_enriched({"name": "Acme", "domain": "acme.com"}, enriched)
    assert company_already_enriched({"name": "Acme", "domain": "www.acme.com"}, enriched)
    assert not company_already_enriched({"name": "Other", "domain": "other.com"}, enriched)


def test_enrich_companies_skips_checkpointed(monkeypatch, tmp_path):
    from vendor_intel.enrichment import smart_enrichment as se

    calls: list[str] = []

    async def fake_crawl(name, domain, *, country="", use_ssc=False):
        calls.append(domain)
        return name, {"domain": domain, "data": {"ok": True}}, None

    monkeypatch.setattr(se, "_crawl_one", fake_crawl)

    class _S:
        pipeline_use_ssc = True
        company_cache_enabled = False

        @staticmethod
        def load():
            return _S()

    monkeypatch.setattr("vendor_intel.config.Settings.load", _S.load)

    ck = tmp_path / "phase3_enrich_test.json"
    save_phase3_checkpoint(
        ck,
        query="T",
        companies=[
            {"name": "A", "domain": "a.com"},
            {"name": "B", "domain": "b.com"},
        ],
        enriched={"A": {"cached": True}, "a.com": {"cached": True}},
        merge=False,
    )

    out = asyncio.run(
        se.enrich_companies(
            [
                {"name": "A", "domain": "a.com"},
                {"name": "B", "domain": "b.com"},
            ],
            limit=10,
            max_concurrent=2,
            checkpoint_path=ck,
            resume=True,
            query="T",
        )
    )
    assert calls == ["b.com"]
    assert "A" in out and out["A"].get("cached") is True
    assert "B" in out
    data = json.loads(ck.read_text(encoding="utf-8"))
    assert "B" in data["enriched"] or "b.com" in data["enriched"]
