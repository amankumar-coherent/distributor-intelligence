"""Mid-classify checkpoint: skip already-classified companies on resume."""
from __future__ import annotations

from vendor_intel.pipeline.checkpoint_resume import (
    company_already_classified,
    load_phase4_checkpoint,
    lookup_classified,
    phase4_classify_path,
    save_phase4_checkpoint,
)


def test_save_load_phase4_checkpoint_merge(tmp_path, monkeypatch):
    monkeypatch.setattr(
        "vendor_intel.pipeline.checkpoint_resume._project_root",
        lambda: tmp_path,
    )
    path = phase4_classify_path(query="Test Market", market_slug="test_market_global")
    assert path.parent.name == "phase4"

    save_phase4_checkpoint(
        path,
        query="Test Market",
        market_slug="test_market_global",
        classified={
            "Acme": {"company": "Acme", "domain": "acme.com", "confidence": 0.8},
            "acme.com": {"company": "Acme", "domain": "acme.com", "confidence": 0.8},
        },
        merge=False,
    )
    save_phase4_checkpoint(
        path,
        query="Test Market",
        market_slug="test_market_global",
        classified={
            "Beta": {"company": "Beta", "domain": "beta.com", "confidence": 0.7},
            "beta.com": {"company": "Beta", "domain": "beta.com", "confidence": 0.7},
        },
        merge=True,
    )
    data = load_phase4_checkpoint(path)
    assert data is not None
    assert "Acme" in data["classified"] and "Beta" in data["classified"]
    assert data["classified_count"] >= 2


def test_company_already_classified_and_lookup():
    classified = {
        "acme.com": {"company": "Acme", "domain": "acme.com", "is_relevant": True},
    }
    assert company_already_classified({"name": "Acme", "domain": "acme.com"}, classified)
    assert company_already_classified({"name": "Acme", "domain": "www.acme.com"}, classified)
    assert not company_already_classified({"name": "Other", "domain": "other.com"}, classified)
    v = lookup_classified({"name": "Acme", "domain": "www.acme.com"}, classified)
    assert v is not None and v["company"] == "Acme"
