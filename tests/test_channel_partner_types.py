"""Unit tests for market-aware channel-partner taxonomy."""
from __future__ import annotations

from vendor_intel.funnel.channel_partner_types import (
    apply_confidence_tier_to_rows,
    channel_partner_mode_enabled,
    confidence_tier,
    has_channel_evidence,
    infer_channel_partner_types,
    is_accepted_channel_type,
    is_production_exportable,
    is_pure_manufacturer_reject,
    map_to_channel_type,
)
from vendor_intel.pipeline.quality_export import normalize_export_role, score_export_row


def test_infer_cyber_types():
    types = infer_channel_partner_types("Global Cybersecurity Market", "MSSP")
    assert "MSSP" in types
    assert "VAR" in types
    assert "Manufacturer" not in types


def test_infer_foam_types():
    types = infer_channel_partner_types("Technical Foam Market", "global")
    assert "Converter" in types or "Fabricator" in types
    assert "Distributor" in types


def test_infer_physical_security():
    types = infer_channel_partner_types("Physical Security Market", "")
    assert "Security Integrator" in types or "Installer" in types


def test_map_channel_types():
    assert map_to_channel_type("VAR", "") == "VAR"
    assert map_to_channel_type("Integrator", "") == "System Integrator"
    assert map_to_channel_type("", "mssp") == "MSSP"
    assert map_to_channel_type("", "foam fabricator") == "Fabricator"


def test_reject_pure_manufacturer():
    assert is_pure_manufacturer_reject("Manufacturer", text="we manufacture foam panels")
    assert not is_pure_manufacturer_reject(
        "Manufacturer",
        signals={"is_distributor": True},
        text="authorized distributor of foam brands",
    )
    assert is_accepted_channel_type("Distributor")
    assert not is_accepted_channel_type("Manufacturer")


def test_channel_mode_flag():
    assert channel_partner_mode_enabled({}) is True
    assert channel_partner_mode_enabled({"discovery_mode": "landscape"}) is False
    assert channel_partner_mode_enabled({"discovery_mode": "channel_partners"}) is True


def test_normalize_export_role_channel():
    assert normalize_export_role("VAR", "") == "VAR"
    assert normalize_export_role("Integrator", "") == "System Integrator"
    assert normalize_export_role("", "authorized distributor") == "Distributor"


def test_export_gate_rejects_manufacturer():
    q, reason = score_export_row(
        {
            "company": "Acme Foam Makers",
            "domain": "acmefoam.example",
            "is_relevant": True,
            "role": "Manufacturer",
            "confidence": 0.9,
            "signals": {"has_products": True, "is_manufacturer": True},
        },
        {"pages": [{"text": "We manufacture polyurethane foam products for OEMs worldwide. " * 5}]},
        {"industry": "Technical Foam Market", "discovery_mode": "channel_partners", "country": "global"},
    )
    assert q == 0.0
    assert reason == "not_channel_partner"


def test_export_gate_accepts_distributor():
    q, reason = score_export_row(
        {
            "company": "Foam Channel Partners",
            "domain": "foamchannel.example",
            "is_relevant": True,
            "role": "Distributor",
            "confidence": 0.85,
            "company_function": "distributor",
            "signals": {"has_products": True, "is_distributor": True, "industry_match": True},
            "role_description": "Distributor of technical foam for industrial packaging",
            "summary": "Authorized distributor of foam brands across North America.",
        },
        {
            "pages": [
                {
                    "text": (
                        "Authorized distributor of technical foam products and channel partner "
                        "for industrial packaging converters worldwide. " * 8
                    )
                }
            ]
        },
        {"industry": "Technical Foam Market", "discovery_mode": "channel_partners", "country": "global"},
    )
    assert reason == ""
    assert q > 0.0


def test_confidence_tiers():
    assert confidence_tier(0.95) == "auto"
    assert confidence_tier(0.90) == "auto"
    assert confidence_tier(0.89) == "high"
    assert confidence_tier(0.70) == "high"
    assert confidence_tier(0.69) == "review"
    assert confidence_tier(0.50) == "review"
    assert confidence_tier(0.49) == "reject"
    assert confidence_tier(None) == "reject"
    assert confidence_tier(95) == "auto"  # 0–100 scale tolerated
    assert confidence_tier(55) == "review"


def test_apply_confidence_tiers_mutates_rows():
    rows = [
        {"company": "A", "is_relevant": True, "confidence": 0.95},
        {"company": "B", "is_relevant": True, "confidence": 0.75},
        {"company": "C", "is_relevant": True, "confidence": 0.55},
        {"company": "D", "is_relevant": True, "confidence": 0.30},
    ]
    counts = apply_confidence_tier_to_rows(rows)
    assert counts == {"auto": 1, "high": 1, "review": 1, "reject": 1}
    assert rows[0]["review_status"] == "Auto Export" and rows[0]["is_relevant"] is True
    assert rows[1]["review_status"] == "High Confidence" and rows[1]["is_relevant"] is True
    assert rows[2]["_review_only"] is True and rows[2]["is_relevant"] is True
    assert rows[3]["is_relevant"] is False and rows[3]["_channel_gate"] == "low_confidence"


def test_production_gate_exports_all_non_reject_tiers():
    assert is_production_exportable({"confidence": 0.95})  # auto
    assert is_production_exportable({"confidence": 0.75})  # high
    assert is_production_exportable({"confidence": 0.55})  # review
    assert not is_production_exportable({"confidence": 0.40})  # reject
    assert has_channel_evidence({"_evidence_verified": True})
    assert not has_channel_evidence({"evidence": "Authorized Distributor", "evidence_url": "u"})


def test_analyst_seeds_bypass_gates_but_enumerated_players_do_not():
    from vendor_intel.funnel.channel_partner_types import apply_channel_gate_to_rows

    # Only ANALYST seeds (is_analyst_seed) bypass the gate. An auto-enumerated
    # market player (is_seed only) that is a pure manufacturer must be dropped.
    rows = [
        {"company": "Analyst Seed", "is_relevant": True,
         "is_seed": True, "is_analyst_seed": True,
         "role": "Manufacturer", "company_function": "manufacturer", "confidence": 0.2},
        {"company": "Enumerated Maker", "is_relevant": True,
         "is_seed": True, "is_analyst_seed": False,
         "role": "Distributor", "company_function": "manufacturer", "confidence": 0.2},
        {"company": "Normal Maker", "is_relevant": True, "is_seed": False,
         "role": "Distributor", "company_function": "manufacturer", "confidence": 0.2},
    ]
    apply_channel_gate_to_rows(rows)
    assert rows[0]["is_relevant"] is True    # analyst seed kept
    assert rows[1]["is_relevant"] is False   # enumerated manufacturer dropped
    assert rows[2]["is_relevant"] is False   # non-seed manufacturer dropped

    apply_confidence_tier_to_rows(rows)
    assert rows[0]["is_relevant"] is True     # analyst seed kept despite conf 0.2
    assert is_production_exportable(rows[0]) is True
    assert is_production_exportable(rows[1]) is False  # enumerated maker: reject tier


def test_landscape_mode_allows_manufacturer():
    q, reason = score_export_row(
        {
            "company": "Acme Foam Makers",
            "domain": "acmefoam.example",
            "is_relevant": True,
            "role": "Manufacturer",
            "confidence": 0.9,
            "signals": {"has_products": True, "is_manufacturer": True, "industry_match": True},
            "role_description": "Manufacturer of polyurethane foam panels",
            "summary": "Industrial manufacturer of technical foam products.",
        },
        {"pages": [{"text": "We manufacture polyurethane foam products for OEMs worldwide. " * 8}]},
        {"industry": "Technical Foam Market", "discovery_mode": "landscape", "country": "global"},
    )
    # May still reject for other quality reasons, but not not_channel_partner
    assert reason != "not_channel_partner"
rket", "discovery_mode": "landscape", "country": "global"},
    )
    # May still reject for other quality reasons, but not not_channel_partner
    assert reason != "not_channel_partner"
r"
rket", "discovery_mode": "landscape", "country": "global"},
    )
    # May still reject for other quality reasons, but not not_channel_partner
    assert reason != "not_channel_partner"
er"
r"
rket", "discovery_mode": "landscape", "country": "global"},
    )
    # May still reject for other quality reasons, but not not_channel_partner
    assert reason != "not_channel_partner"
