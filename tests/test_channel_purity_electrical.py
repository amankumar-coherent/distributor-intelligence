"""Tests for electrical channel-purity hard rules."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.intelligence.channel_purity import classify_channel_purity


def test_electrical_oem_as_distributor_dropped():
    for name in (
        "Legrand South Africa (as distributor)",
        "Hyundai Electric & Energy Systems Co., Ltd. (as distributor)",
        "LS Electric Co., Ltd. (as distributor)",
        "Hager South Africa (as distributor)",
        "Hubbell Canada",
        "Phoenix Contact South Africa",
        "Legrand India Pvt. Ltd.",
        "Havells (as separate entity)",
    ):
        v = classify_channel_purity(name, family="electrical")
        assert v.keep is False, name
        assert v.label.startswith("drop"), name


def test_fake_templated_wholesale_dropped():
    v = classify_channel_purity(
        "Electrical Wholesale (Cameroon)",
        family="electrical",
        website="https://www.electricalwholesale.cm",
    )
    assert v.keep is False
    assert v.reason == "fake_templated_channel"


def test_real_distributor_kept():
    for name in ("Rexel", "Sonepar", "Graybar Electric", "Consolidated Electrical Distributors (CED)"):
        v = classify_channel_purity(
            name,
            family="electrical",
            specialty="Electrical wholesale distribution",
            distribution_type="Distributor",
        )
        assert v.keep is True, name
