"""Core Categories, Specialty Focus, Key Brands and sales channel.

All four were asked on EVERY company and none had any extraction code, so
they sat at exactly 0% across 1,254 rows regardless of how good Google's
answer was — roughly 5,000 wasted queries.

The strings below are the real live answers for
Cantek Teknik Hizmetler Ticaret Ltd Sti.
"""
from __future__ import annotations

from vendor_intel.discovery.chatgpt_company_screen import _extract_fields_from_prose as ex

LIVE_CATEGORIES = (
    "Cantek distributes medical equipment, consumables, and water treatment systems, "
    "primarily for hospitals and dialysis centers. Its medical product categories "
    "include hemodialysis devices (from brands like Kawasumi and Toray)."
)
LIVE_SPECIALTY = (
    "Cantek Teknik Hizmetler Ltd is a Turkish company headquartered in Antalya, "
    "specializing in the healthcare sector, specifically hemodialysis equipment, "
    "dialysis water purification systems, and medical technical services."
)
LIVE_BRANDS = (
    "Cantek is the authorized technical service and distributor for Kawasumi and Toray "
    "brand hemodialysis devices in Turkey. It distributes Kawasumi hemodialysis devices."
)
LIVE_CHANNEL = (
    "Cantek sells offline through traditional B2B distribution channels, direct "
    "procurement contracts, and technical service networks. It does not operate a "
    "direct online e-commerce shop."
)


class TestCoreCategories:
    def test_live_answer(self):
        got = ex(LIVE_CATEGORIES, ["Core Categories"])["Core Categories"]
        assert "medical equipment" in got

    def test_trailing_audience_clause_is_dropped(self):
        """'primarily for hospitals and dialysis centers' describes WHO they
        sell to, not a product category."""
        got = ex(LIVE_CATEGORIES, ["Core Categories"])["Core Categories"]
        assert "hospitals" not in got

    def test_not_asked_means_not_extracted(self):
        assert "Core Categories" not in ex(LIVE_CATEGORIES, ["Website"])


class TestSpecialtyFocus:
    def test_live_answer(self):
        got = ex(LIVE_SPECIALTY, ["Specialty Focus"])["Specialty Focus"]
        assert "hemodialysis" in got

    def test_value_stays_cell_sized(self):
        got = ex(LIVE_SPECIALTY, ["Specialty Focus"])["Specialty Focus"]
        assert len(got) <= 120


class TestKeyBrands:
    def test_live_answer_splits_the_brand_names(self):
        assert ex(LIVE_BRANDS, ["Key Brands Represented"])["Key Brands Represented"] == (
            "Kawasumi, Toray"
        )

    def test_trailing_product_nouns_are_stripped(self):
        """'Kawasumi and Toray brand hemodialysis devices' must not yield a
        brand called 'Toray brand hemodialysis devices'."""
        got = ex(LIVE_BRANDS, ["Key Brands Represented"])["Key Brands Represented"]
        assert "hemodialysis" not in got.lower()

    def test_brands_like_phrasing(self):
        got = ex("It carries brands like Philips and Siemens.", ["Key Brands Represented"])
        assert got["Key Brands Represented"] == "Philips, Siemens"

    def test_no_brands_mentioned(self):
        assert ex("Acme is a distributor.", ["Key Brands Represented"]) == {}


class TestSalesChannel:
    def test_offline_despite_the_word_online_appearing(self):
        """'does NOT operate a direct online e-commerce shop' is the
        strongest possible offline signal — reading it as Online inverts
        the answer."""
        assert ex(LIVE_CHANNEL, ["Online / Offline / Both"])["Online / Offline / Both"] == "Offline"

    def test_online(self):
        got = ex("Acme sells online through its e-commerce store.", ["Online / Offline / Both"])
        assert got["Online / Offline / Both"] == "Online"

    def test_both_explicit(self):
        got = ex("Acme sells both online and offline.", ["Online / Offline / Both"])
        assert got["Online / Offline / Both"] == "Both"

    def test_no_channel_signal(self):
        assert ex("Acme is a company.", ["Online / Offline / Both"]) == {}
