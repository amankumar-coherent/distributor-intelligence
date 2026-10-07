"""The Kompass -> gap-fill bridge.

The risk this file guards is silent DATA LOSS: gap-fill is a research
engine, so if a good Kompass value (a phone number straight off the
company's own Kompass listing) is presented as a gap, a weaker web guess
can replace it. Seeding and the write-back guard must both hold.
"""
from __future__ import annotations

import pytest

from vendor_intel.discovery.kompass_distributors import VerifiedDistributor
from vendor_intel.pipeline.kompass_gap_adapter import apply_gap_row, to_gap_row

NOT_DISCLOSED = "Not publicly disclosed"


def make(**kw) -> VerifiedDistributor:
    base = dict(
        company_name="Inycom S.A.",
        market="Wearable Medical Devices",
        kompass_company_id="ES1165858",
        kompass_profile_url="https://lb.kompass.com/c/inycom/",
        kompass_product_id="23200",
        kompass_company_type="Distributor",
        is_actual_market_distributor=True,
        verification_confidence=60,
    )
    base.update(kw)
    return VerifiedDistributor(**base)


class TestSeeding:
    def test_known_values_are_seeded_so_gap_fill_skips_them(self):
        row = to_gap_row(make(website="inycom.es", phone="+34 976 013 300"))
        assert row["_crawl"]["website"] == "inycom.es"
        assert row["_crawl"]["phone"] == "+34 976 013 300"

    def test_placeholder_is_not_seeded_and_stays_a_gap(self):
        """'Not publicly disclosed' is a display placeholder. Seeding it
        would tell gap-fill the field is done and suppress the research."""
        row = to_gap_row(make(employees=NOT_DISCLOSED, headquarters=""))
        assert "employee_count" not in row["_crawl"]
        assert "headquarters" not in row["_crawl"]

    def test_market_name_as_core_categories_is_treated_as_a_gap(self):
        row = to_gap_row(make(products="Wearable Medical Devices"))
        assert "core_categories" not in row["_crawl"]

    def test_real_core_categories_survives(self):
        row = to_gap_row(make(products="Lab analysers, patient monitors"))
        assert row["_crawl"]["core_categories"] == "Lab analysers, patient monitors"

    def test_extra_columns_are_seeded_under_their_gap_key(self):
        row = to_gap_row(make(extra_columns={"Role": "Purchasing Manager"}))
        assert row["_crawl"]["contact_role"] == "Purchasing Manager"

    def test_company_name_is_carried_for_search(self):
        row = to_gap_row(make())
        assert row["company"] == "Inycom S.A."
        assert row["brand"] == "Inycom S.A."


class TestWriteBack:
    def test_learned_values_land_on_the_dataclass(self):
        d = make()
        changed = apply_gap_row(d, {"_crawl": {"linkedin": "linkedin.com/company/inycom",
                                               "email": "info@inycom.es"}})
        assert d.email == "info@inycom.es"
        assert d.extra_columns["LinkedIn"] == "linkedin.com/company/inycom"
        assert set(changed) == {"email", "LinkedIn"}

    def test_never_overwrites_an_existing_kompass_value(self):
        """The core safety property. Kompass gave us the phone number off
        the company's own listing; a web-scraped guess must not win."""
        d = make(phone="+34 976 013 300")
        apply_gap_row(d, {"_crawl": {"phone": "+34 900 000 000"}})
        assert d.phone == "+34 976 013 300"

    def test_fills_over_a_placeholder(self):
        d = make(employees=NOT_DISCLOSED)
        apply_gap_row(d, {"_crawl": {"employee_count": "700"}})
        assert d.employees == "700"

    def test_ignores_a_placeholder_coming_back_from_gap_fill(self):
        d = make()
        apply_gap_row(d, {"_crawl": {"ownership": NOT_DISCLOSED}})
        assert not (d.extra_columns or {}).get("Ownership")

    def test_missing_crawl_is_harmless(self):
        d = make(website="inycom.es")
        assert apply_gap_row(d, {}) == []
        assert d.website == "inycom.es"

    def test_round_trip_changes_nothing_when_nothing_was_learned(self):
        """Seeding then writing back the same row must be a no-op —
        otherwise a refill would churn the checkpoint on every pass."""
        d = make(website="inycom.es", phone="+34 976 013 300", founded_year="1982")
        assert apply_gap_row(d, to_gap_row(d)) == []


class TestGeographyColumns:
    """The two columns the user specifically wants filled with NAMED
    countries rather than 'Global'."""

    def test_regions_and_offices_map_to_the_two_geography_columns(self):
        d = make()
        apply_gap_row(d, {"_crawl": {"regions_served": "Spain, Ecuador",
                                     "offices": "Zaragoza, Madrid"}})
        assert d.extra_columns["Continent / Geography"] == "Spain, Ecuador"
        assert d.extra_columns["Operational Presence"] == "Zaragoza, Madrid"

    def test_existing_named_countries_are_seeded_not_re_researched(self):
        row = to_gap_row(make(extra_columns={"Continent / Geography": "Spain, Ecuador"}))
        assert row["_crawl"]["regions_served"] == "Spain, Ecuador"
