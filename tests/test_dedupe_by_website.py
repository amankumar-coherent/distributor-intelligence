"""The same company was listed up to eight times under name variants.

Measured on the finished workbooks:

    Dry Eyes : 373 rows -> 322 unique (51 redundant, 13%)
    Solar    : 401 rows -> 325 unique (76 redundant, 19%)

The pipeline already dedupes on a normalised company name, but these get
past it because the names genuinely differ:

    powernsun.com   PowernSun
                    Powernsun (Power n Sun Solar Systems)
                    Power & Sun
                    Power n Sun
                    Power & Sun FZCO
                    Powernsun Group
                    Power N Sun Solar Equipment's Trading LLC

Seven rows, one Dubai company. A shared website is far stronger evidence
of identity than a name, because a company controls one primary domain.

The hard part is that a shared domain is NOT proof on its own. A
multinational runs one domain across genuinely distinct national
distributors:

    krannich-solar.com  Krannich Solar GmbH & Co. KG   Germany
                        Krannich Solar Turkey          Turkey
                        Krannich Solar d.o.o.          Croatia
                        Krannich Solar Energy (Pty)    South Africa
                        Krannich Solar Private Limited India

Those are five real distributors with different geography and contacts.
Merging them would destroy information, so the rule is domain AND
country.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location(
    "dedupe_by_website", _ROOT / "dedupe_by_website.py"
)
dd = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(dd)


def _row(company, website, hq, **extra):
    base = {"Company": company, "Website": website, "Headquarters": hq}
    base.update(extra)
    return base


# --- domain extraction ---------------------------------------------------


@pytest.mark.parametrize(
    "raw",
    ["https://www.powernsun.com", "http://powernsun.com/",
     "www.powernsun.com", "powernsun.com", "https://powernsun.com/products",
     "HTTPS://WWW.POWERNSUN.COM"],
)
def test_domains_normalise_to_the_same_key(raw):
    assert dd.domain_of(raw) == "powernsun.com"


def test_different_domains_stay_different():
    assert dd.domain_of("https://a.example") != dd.domain_of("https://b.example")


def test_a_missing_website_has_no_domain():
    assert dd.domain_of("") == ""
    assert dd.domain_of(None) == ""


# --- country, which decides whether a shared domain merges ---------------


def test_country_comes_from_the_end_of_headquarters():
    assert dd.country_of({"Headquarters": "Dubai, United Arab Emirates"}) == \
        dd.country_of({"Headquarters": "Sharjah, United Arab Emirates"})


def test_country_comparison_ignores_case_and_punctuation():
    assert dd.country_of({"Headquarters": "Weil der Stadt, Germany"}) == \
        dd.country_of({"Headquarters": "Berlin,  germany "})


def test_different_countries_do_not_match():
    assert dd.country_of({"Headquarters": "Istanbul, Turkey"}) != \
        dd.country_of({"Headquarters": "Zagreb, Croatia"})


# --- the merge rule ------------------------------------------------------


def test_name_variants_on_one_domain_and_country_merge():
    """The PowernSun case: seven rows, one company."""
    rows = [
        _row("PowernSun", "https://powernsun.com", "Dubai, United Arab Emirates"),
        _row("Power & Sun", "https://www.powernsun.com", "Dubai, United Arab Emirates"),
        _row("Power n Sun", "powernsun.com", "Dubai, UAE"),
    ]
    out, merges = dd.plan(rows)
    # All three collapse: the alias table makes "UAE" and "United Arab
    # Emirates" one country, so the spelling cannot hide the duplicate.
    assert len(out) == 1
    assert merges


def test_national_subsidiaries_are_never_merged():
    """The Krannich case. These are five real distributors."""
    rows = [
        _row("Krannich Solar GmbH & Co. KG", "https://krannich-solar.com",
             "Weil der Stadt, Germany"),
        _row("Krannich Solar Turkey", "https://krannich-solar.com",
             "Istanbul, Turkey"),
        _row("Krannich Solar d.o.o.", "https://krannich-solar.com",
             "Zagreb, Croatia"),
        _row("Krannich Solar Energy (Pty) Ltd", "https://krannich-solar.com",
             "Cape Town, South Africa"),
        _row("Krannich Solar Private Limited", "https://krannich-solar.com",
             "Pune, India"),
    ]
    out, merges = dd.plan(rows)
    assert len(out) == 5, "each national entity must survive"
    assert merges == []


def test_a_same_country_variant_merges_while_subsidiaries_survive():
    """Both behaviours at once, which is the real data shape."""
    rows = [
        _row("Krannich Solar GmbH & Co. KG", "https://krannich-solar.com",
             "Weil der Stadt, Germany"),
        _row("Krannich Solar", "https://krannich-solar.com",
             "Weil der Stadt, Germany"),          # variant of the above
        _row("Krannich Solar Turkey", "https://krannich-solar.com",
             "Istanbul, Turkey"),                 # distinct entity
    ]
    out, _ = dd.plan(rows)
    assert len(out) == 2
    countries = {dd.country_of(r) for r in out}
    assert len(countries) == 2


def test_rows_without_a_website_are_never_merged():
    """No domain means no identity evidence, so nothing may be collapsed."""
    rows = [
        _row("Alpha Trading", "", "Nairobi, Kenya"),
        _row("Beta Trading", "", "Nairobi, Kenya"),
    ]
    out, merges = dd.plan(rows)
    assert len(out) == 2
    assert merges == []


# --- merging must never lose a value -------------------------------------


def test_the_most_complete_row_survives():
    rows = [
        _row("Solar Hive", "https://solarhive.com.sg", "Singapore"),
        _row("SolarHive Pte Ltd", "https://solarhive.com.sg", "Singapore",
             Founded="2015", Employees="50-100", Ownership="Private Company"),
    ]
    out, _ = dd.plan(rows)
    assert len(out) == 1
    assert out[0]["Founded"] == "2015"


def test_blanks_are_filled_from_the_siblings():
    """If one variant recorded a contact and another a phone number, the
    survivor must carry both -- merging is not allowed to lose data."""
    rows = [
        _row("Solar Hive", "https://solarhive.com.sg", "Singapore",
             Founded="2015", Employees="50-100", Ownership="Private",
             **{"Contact Person": "Jane Tan"}),
        _row("SolarHive Pte Ltd", "https://solarhive.com.sg", "Singapore",
             Founded="2015", Employees="50-100", Ownership="Private",
             **{"Office No.": "+65 6555 0100"}),
    ]
    out, _ = dd.plan(rows)
    assert len(out) == 1
    assert out[0]["Contact Person"] == "Jane Tan"
    assert out[0]["Office No."] == "+65 6555 0100"


def test_an_existing_value_is_not_overwritten_by_a_sibling():
    rows = [
        _row("A Ltd", "https://a.example", "Lagos, Nigeria",
             Founded="2001", Employees="10-50"),
        _row("A Limited", "https://a.example", "Lagos, Nigeria",
             Founded="1999"),
    ]
    out, _ = dd.plan(rows)
    assert len(out) == 1
    assert out[0]["Founded"] == "2001", "the more complete row wins outright"


def test_numbering_is_resequenced():
    """A gap in the # column looks like a lost row."""
    rows = [
        _row("A", "https://a.example", "Lagos, Nigeria", **{"#": 1}),
        _row("A Ltd", "https://a.example", "Lagos, Nigeria", **{"#": 2}),
        _row("B", "https://b.example", "Accra, Ghana", **{"#": 3}),
    ]
    out, _ = dd.plan(rows)
    assert [r["#"] for r in out] == list(range(1, len(out) + 1))


def test_a_clean_list_is_left_completely_alone():
    rows = [
        _row("A", "https://a.example", "Lagos, Nigeria"),
        _row("B", "https://b.example", "Accra, Ghana"),
        _row("C", "https://c.example", "Nairobi, Kenya"),
    ]
    out, merges = dd.plan(rows)
    assert len(out) == 3
    assert merges == []


def test_no_company_is_ever_invented():
    """Every surviving company name must have been in the input."""
    rows = [
        _row("PowernSun", "https://powernsun.com", "Dubai, UAE"),
        _row("Power & Sun", "https://powernsun.com", "Dubai, UAE"),
        _row("Segen", "https://segensolar.co.za", "Cape Town, South Africa"),
    ]
    source = {r["Company"] for r in rows}
    out, _ = dd.plan(rows)
    for row in out:
        assert row["Company"] in source


# --- country aliases -----------------------------------------------------
#
# Measured in the live data: "UAE" (12 rows) and "United Arab Emirates"
# (65), "USA" (5) and "United States" (26), "UK" (2) and "United Kingdom"
# (13). Country is half the identity key, so treating these as different
# countries lets a same-domain duplicate through under the other
# spelling -- exactly the failure the key exists to prevent.


@pytest.mark.parametrize(
    "a,b",
    [
        ("UAE", "United Arab Emirates"),
        ("USA", "United States"),
        ("U.S.A.", "United States"),
        ("UK", "United Kingdom"),
        ("Great Britain", "United Kingdom"),
        ("Turkiye", "Turkey"),
        ("Holland", "Netherlands"),
        ("Czechia", "Czech Republic"),
        ("KSA", "Saudi Arabia"),
        ("RSA", "South Africa"),
    ],
)
def test_country_aliases_collapse(a, b):
    assert dd.country_of({"Headquarters": f"Somewhere, {a}"}) == \
        dd.country_of({"Headquarters": f"Elsewhere, {b}"})


@pytest.mark.parametrize(
    "a,b",
    [
        ("Turkey", "Germany"),
        ("United States", "United Kingdom"),
        ("Austria", "Australia"),
        ("Niger", "Nigeria"),
        ("South Africa", "South Korea"),
    ],
)
def test_distinct_countries_stay_distinct(a, b):
    """A false merge here deletes a real national distributor, so the
    alias table must not over-collapse."""
    assert dd.country_of({"Headquarters": f"X, {a}"}) != \
        dd.country_of({"Headquarters": f"Y, {b}"})


def test_a_duplicate_written_with_an_alias_is_caught():
    """The concrete gap: the same Dubai company, one row saying UAE and
    the other United Arab Emirates."""
    rows = [
        _row("Fortune Emirates", "https://fortune.example", "Dubai, UAE"),
        _row("Fortune Emirates General Trading", "https://www.fortune.example",
             "Dubai, United Arab Emirates"),
    ]
    out, merges = dd.plan(rows)
    assert len(out) == 1, "the alias spelling must not hide the duplicate"
    assert merges


def test_the_tool_and_the_pipeline_share_one_alias_table():
    """Two tables would drift, and then the cleanup tool and the
    collection-time check would disagree about what is a duplicate."""
    from vendor_intel.pipeline.market_research.state import normalise_country

    for value in ("UAE", "United Arab Emirates", "USA", "Turkey", "Holland"):
        assert dd.country_of({"Headquarters": f"City, {value}"}) == \
            normalise_country(value)
