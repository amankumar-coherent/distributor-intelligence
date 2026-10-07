"""Operational Presence must include every country with a physical site.

The two columns have definitions the user gave explicitly:

  * Continent / Geography = countries with a PHYSICAL SITE (HQ, office,
    warehouse, plant)
  * Operational Presence  = every country it does business in, WIDER

A country with a physical site is necessarily one it does business in, so
Operational Presence listing FEWER countries is impossible rather than
merely unusual. Found in the finished Dry Eyes file:

    Zuellig Pharma Vietnam
      CG: APAC(Vietnam, Singapore, Philippines, Thailand, Malaysia, Indonesia)
      OP: Vietnam; Cambodia; Laos        <- 3 countries, 6 sites

The repair is a union, so it invents nothing: every country added is
already stated in the other column.

Most of this file guards the PARSER, because my first version of it was
wrong in the dangerous direction. `_REGIONS` did not contain the short
forms "ME" and "AF" that the live data actually uses, so a dry run
reported 132 rows to "fix" and would have appended ME and AF as if they
were countries -- corrupting 130 good rows to fix 2 bad ones. A parser
that over-reports is worse here than no parser at all.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location(
    "fix_geography_superset", _ROOT / "fix_geography_superset.py"
)
fx = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(fx)

CG, OP = fx.CG, fx.OP


# --- the parser, which must never invent a country -----------------------


@pytest.mark.parametrize(
    "label",
    ["ME", "AF", "EU", "NA", "SA", "AS", "APAC", "EMEA", "LATAM", "MEA",
     "ANZ", "SEA", "GCC", "MENA", "APJ", "ROW"],
)
def test_region_acronyms_are_not_countries(label):
    """The 132-row near-miss. "ME" and "AF" were missing from _REGIONS, so
    they parsed as countries and would have been appended to 130 rows
    that were already correct."""
    assert fx.countries_in(f"{label}(Bahrain)") == ["Bahrain"]
    assert label not in fx.countries_in(f"{label}(Bahrain)")


@pytest.mark.parametrize(
    "label",
    ["Africa", "Europe", "Asia", "Middle East", "North America",
     "Southeast Asia", "Sub-Saharan Africa", "Global", "Worldwide"],
)
def test_spelled_out_regions_are_not_countries(label):
    assert fx.countries_in(f"{label}(Kenya, Uganda)") == ["Kenya", "Uganda"]


def test_the_region_wrapped_format_is_parsed():
    got = fx.countries_in("APAC(India, Singapore, Japan); EU(Germany, France)")
    assert got == ["India", "Singapore", "Japan", "Germany", "France"]


def test_the_flat_format_is_parsed():
    assert fx.countries_in("India; Singapore; Germany") == [
        "India", "Singapore", "Germany"]


def test_mixed_formats_are_parsed():
    got = fx.countries_in("ME(Bahrain); Saudi Arabia; AF(Kenya)")
    assert got == ["Bahrain", "Saudi Arabia", "Kenya"]


def test_duplicates_collapse_case_insensitively():
    assert fx.countries_in("India; india; INDIA") == ["India"]


def test_multiword_countries_survive():
    got = fx.countries_in("ME(United Arab Emirates, Saudi Arabia); South Africa")
    assert "United Arab Emirates" in got
    assert "Saudi Arabia" in got
    assert "South Africa" in got


def test_empty_and_junk_parse_to_nothing():
    for value in ("", None, ";", " ; , ; ", "APAC", "EMEA; APAC"):
        assert fx.countries_in(value) == [], repr(value)


# --- the widening rule ---------------------------------------------------


def test_a_narrower_operational_presence_is_widened():
    """The exact live row."""
    row = {
        CG: "APAC(Vietnam, Singapore, Philippines, Thailand, Malaysia, Indonesia)",
        OP: "Vietnam; Cambodia; Laos",
    }
    got = fx.widen(row)
    assert got is not None
    countries = {c.lower() for c in fx.countries_in(got)}
    # every site country is now present
    for site in ("vietnam", "singapore", "philippines", "thailand",
                 "malaysia", "indonesia"):
        assert site in countries, site
    # and the sell-only countries are NOT lost
    assert "cambodia" in countries
    assert "laos" in countries


def test_an_already_wider_value_is_left_alone():
    """Selling somewhere with no office is entirely normal and must not be
    'corrected'."""
    row = {CG: "AF(South Africa, Kenya)",
           OP: "South Africa; Kenya; Namibia; Botswana"}
    assert fx.widen(row) is None


def test_an_equal_value_is_left_alone():
    row = {CG: "APAC(India)", OP: "India"}
    assert fx.widen(row) is None


def test_a_region_only_site_value_is_left_alone():
    """No parseable country means no basis for widening."""
    assert fx.widen({CG: "APAC", OP: "India"}) is None
    assert fx.widen({CG: "", OP: "India"}) is None


def test_nothing_is_invented():
    """Every country in the result must come from one of the two inputs --
    the property that makes this safe to run without a model call."""
    row = {CG: "EU(Switzerland); APAC(Thailand, Cambodia)",
           OP: "Thailand; Australia"}
    got = fx.widen(row)
    source = {c.lower() for c in
              fx.countries_in(row[CG]) + fx.countries_in(row[OP])}
    for c in fx.countries_in(got):
        assert c.lower() in source, c


def test_existing_order_is_preserved():
    """Reshuffling would make a reviewed value look changed when it is
    not."""
    row = {CG: "APAC(Japan)", OP: "India; Singapore"}
    got = fx.widen(row)
    assert got.startswith("India; Singapore")


def test_continent_geography_is_never_modified():
    """Removing a site claim would be a factual assertion this script has
    no basis for. Only OP is widened."""
    src = (_ROOT / "fix_geography_superset.py").read_text(encoding="utf-8")
    fn = src[src.index("def process("):]
    assert "row[OP] = new" in fn
    assert "row[CG] =" not in fn


# --- the live data -------------------------------------------------------


def test_the_real_dry_eyes_state_has_only_the_two_known_rows():
    """Guards the parser against over-reporting on real data: the first
    version claimed 132 rows needed fixing when the true answer was 2."""
    import json

    from vendor_intel.pipeline.market_research.state import STATE_DIR, market_slug

    market = "Global Dry Eyes Treatment Market - Distributors Only"
    path = STATE_DIR / f"{market_slug(market)}.json"
    if not path.is_file():
        pytest.skip("state file not present in this checkout")

    companies = json.loads(path.read_text(encoding="utf-8")).get("companies") or []
    if not companies:
        pytest.skip("state file is empty")

    needing = [r for r in companies if fx.widen(r)]
    assert len(needing) <= 5, (
        f"{len(needing)} of {len(companies)} rows flagged — the parser is "
        "probably treating a region label as a country again"
    )
