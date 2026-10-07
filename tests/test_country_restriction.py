"""A market may be restricted to a fixed list of HQ countries.

The existing geo steer only ever PREFERS under-covered countries; it has
no way to say "nowhere else counts". That is right for a global market
and wrong for one defined as "distributors in these 16 countries only",
where a German company is not a weaker answer but an incorrect one.

Enforced twice on purpose:

  * in the PROMPT, so the backend spends its answer on countries that can
    actually be accepted;
  * at COLLECTION (batch_runner), because a prompt is a request and the
    model does not always honour it.

The prompt saves queries. The collection gate is what makes it a promise.
"""
from __future__ import annotations

import pytest

from vendor_intel.pipeline.market_research import discovery

MARKET = "Test Pepper Market - Distributors Only"
ALLOWED = ("Japan", "South Korea", "Vietnam", "India")


@pytest.fixture(autouse=True)
def _restriction():
    """Register the restriction, and always remove it afterwards.

    The registry is module-level state; a leaked entry would silently
    restrict an unrelated test.
    """
    discovery.set_country_restriction(MARKET, ALLOWED)
    yield
    discovery._COUNTRY_RESTRICTIONS.pop(MARKET.strip().lower(), None)


# --- the registry --------------------------------------------------------


def test_a_registered_market_reports_its_countries():
    assert discovery.country_restriction_for(MARKET) == ALLOWED


def test_an_unregistered_market_is_unrestricted():
    assert discovery.country_restriction_for("Some Other Market") == ()


def test_an_unrestricted_market_allows_any_country():
    for country in ("Germany", "Brazil", "Japan", ""):
        assert discovery.country_is_allowed("Some Other Market", country)


# --- the decision --------------------------------------------------------


@pytest.mark.parametrize("country", ["Japan", "South Korea", "Vietnam", "India"])
def test_listed_countries_are_allowed(country):
    assert discovery.country_is_allowed(MARKET, country)


@pytest.mark.parametrize("country", ["Germany", "United States", "Brazil", "Kenya"])
def test_unlisted_countries_are_rejected(country):
    assert not discovery.country_is_allowed(MARKET, country)


def test_matching_ignores_case_and_spacing():
    for spelling in ("japan", "  Japan  ", "JAPAN"):
        assert discovery.country_is_allowed(MARKET, spelling)


def test_common_aliases_match():
    """"Korea", "South Korea" and "Republic of Korea" are one country.

    Without alias handling a real Korean distributor would be thrown away
    for spelling its own country differently.
    """
    assert discovery.country_is_allowed(MARKET, "Korea")
    assert discovery.country_is_allowed(MARKET, "Republic of Korea")


def test_an_unknown_country_is_rejected_not_assumed():
    """Fail closed. A row whose country cannot be established has not been
    shown to satisfy the restriction, and the restriction is the point of
    the market."""
    assert not discovery.country_is_allowed(MARKET, "")
    assert not discovery.country_is_allowed(MARKET, "   ")


# --- the prompt ----------------------------------------------------------


def test_the_prompt_names_every_allowed_country():
    prompt = discovery.build_draft_prompt(
        MARKET, excluded=[], count=10, collected=[])
    for country in ALLOWED:
        assert country in prompt, country


def test_the_prompt_states_the_restriction_as_a_rule():
    prompt = discovery.build_draft_prompt(
        MARKET, excluded=[], count=10, collected=[])
    assert "ONLY these countries count" in prompt
    assert "must NOT be returned" in prompt


def test_the_prompt_drops_the_global_coverage_demand():
    """The two instructions are contradictory; the restriction wins."""
    prompt = discovery.build_draft_prompt(
        MARKET, excluded=[], count=10, collected=[])
    assert "GLOBAL COVERAGE IS A HARD REQUIREMENT" not in prompt
    assert "Latin America" not in prompt


def test_the_steer_stays_inside_the_allowed_set():
    """Steering at a country the gate will reject wastes the query."""
    steer = discovery.format_geo_steer([], market=MARKET)
    focus = steer.split("Focus this batch on:")[1]
    named = [c.strip(" .") for c in focus.split(",")]
    assert set(named) <= set(ALLOWED), named


def test_the_steer_prefers_countries_with_no_companies_yet():
    """Steering at a country that already has companies wastes the query.

    Parsed from the END of the steer rather than a fixed marker: the
    first query and later queries word the restriction differently (the
    full country list is stated once, then dropped for URL budget), but
    both end with the batch's target countries.
    """
    rows = [{"Headquarters": "Tokyo, Japan"}] * 10
    steer = discovery.format_geo_steer(rows, market=MARKET)
    targets = steer.rsplit(":", 1)[1]
    assert "Japan" not in targets, "already covered, should not be re-targeted"
    assert any(c in targets for c in ALLOWED), "must still name real targets"


def test_the_restriction_survives_url_budget_shrinking():
    """The shrink stages blank the geo steer as 'a coverage nudge, not
    correctness'. For a restricted market the steer CARRIES the country
    rule, so dropping it would silently unrestrict the market mid-run."""
    from vendor_intel.scraping.google_ai_mode import (
        MAX_URL_CHARS,
        ai_mode_url_length,
    )

    excluded = [f"Company Number {i} Limited" for i in range(1200)]
    prompt = discovery.build_draft_prompt(
        MARKET, excluded=excluded, count=10,
        collected=[{"Headquarters": f"City {i}, Japan"} for i in range(1200)])

    full = discovery._DRAFT_SYSTEM + "\n\n" + prompt
    assert ai_mode_url_length(full) <= MAX_URL_CHARS, "must still fit the URL"
    # Asserted as a PROPERTY, not a phrase. Later queries state the rule
    # compactly ("ONLY companies headquartered in these countries count")
    # instead of listing every country, because the full list overflowed
    # the URL -- which truncated the JSON reply template and turned four
    # batches into unusable prose. What must never disappear is the rule
    # itself and a set of target countries.
    assert "must NOT be returned" in prompt, (
        "the country restriction must never be shrunk away")
    assert any(c in prompt for c in ALLOWED), (
        "the prompt must still name countries to target")
