"""Validators for repairing a market workbook's geography/contact columns.

Measured on the real GLP1_Receptor_Agonist_Distributors.xlsx (270 rows):
  - Continent / Geography named NO country in 270/270 rows
  - Operational Presence used counts/hand-waving in 59
  - Contact Person was a team or company name in 128 (47%)
  - LinkedIn held 2 real /in/ profiles; 156 were SEARCH URLs
  - Email equalled Office Email in 268 (99%)

The prompt asks for better data; these validators are what guarantee it,
because a plausible invention cannot be distinguished from real data
downstream.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_SPEC = importlib.util.spec_from_file_location(
    "fix_glp1_columns", Path(__file__).resolve().parents[1] / "fix_glp1_columns.py"
)
fx = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(fx)


# --- geography ------------------------------------------------------------


@pytest.mark.parametrize("bad", [
    "Asia", "Europe", "North America", "Africa / Middle East",
    "Asia / Global", "Global", "Worldwide",
    "APAC", "EMEA", "North America / Global",
    "India + 40 countries", "50+ countries", "USA, Canada + international",
    "India + international export markets", "India - nationwide",
    "India; international pharmaceutical buyers",
])
def test_vague_geography_is_rejected(bad):
    """Exactly the values found in the file, each of which names no
    country or substitutes a count for the list."""
    assert not fx.valid_geo(bad), bad


@pytest.mark.parametrize("good", [
    "Asia(India, Singapore); Europe(Germany, United Kingdom)",
    "North America(United States, Canada)",
    "India; Nepal; Sri Lanka; Bangladesh",
    "Asia(India); Middle East(UAE, Saudi Arabia); Africa(Kenya)",
])
def test_country_level_geography_is_accepted(good):
    assert fx.valid_geo(good), good


def test_old_dash_format_is_accepted_because_it_names_a_country():
    """"Europe - United Kingdom" is the file's old malformed format, but it
    DOES name a real country, so rejecting it would discard usable data.
    The parser must strip the region prefix rather than treat the whole
    string as one fake country name."""
    assert fx.countries_in("Europe - United Kingdom") == ["United Kingdom"]
    assert fx.valid_geo("Europe - United Kingdom")


def test_region_slash_global_names_no_country():
    """"Asia / Global" looks similar but names nothing real."""
    assert fx.countries_in("Asia / Global") == []
    assert not fx.valid_geo("Asia / Global")


def test_single_country_is_valid():
    """A regional distributor may genuinely operate in one country.
    Requiring a separator rejected "India" and overwrote 159 correct
    values with NPD in the first real run."""
    assert fx.valid_geo("India")
    assert fx.valid_geo("Asia(India)")


def test_empty_geography_is_not_valid():
    assert not fx.valid_geo("")
    assert not fx.valid_geo("N/A")
    assert not fx.valid_geo("NPD")


# --- contact person -------------------------------------------------------


@pytest.mark.parametrize("bad", [
    "ALDAWLIA Management Team", "A.K. Pharma Management", "IMCD India Management",
    "Actylis Management", "Pharmaforte Sales Team", "Ancanex Pharma Trading Team",
    "Sales / Distribution", "Customer Service", "Suhana", "Krish",
    "Medilink Healthcare", "Skites Pharma Pvt Ltd", "Board of Directors",
])
def test_team_and_company_names_are_rejected(bad):
    """47% of the column was values like these - not usable contacts."""
    assert not fx.is_person_name(bad), bad


@pytest.mark.parametrize("good", [
    "Steven Collis", "Brian Tyler", "Jason Hollar", "David Foell",
    "Diane M. Smith", "Abhishek Sanjaykumar Tiwari",
])
def test_real_person_names_are_accepted(good):
    assert fx.is_person_name(good), good


def test_empty_person_is_rejected():
    assert not fx.is_person_name("")
    assert not fx.is_person_name("N/A")


# --- LinkedIn -------------------------------------------------------------


def test_search_url_is_rejected():
    """156 of 270 rows held a search URL. Nobody researching a contact
    writes one - it is proof of construction."""
    assert fx.clean_linkedin(
        "https://www.linkedin.com/search/results/companies/?keywords=ALDAWLIA"
    ) == ""


def test_company_page_is_rejected():
    assert fx.clean_linkedin("https://www.linkedin.com/company/cencoraglobal/") == ""
    assert fx.clean_linkedin("https://in.linkedin.com/company/aark-pharma") == ""


def test_personal_profile_is_kept():
    assert fx.clean_linkedin("https://www.linkedin.com/in/jane-doe/") == (
        "https://www.linkedin.com/in/jane-doe"
    )


def test_bare_personal_profile_gets_a_scheme():
    assert fx.clean_linkedin("linkedin.com/in/jane-doe") == (
        "https://www.linkedin.com/in/jane-doe"
    )


# --- email ----------------------------------------------------------------


def test_generic_mailbox_at_own_domain_is_rejected():
    """A generic mailbox at the company's own domain is the construction
    being removed: 59% of the column, and indistinguishable from a guess."""
    for local in ("info", "sales", "contact", "rx", "customer.service"):
        assert fx.clean_email(f"{local}@acme.com", "https://acme.com") == "", local


def test_specific_address_at_own_domain_is_kept():
    assert fx.clean_email("j.smith@acme.com", "https://acme.com") == "j.smith@acme.com"


def test_generic_at_a_different_domain_is_kept():
    """Not derived from this website, so the heuristic must not fire."""
    assert fx.clean_email("info@acme-group.net", "https://acme.com") == (
        "info@acme-group.net"
    )


def test_office_email_copied_from_email_is_rejected():
    """Email == Office Email in 268/270 rows - one field copied into the
    other, which is not independent evidence."""
    assert fx.clean_email(
        "j.smith@acme.com", "https://acme.com", other="j.smith@acme.com"
    ) == ""


def test_malformed_email_is_rejected():
    assert fx.clean_email("not-an-email", "https://acme.com") == ""
    assert fx.clean_email("a@b", "https://acme.com") == ""


# --- phone ----------------------------------------------------------------


def test_short_phone_is_rejected():
    for bad in ("123", "+1", "", "n/a"):
        assert fx.clean_phone(bad) == "", bad


def test_real_phone_is_kept():
    assert fx.clean_phone("+1 610 727 7000") == "+1 610 727 7000"


# --- answer alignment -----------------------------------------------------


def test_answers_are_matched_by_name_not_position():
    """The model reorders entries; position-matching would attribute one
    company's contact details to another."""
    rows = [{"Company": "Alpha Ltd"}, {"Company": "Beta Corp"}]
    answers = [
        {"company": "Beta Corp", "office_no": "+44 20 1111 1111"},
        {"company": "Alpha Ltd", "office_no": "+1 555 222 2222"},
    ]
    out = fx._match_by_name(rows, answers)
    assert out["Alpha Ltd"]["office_no"] == "+1 555 222 2222"
    assert out["Beta Corp"]["office_no"] == "+44 20 1111 1111"


def test_unmatched_answer_is_not_forced_onto_a_company():
    rows = [{"Company": "Alpha Ltd"}, {"Company": "Beta Corp"}, {"Company": "Gamma SA"}]
    answers = [{"company": "Someone Else", "office_no": "+1 555 000 0000"}]
    out = fx._match_by_name(rows, answers)
    assert all(v.get("office_no") != "+1 555 000 0000" for v in out.values())


def test_echoed_prompt_template_is_ignored():
    """The prompt ENDS with an empty JSON template, which AI Mode echoes
    back; a naive find on the first bracket locks onto it and returns
    nothing."""
    echoed = 'Reply with ONLY a JSON array...\n[{"company": "", "role": ""}]'
    real = '[{"company": "Acme Inc", "role": "VP Sales"}]'
    got = fx._extract_array(echoed + "\n" + real, "company")
    assert len(got) == 1 and got[0]["role"] == "VP Sales"


def test_first_record_is_not_dropped():
    """Slicing PAST the marker instead of FROM it silently lost the first
    company of every batch in an earlier version of this parser."""
    real = (
        '[{"company": "First Co", "role": "A"}, '
        '{"company": "Second Co", "role": "B"}]'
    )
    got = fx._extract_array('template [{"company": ""}]\n' + real, "company")
    assert [g["company"] for g in got] == ["First Co", "Second Co"]


def test_prompts_forbid_the_observed_failures():
    """The instructions must actually reach the model."""
    # The geo prompt was deliberately SIMPLIFIED: a version dense with
    # quoted forbidden phrases ("Global", "50+ countries", "+ international")
    # made AI Mode treat the request as a SEARCH QUERY and return web
    # results instead of generating an answer (observed live: 0 companies
    # answered twice in a row). The constraints are still stated, just in
    # plain prose.
    geo = fx._GEO_PROMPT.lower()
    assert "never answer with only a region name" in geo
    assert "never with a number of countries" in geo
    assert "global or" in geo and "international" in geo

    con = fx._CONTACT_PROMPT.lower()
    assert "do not hallucinate" in con
    assert "management team" in con
    assert "search url" in con
    assert "do not simply repeat" in con


# --- the two columns are RELATED, not independent -------------------------
#
# Per the user's definitions, restated:
#   Continent / Geography = countries with a physical site
#   Operational Presence  = every country it does business in, WIDER
#
# Observed bug: the model filled continent_geography and returned an EMPTY
# operational_presence, which the validator turned into "NPD" - discarding
# information already in hand. A company necessarily operates wherever it
# has a physical site, so the physical countries are a valid floor.


def test_countries_are_extracted_from_a_region_list():
    cg = "North America(United States, Canada); Europe(Germany, France)"
    assert fx.countries_in(cg) == [
        "United States", "Canada", "Germany", "France",
    ]


def test_countries_are_extracted_from_a_plain_list():
    assert fx.countries_in("India; Nepal; Sri Lanka") == [
        "India", "Nepal", "Sri Lanka",
    ]


def test_region_words_are_not_mistaken_for_countries():
    assert "Europe" not in fx.countries_in("Europe(Germany); Asia(India)")


def test_empty_operational_is_backfilled_from_physical():
    """The reported symptom: only one column filled, the other NPD."""
    cg = "North America(United States, Canada); Asia(India)"
    out = fx.widen_operational(cg, "")
    assert fx.countries_in(out) == ["United States", "Canada", "India"]
    assert fx.valid_geo(out), "the backfilled value must pass validation"


def test_operational_keeps_extra_countries_and_adds_physical_ones():
    """Countries the model supplied must never be dropped by the merge."""
    cg = "Asia(India)"
    out = fx.widen_operational(cg, "Brazil; Mexico")
    got = fx.countries_in(out)
    assert set(got) == {"Brazil", "Mexico", "India"}


def test_operational_is_never_narrower_than_physical():
    cg = "Europe(Germany, France, Italy)"
    out = fx.widen_operational(cg, "Germany")
    assert len(fx.countries_in(out)) >= len(fx.countries_in(cg))


def test_no_duplicate_countries_after_merge():
    cg = "Asia(India, Singapore)"
    out = fx.widen_operational(cg, "India; Singapore; Japan")
    got = [c.casefold() for c in fx.countries_in(out)]
    assert len(got) == len(set(got)), got


def test_both_empty_stays_empty():
    """Nothing known must not be invented into a country list."""
    assert fx.widen_operational("", "") == ""


def test_geo_prompt_states_both_definitions():
    """The user supplied these definitions explicitly; they must reach the
    model, and the wider/narrower relationship must be stated."""
    geo = fx._GEO_PROMPT.lower()
    assert "physical site" in geo
    assert "wider" in geo
    assert "does business" in geo
    assert "every country individually" in geo


# --- transient refusals must be retried, not skipped ---------------------
#
# Google refuses BURSTS of requests ("It looks like there's no response
# available for this search. Try asking something else.") and recovers on
# its own - proven by re-running the identical prompt for the identical
# company and getting a 52-country answer. Giving up after one refusal
# discarded 4 consecutive batches for no reason.


def test_batch_is_retried_after_a_refusal(monkeypatch):
    calls = {"n": 0}

    def _fake_chat(_sys, _prompt):
        calls["n"] += 1
        if calls["n"] < 3:
            return "It looks like there's no response available for this search."
        return '[{"company": "Acme", "continent_geography": "Asia(India)", "operational_presence": "India"}]'

    monkeypatch.setattr(fx.ai_mode_client, "chat", _fake_chat)
    monkeypatch.setattr(fx.time, "sleep", lambda _s: None)

    out = fx._ask("sys", "prompt", "company")

    assert out and out[0]["company"] == "Acme"
    assert calls["n"] == 3, "must keep retrying until an answer arrives"


def test_batch_retries_are_bounded(monkeypatch):
    """A permanently refused prompt must not retry forever."""
    calls = {"n": 0}

    def _always_refuse(_sys, _prompt):
        calls["n"] += 1
        return "It looks like there's no response available for this search."

    monkeypatch.setattr(fx.ai_mode_client, "chat", _always_refuse)
    monkeypatch.setattr(fx.time, "sleep", lambda _s: None)

    assert fx._ask("sys", "prompt", "company") == []
    assert calls["n"] == fx._BATCH_RETRIES


def test_batch_retries_an_exception_too(monkeypatch):
    """AIModeUnavailable is raised for a CAPTCHA/timeout; that is also
    transient and must not kill the batch on the first occurrence."""
    calls = {"n": 0}

    def _boom_then_ok(_sys, _prompt):
        calls["n"] += 1
        if calls["n"] == 1:
            raise RuntimeError("captcha")
        return '[{"company": "Acme", "continent_geography": "Asia(India)", "operational_presence": "India"}]'

    monkeypatch.setattr(fx.ai_mode_client, "chat", _boom_then_ok)
    monkeypatch.setattr(fx.time, "sleep", lambda _s: None)

    assert fx._ask("sys", "prompt", "company")
    assert calls["n"] == 2


def test_batch_backoff_grows(monkeypatch):
    """An immediate retry usually returns the same refusal, so the wait
    must increase rather than hammering Google."""
    monkeypatch.setattr(
        fx.ai_mode_client, "chat",
        lambda _s, _p: "It looks like there's no response available for this search.",
    )
    slept = []
    monkeypatch.setattr(fx.time, "sleep", lambda s: slept.append(s))

    fx._ask("sys", "prompt", "company")

    assert len(slept) == fx._BATCH_RETRIES - 1
    assert slept == sorted(slept), f"backoff must not shrink: {slept}"
    assert slept[0] > 0


def test_first_attempt_success_does_not_sleep(monkeypatch):
    """The happy path must stay fast - no wasted backoff."""
    monkeypatch.setattr(
        fx.ai_mode_client, "chat",
        lambda _s, _p: '[{"company": "Acme", "continent_geography": "Asia(India)", "operational_presence": "India"}]',
    )
    slept = []
    monkeypatch.setattr(fx.time, "sleep", lambda s: slept.append(s))

    assert fx._ask("sys", "prompt", "company")
    assert slept == []
