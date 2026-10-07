"""contact_lookup.py: re-asking AI Mode for real contact data only.

The original discovery prompt required all 21 columns non-blank, which
forced construction of contact details (94% of exported Office Email
values were info@/sales@ at the company's own domain, LinkedIn URLs were
built from names, phones from country calling codes). This module re-asks
with explicit permission to say "unknown".

These tests pin the anti-fabrication rules, because the whole value of the
second pass is that it does NOT re-accept a constructed value.
"""
from __future__ import annotations

import json
from unittest.mock import patch

import pytest

from vendor_intel.pipeline.market_research import contact_lookup


def _row(company="Acme Inc", website="https://acme.com", hq="Austin, United States"):
    return {"Company": company, "Website": website, "Headquarters": hq}


# --- "unknown" recognition -----------------------------------------------


@pytest.mark.parametrize(
    "value",
    [
        "", "-", "N/A", "n/a", "na", "None", "null", "nil", "unknown",
        "Not known", "not available", "Not Found", "not published",
        "not disclosed", "no information found", "no email available",
        "TBD", "undisclosed", "Not applicable", "  N/A  ", "unknown.",
    ],
)
def test_unknown_phrasings_are_all_treated_as_blank(value):
    """The model expresses "I don't know" many ways; every one must become
    a blank, or prose like "not publicly available" gets exported as data."""
    assert contact_lookup.is_unknown(value), value


@pytest.mark.parametrize(
    "value",
    ["j.smith@acme.com", "Jane Doe", "VP Sales", "+1 512 555 0100", "Nanjing"],
)
def test_real_values_are_not_treated_as_unknown(value):
    assert not contact_lookup.is_unknown(value)


# --- constructed-email detection -----------------------------------------


@pytest.mark.parametrize("local", ["info", "sales", "contact", "support", "enquiry"])
def test_generic_mailbox_at_own_domain_is_rejected(local):
    """This is the exact pattern the old prompt was told to construct
    ("website acme.com -> info@acme.com"), so it cannot be distinguished
    from an invention and must not be accepted as verified."""
    assert contact_lookup._looks_constructed_email(f"{local}@acme.com", "https://acme.com")


def test_specific_address_at_own_domain_is_kept():
    """A named mailbox is not the constructed pattern — the model would
    have to actually know it."""
    assert not contact_lookup._looks_constructed_email("j.smith@acme.com", "https://acme.com")


def test_generic_mailbox_at_a_different_domain_is_kept():
    """info@ at some OTHER domain wasn't derived from this website, so the
    construction heuristic must not fire."""
    assert not contact_lookup._looks_constructed_email("info@acme-group.net", "https://acme.com")


def test_constructed_check_handles_subdomains_and_www():
    assert contact_lookup._looks_constructed_email("info@mail.acme.com", "https://www.acme.com")


def test_constructed_check_ignores_empty_input():
    assert not contact_lookup._looks_constructed_email("", "https://acme.com")
    assert not contact_lookup._looks_constructed_email("info@acme.com", "")


# --- sanitisation ---------------------------------------------------------


def test_sanitize_blanks_unknown_values():
    answer = {
        "contact_person": "unknown", "role": "N/A", "email": "not available",
        "linkedin": "-", "office_no": "none", "office_email": "",
    }
    out = contact_lookup._sanitize(answer, "https://acme.com", strict=True)
    assert all(v == "" for v in out.values()), out


def test_sanitize_keeps_a_real_specific_email():
    answer = {"email": "j.smith@acme.com", "contact_person": "", "role": "",
              "linkedin": "", "office_no": "", "office_email": ""}
    out = contact_lookup._sanitize(answer, "https://acme.com", strict=True)
    assert out["email"] == "j.smith@acme.com"


def test_sanitize_rejects_generic_email_in_strict_mode():
    answer = {"email": "info@acme.com", "contact_person": "", "role": "",
              "linkedin": "", "office_no": "", "office_email": "info@acme.com"}
    out = contact_lookup._sanitize(answer, "https://acme.com", strict=True)
    assert out["email"] == ""
    assert out["office_email"] == ""


def test_sanitize_keeps_generic_email_when_not_strict():
    """--allow-generic-email opts back in deliberately."""
    answer = {"email": "info@acme.com", "contact_person": "", "role": "",
              "linkedin": "", "office_no": "", "office_email": ""}
    out = contact_lookup._sanitize(answer, "https://acme.com", strict=False)
    assert out["email"] == "info@acme.com"


def test_sanitize_rejects_malformed_email():
    answer = {"email": "not-an-email", "contact_person": "", "role": "",
              "linkedin": "", "office_no": "", "office_email": "a@b"}
    out = contact_lookup._sanitize(answer, "https://acme.com", strict=True)
    assert out["email"] == ""
    assert out["office_email"] == ""  # no TLD


def test_sanitize_requires_person_and_role_together():
    """A title with nobody attached, or a name with no role, is not a
    usable contact and half of it is likely invented."""
    out = contact_lookup._sanitize(
        {"contact_person": "Jane Doe", "role": "", "email": "", "linkedin": "",
         "office_no": "", "office_email": ""},
        "https://acme.com", strict=True,
    )
    assert out["contact_person"] == "" and out["role"] == ""


def test_sanitize_drops_linkedin_without_a_named_person():
    """A profile URL with no person can't be attributed to anyone."""
    out = contact_lookup._sanitize(
        {"contact_person": "", "role": "", "email": "",
         "linkedin": "https://www.linkedin.com/in/someone",
         "office_no": "", "office_email": ""},
        "https://acme.com", strict=True,
    )
    assert out["linkedin"] == ""


def test_sanitize_rejects_linkedin_company_page():
    out = contact_lookup._sanitize(
        {"contact_person": "Jane Doe", "role": "VP Sales", "email": "",
         "linkedin": "https://www.linkedin.com/company/acme",
         "office_no": "", "office_email": ""},
        "https://acme.com", strict=True,
    )
    assert out["linkedin"] == ""


def test_sanitize_keeps_and_normalises_a_real_profile_url():
    out = contact_lookup._sanitize(
        {"contact_person": "Jane Doe", "role": "VP Sales", "email": "",
         "linkedin": "linkedin.com/in/jane-doe", "office_no": "", "office_email": ""},
        "https://acme.com", strict=True,
    )
    assert out["linkedin"] == "https://www.linkedin.com/in/jane-doe"


@pytest.mark.parametrize("phone", ["123", "+1", "55", "abc"])
def test_sanitize_rejects_too_short_phone(phone):
    """A real number has enough digits; a fragment is noise."""
    out = contact_lookup._sanitize(
        {"office_no": phone, "contact_person": "", "role": "", "email": "",
         "linkedin": "", "office_email": ""},
        "https://acme.com", strict=True,
    )
    assert out["office_no"] == ""


def test_sanitize_keeps_a_real_phone():
    out = contact_lookup._sanitize(
        {"office_no": "+1 512-555-0100", "contact_person": "", "role": "",
         "email": "", "linkedin": "", "office_email": ""},
        "https://acme.com", strict=True,
    )
    assert out["office_no"] == "+1 512-555-0100"


# --- batching and answer alignment ---------------------------------------


def test_batch_size_is_fifteen():
    """Explicitly requested: 15 companies at a time."""
    assert contact_lookup.BATCH_SIZE == 15


def test_lookup_matches_answers_to_companies_by_name():
    rows = [_row("Alpha Ltd", "https://alpha.com"), _row("Beta Corp", "https://beta.com")]
    reply = json.dumps([
        {"company": "Beta Corp", "contact_person": "Bob Roe", "role": "VP Sales",
         "email": "b.roe@beta.com", "linkedin": "", "office_no": "", "office_email": ""},
        {"company": "Alpha Ltd", "contact_person": "Ann Ray", "role": "Head of Sales",
         "email": "a.ray@alpha.com", "linkedin": "", "office_no": "", "office_email": ""},
    ])
    with patch.object(contact_lookup.ai_mode_client, "chat", return_value=reply):
        out = contact_lookup.lookup_contacts(rows)

    # Answers came back in the WRONG order; matching by name must prevent
    # writing Beta's contact onto Alpha.
    assert out["Alpha Ltd"]["email"] == "a.ray@alpha.com"
    assert out["Beta Corp"]["email"] == "b.roe@beta.com"


def test_lookup_does_not_misalign_when_an_answer_is_missing():
    """If the model returns fewer entries and names don't match, falling
    back to position would attribute one company's details to another."""
    rows = [_row("Alpha Ltd"), _row("Beta Corp"), _row("Gamma SA")]
    reply = json.dumps([
        {"company": "Totally Different Co", "contact_person": "X Y", "role": "VP",
         "email": "x@other.com", "linkedin": "", "office_no": "", "office_email": ""},
    ])
    with patch.object(contact_lookup.ai_mode_client, "chat", return_value=reply):
        out = contact_lookup.lookup_contacts(rows)

    for company in ("Alpha Ltd", "Beta Corp", "Gamma SA"):
        assert out.get(company, {}).get("email", "") != "x@other.com"


def test_lookup_ignores_echoed_prompt_template():
    """AI Mode echoes the question, and this prompt ENDS with an empty JSON
    template — a naive find("[") locks onto it and yields nothing (the same
    bug already fixed in discovery.py)."""
    echoed = 'Reply with ONLY a JSON array...\n[{"company": "", "contact_person": ""}]'
    real = json.dumps([
        {"company": "Acme Inc", "contact_person": "Jane Doe", "role": "VP Sales",
         "email": "j.doe@acme.com", "linkedin": "", "office_no": "", "office_email": ""},
    ])
    with patch.object(contact_lookup.ai_mode_client, "chat", return_value=echoed + "\n" + real):
        out = contact_lookup.lookup_contacts([_row()])
    assert out["Acme Inc"]["email"] == "j.doe@acme.com"


def test_lookup_returns_empty_when_ai_mode_unavailable():
    """A CAPTCHA/timeout must not be mistaken for "no contact exists" —
    the caller leaves those rows unchanged rather than blanking them."""
    with patch.object(
        contact_lookup.ai_mode_client, "chat", side_effect=RuntimeError("captcha")
    ):
        assert contact_lookup.lookup_contacts([_row()]) == {}


def test_lookup_returns_empty_on_unparseable_answer():
    with patch.object(contact_lookup.ai_mode_client, "chat", return_value="no idea, sorry"):
        assert contact_lookup.lookup_contacts([_row()]) == {}


def test_lookup_handles_empty_input():
    assert contact_lookup.lookup_contacts([]) == {}


def test_prompt_forbids_fabrication_explicitly():
    """The instruction the user asked for must actually reach the model."""
    prompt = contact_lookup._PROMPT
    low = prompt.lower()
    assert "do not hallucinate" in low
    assert "forbidden" in low
    assert "never invent a phone number" in low
    assert "never build a linkedin url" in low
    assert '""' in prompt  # empty-string answer is spelled out


def test_prompt_includes_every_company_in_the_batch():
    rows = [_row(f"Company {i}") for i in range(15)]
    text = contact_lookup._format_companies(rows)
    for i in range(15):
        assert f"Company {i}" in text
