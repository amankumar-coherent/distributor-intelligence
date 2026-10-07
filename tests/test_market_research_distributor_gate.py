"""Distributor/channel-partner-only gating: manufacturers/OEMs must never
survive into a final batch row, even if the research draft mislabels one.

Per explicit user instruction, batch_runner._build_row no longer calls
verify.py's real-fetch functions (verify_website/verify_is_distributor/
verify_contact_fields/verify_linkedin) — the draft's own field values are
used directly. Those verify.py functions still exist and are independently
tested (test_market_research_verify.py covers verify_contact_fields; the
verify_is_distributor tests below still exercise that function directly,
even though _build_row no longer calls it).
"""
from __future__ import annotations

from unittest.mock import patch

from vendor_intel.pipeline.market_research import batch_runner, verify


def _item(**overrides) -> dict:
    base = {
        "company": "Acme Distribution Inc",
        "website": "https://acmedist.example.com",
        "founded": "1998",
        "headquarters": "Austin, United States",
        "continent_geography": "",
        "operational_presence": "",
        "ownership": "Private Company",
        "employees": "250-500",
        "core_categories": "Semiconductor distribution",
        "specialty_focus": "Analog ICs",
        "key_products_technologies": "Analog ICs",
        "online_offline_both": "Both",
        "contact_person": "",
        "role": "",
        "email": "",
        "linkedin": "",
        "office_no": "",
        "office_email": "",
        "verdict": "distributor",
        "verdict_reason": "Resells third-party semiconductor components.",
    }
    base.update(overrides)
    return base


def test_build_row_rejects_manufacturer_verdict():
    assert batch_runner._build_row(_item(verdict="manufacturer")) is None


def test_build_row_rejects_unknown_verdict():
    assert batch_runner._build_row(_item(verdict="unknown")) is None


def test_build_row_rejects_other_verdict():
    assert batch_runner._build_row(_item(verdict="other")) is None


def test_build_row_rejects_channel_partner_verdict():
    """Channel partners / resellers / VARs / dealers are explicitly NOT
    distributors per the user's requirement — only genuine wholesale or
    authorized distributors count, so a "channel_partner" verdict must be
    rejected just like manufacturer and system_integrator."""
    assert batch_runner._build_row(_item(verdict="channel_partner")) is None


def test_build_row_rejects_system_integrator_verdict():
    """System integrators kept slipping through as "distributor" in
    practice — the prompt now has a dedicated "system_integrator" verdict
    category, and it must be rejected the same as manufacturer/other."""
    assert batch_runner._build_row(_item(verdict="system_integrator")) is None


def test_build_row_rejects_missing_website_even_with_distributor_verdict():
    assert batch_runner._build_row(_item(verdict="distributor", website="")) is None


def test_build_row_uses_draft_fields_directly_without_any_real_fetch():
    """With verification removed, a distributor-verdict draft with a
    website should produce a row built entirely from the draft's own
    values — no call to verify.py's real-fetch functions."""
    item = _item(
        verdict="distributor",
        contact_person="Jane Doe",
        role="VP Sales",
        email="sales@acmedist.example.com",
        linkedin="https://linkedin.com/company/acme-dist",
        office_no="+1 512-555-0100",
        office_email="office@acmedist.example.com",
    )
    with patch.object(verify, "verify_website") as mock_verify_website, patch.object(
        verify, "verify_is_distributor"
    ) as mock_is_distributor, patch.object(
        verify, "verify_contact_fields"
    ) as mock_contact_fields, patch.object(
        verify, "verify_linkedin"
    ) as mock_linkedin:
        row = batch_runner._build_row(item)

    mock_verify_website.assert_not_called()
    mock_is_distributor.assert_not_called()
    mock_contact_fields.assert_not_called()
    mock_linkedin.assert_not_called()

    assert row is not None
    assert row["Company"] == "Acme Distribution Inc"
    assert row["Website"] == "https://acmedist.example.com"
    assert row["Contact Person"] == "Jane Doe"
    assert row["Role"] == "VP Sales"
    assert row["Email"] == "sales@acmedist.example.com"
    # A /company/ page is NOT the contact person's profile and is no longer
    # rebuilt into one — see _real_linkedin_or_blank.
    assert row["LinkedIn"] == ""
    assert row["Office No."] == "+1 512-555-0100"
    assert row["Office Website"] == "https://acmedist.example.com"
    assert row["Office Email"] == "office@acmedist.example.com"


def test_build_row_office_email_stays_blank_when_unset():
    """Office Email no longer falls back to Email. Copying one field into
    another manufactured a value that looked independently researched; per
    explicit user instruction these columns hold real data or nothing."""
    item = _item(verdict="distributor", email="sales@acmedist.example.com", office_email="")
    row = batch_runner._build_row(item)
    assert row["Office Email"] == ""


def test_build_row_keeps_company_when_contact_fields_are_empty():
    """Blanking an unverifiable contact field must NOT drop the company —
    otherwise switching off fabrication would delete the whole dataset."""
    item = _item(
        verdict="distributor",
        contact_person="",
        role="",
        email="",
        linkedin="",
        office_no="",
        office_email="",
    )
    row = batch_runner._build_row(item)

    assert row is not None
    assert row["Company"] == "Acme Distribution Inc"
    assert row["Contact Person"] == ""
    assert row["Email"] == ""


def test_build_row_normalizes_placeholder_contacts_to_blank():
    """"N/A"/"unknown" must become a true blank, or a placeholder string
    would be exported as though it were real data."""
    for junk in ("N/A", "unknown", "none", "TBD", "-"):
        item = _item(verdict="distributor", email=junk, office_no=junk)
        row = batch_runner._build_row(item)
        assert row["Email"] == "", junk
        assert row["Office No."] == "", junk


def test_build_row_converts_bare_employee_number_into_a_range():
    """"about 23,000" / "23,000" style single numbers must become a range
    like "20,000-25,000" — the user explicitly said not to want a bare
    number in the Employees column."""
    for raw in ("23,000", "about 23,000", "~23000", "23000+"):
        item = _item(verdict="distributor", employees=raw)
        row = batch_runner._build_row(item)
        assert row["Employees"] == "20,000-25,000", raw


def test_build_row_leaves_already_ranged_employees_unchanged():
    item = _item(verdict="distributor", employees="10,000-25,000")
    row = batch_runner._build_row(item)
    assert row["Employees"] == "10,000-25,000"


def test_linkedin_is_never_built_from_a_persons_name():
    """The old code turned a name into "/in/<name-slug>", which looked
    researched but was generated — many such URLs point at nobody, or at a
    different person with the same name. Must now be blank instead."""
    assert batch_runner._real_linkedin_or_blank("linkedin.com") == ""
    assert batch_runner._real_linkedin_or_blank("") == ""


def test_linkedin_company_page_is_rejected_not_repaired():
    """A /company/ page is the wrong entity for this column. Rejecting it
    to "" is honest; converting it to a /in/ slug was not."""
    assert (
        batch_runner._real_linkedin_or_blank("https://www.linkedin.com/company/acme") == ""
    )


def test_linkedin_keeps_a_real_profile_url():
    assert (
        batch_runner._real_linkedin_or_blank("https://www.linkedin.com/in/paul-bay/")
        == "https://www.linkedin.com/in/paul-bay"
    )


def test_linkedin_adds_scheme_to_a_bare_profile_url():
    assert (
        batch_runner._real_linkedin_or_blank("linkedin.com/in/paul-bay")
        == "https://www.linkedin.com/in/paul-bay"
    )


def test_linkedin_placeholder_becomes_blank():
    for junk in ("N/A", "unknown", "none", "-"):
        assert batch_runner._real_linkedin_or_blank(junk) == "", junk


def test_normalize_employees_bands_small_and_large_numbers():
    assert batch_runner._normalize_employees("45") == "40-50"
    assert batch_runner._normalize_employees("450") == "400-500"
    assert batch_runner._normalize_employees("4,500") == "4,000-5,000"
    assert batch_runner._normalize_employees("450,000") == "450,000-455,000"
    # Already a range or non-numeric text is left untouched.
    assert batch_runner._normalize_employees("1,000-5,000") == "1,000-5,000"
    assert batch_runner._normalize_employees("Not disclosed") == "Not disclosed"


def test_verify_is_distributor_false_on_manufacturer_self_description():
    fake_result = type("R", (), {"alive": True, "text": "We manufacture premium analog ICs in-house."})()
    with patch(
        "vendor_intel.scraping.httpx_fetch.fetch_page_httpx", return_value=fake_result
    ):
        assert verify.verify_is_distributor("https://acme.example.com") is False


def test_verify_is_distributor_true_when_no_manufacturer_language():
    fake_result = type("R", (), {"alive": True, "text": "We are the leading distributor of analog ICs."})()
    with patch(
        "vendor_intel.scraping.httpx_fetch.fetch_page_httpx", return_value=fake_result
    ):
        assert verify.verify_is_distributor("https://acme.example.com") is True


def test_verify_is_distributor_false_for_empty_website():
    assert verify.verify_is_distributor("") is False
