"""verify.py: _run_async loop lifecycle and contact-field fetch behavior.

No real network call is made — extract_contact_from_website /
extract_leadership_from_website are mocked throughout.
"""
from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock, patch

from vendor_intel.pipeline.market_research import verify


def test_run_async_executes_coroutine_and_returns_result():
    async def _coro():
        return "done"

    assert verify._run_async(_coro()) == "done"


def test_run_async_can_be_called_many_times_in_a_row_without_error():
    """Regression: calling asyncio.run() repeatedly (once/twice per company,
    across dozens of companies per batch) surfaced a noisy unretrieved
    "RuntimeError: Event loop is closed" from httpx's AsyncClient cleanup
    task firing after its loop had already been torn down (Windows
    ProactorEventLoop). _run_async must tolerate being called many times in
    a tight loop without that happening."""
    async def _coro(n):
        await asyncio.sleep(0)
        return n

    results = [verify._run_async(_coro(i)) for i in range(25)]
    assert results == list(range(25))


def test_verify_contact_fields_uses_single_run_async_call():
    """verify_contact_fields must make exactly one _run_async call per
    company (contact + leadership-fallback combined into one coroutine),
    not two separate event-loop lifecycles — halves the loop churn that
    caused the Windows asyncio cleanup bug."""
    with patch.object(verify, "_run_async", wraps=verify._run_async) as mock_run_async, patch(
        "vendor_intel.pipeline.market_research.verify.extract_contact_from_website",
        new=AsyncMock(return_value={"email": "sales@acme.com", "phone": "+1 555-0100"}),
    ), patch(
        "vendor_intel.pipeline.market_research.verify.extract_leadership_from_website",
        new=AsyncMock(return_value={}),
    ):
        result = verify.verify_contact_fields("https://acme.example.com")

    assert mock_run_async.call_count == 1
    assert result["Email"] == "sales@acme.com"
    assert result["Office No."] == "+1 555-0100"


def test_verify_contact_fields_falls_back_to_leadership_for_acceptable_role():
    with patch(
        "vendor_intel.pipeline.market_research.verify.extract_contact_from_website",
        new=AsyncMock(return_value={"contact_person": "Alex CEO", "contact_role": "CEO"}),
    ), patch(
        "vendor_intel.pipeline.market_research.verify.extract_leadership_from_website",
        new=AsyncMock(return_value={"contact_person": "Jane Doe", "contact_role": "VP Sales"}),
    ):
        result = verify.verify_contact_fields("https://acme.example.com")

    # CEO isn't an acceptable role (sales/procurement/purchasing/commercial/
    # channel only) so it falls through to the leadership page's VP Sales.
    assert result["Contact Person"] == "Jane Doe"
    assert result["Role"] == "VP Sales"


def test_verify_contact_fields_no_contact_person_when_no_acceptable_role_found():
    with patch(
        "vendor_intel.pipeline.market_research.verify.extract_contact_from_website",
        new=AsyncMock(return_value={"contact_person": "Alex CEO", "contact_role": "CEO"}),
    ), patch(
        "vendor_intel.pipeline.market_research.verify.extract_leadership_from_website",
        new=AsyncMock(return_value={"contact_person": "Sam Founder", "contact_role": "Founder"}),
    ):
        result = verify.verify_contact_fields("https://acme.example.com")

    assert "Contact Person" not in result
    assert "Role" not in result


def test_verify_contact_fields_empty_website_returns_empty_dict():
    assert verify.verify_contact_fields("") == {}


def test_verify_linkedin_rejects_company_page_url():
    """The LinkedIn column is the CONTACT PERSON's own profile, not the
    company's page — a /company/ URL must be rejected even though it's a
    valid-looking linkedin.com URL."""
    assert verify.verify_linkedin("https://www.linkedin.com/company/acme-inc/", "Jane Doe") == ""


def test_verify_linkedin_accepts_personal_profile_url_on_identity_match():
    fake_result = type(
        "R", (), {"alive": True, "title": "Jane Doe - VP Sales - Acme", "text": "Jane Doe is VP of Sales"}
    )()
    with patch(
        "vendor_intel.scraping.httpx_fetch.fetch_page_httpx", return_value=fake_result
    ):
        result = verify.verify_linkedin("https://www.linkedin.com/in/jane-doe/", "Jane Doe")
    assert result == "https://www.linkedin.com/in/jane-doe/"


def test_verify_linkedin_rejects_when_name_does_not_match_profile():
    fake_result = type(
        "R", (), {"alive": True, "title": "John Smith - CEO - Other Co", "text": "John Smith is CEO"}
    )()
    with patch(
        "vendor_intel.scraping.httpx_fetch.fetch_page_httpx", return_value=fake_result
    ):
        result = verify.verify_linkedin("https://www.linkedin.com/in/john-smith/", "Jane Doe")
    assert result == ""


def test_verify_linkedin_rejects_empty_url():
    assert verify.verify_linkedin("", "Jane Doe") == ""
