"""Strict per-field verification: only a real fetched source can confirm a
field. A ChatGPT-OAuth draft answer is never trusted on its own.

NOTE: as of the current default flow, batch_runner.py does NOT call these
functions (per explicit user instruction to trust the research draft
directly instead) — this module is kept for anyone who wants to re-enable
real-source verification, and is independently tested.

Reuses, rather than reimplements:
  - website identity/reachability -> pipeline.website_verification.verify_and_fix_website
  - contact/email/phone/named contact -> enrichment.gap_fill.contact_page
    (fetches the company's real Contact/About/Leadership pages)
  - LinkedIn URL shape -> same pattern as heuristic_extract._LINKEDIN_RE,
    plus a real fetch + identity check before acceptance. Per explicit user
    instruction the LinkedIn column is the CONTACT PERSON's own personal
    profile ("/in/<slug>"), not the company's page ("/company/<slug>") —
    verify_linkedin only accepts and identity-checks a /in/ URL against the
    contact person's name.

Any field verify.py cannot confirm this way is left out of the returned
dict entirely (not filled with ""/N/A) — schema.finalize_row treats a
missing key as unverified and drops the whole company.
"""
from __future__ import annotations

import asyncio
import re
from typing import Any

from vendor_intel.enrichment.gap_fill.contact_page import (
    extract_contact_from_website,
    extract_leadership_from_website,
)
from vendor_intel.pipeline.website_verification import verify_and_fix_website

# Per explicit user instruction, the LinkedIn column is the CONTACT
# PERSON's own profile ("/in/<slug>"), not the company's page
# ("/company/<slug>") — this must reject a company-page URL, not just
# validate "is this linkedin.com at all".
_LINKEDIN_RE = re.compile(
    r"^https?://(?:www\.)?linkedin\.com/in/[A-Za-z0-9\-_%]+/?$", re.I
)

_ACCEPTABLE_ROLE = re.compile(
    r"\b(sales|procurement|purchasing|commercial|channel)\b", re.I
)

# Deterministic cross-check against the research draft's self-reported
# "verdict" — a company whose own website plainly describes itself as a
# manufacturer/OEM/producer is excluded even if the LLM draft said
# "distributor" (mirrors discovery.kompass_distributors' hard rule: any
# manufacturer/producer signal on the company's own site disqualifies it,
# never trusting the LLM's verdict alone for a company/manufacturer
# distinction this consequential).
_MANUFACTURER_SELF_DESC = re.compile(
    r"\b(we manufacture|we are a manufacturer|leading manufacturer|"
    r"original equipment manufacturer\b|\bOEM\b|we produce|our factory|"
    r"our manufacturing facilit|in-house manufacturing|"
    r"design and manufacture)\b",
    re.I,
)


def _run_async(coro):
    """Runs one coroutine to completion via a fresh event loop.

    Deliberately NOT asyncio.run(): asyncio.run() closes the loop
    immediately after the coroutine returns, before httpx.AsyncClient's own
    connection-cleanup task (scheduled inside its __aexit__) gets a chance
    to run — on Windows' ProactorEventLoop this surfaces as a noisy
    unretrieved "RuntimeError: Event loop is closed" exception once this
    function is called many times in a row (once or twice per company,
    across dozens of companies per batch). Explicitly running
    loop.shutdown_asyncgens() before closing gives any such pending cleanup
    a chance to finish on the same loop it was scheduled on.

    Also restores the thread's previous event loop afterwards. Without
    that, the closed loop stayed installed as the thread's current loop,
    and Patchright's sync API -- which refuses to run when it detects an
    event loop on the calling thread -- failed every subsequent query with
    "It looks like you are using Playwright Sync API inside the asyncio
    loop. Please use the Async API instead." AI Mode and this verifier run
    on the same thread, so that error was permanent once triggered: every
    later query failed identically until the process was killed.
    """
    try:
        previous = asyncio.get_event_loop_policy().get_event_loop()
    except Exception:  # noqa: BLE001
        previous = None

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        return loop.run_until_complete(coro)
    finally:
        try:
            loop.run_until_complete(loop.shutdown_asyncgens())
        except Exception:
            pass
        loop.close()
        # Never leave a CLOSED loop installed on the thread. Restore the
        # original if there was a usable one, otherwise clear it, so a
        # following sync-Playwright call sees no loop at all.
        try:
            keep = previous if (previous is not None
                                and not previous.is_closed()) else None
            asyncio.set_event_loop(keep)
        except Exception:  # noqa: BLE001
            pass


def verify_website(company: str, headquarters: str, draft_website: str) -> str:
    """Returns the verified official website, or "" if none could be confirmed."""
    result = verify_and_fix_website(company, headquarters, draft_website)
    if result.is_reachable and result.is_identity_match:
        return result.website
    return ""


def verify_is_distributor(website: str) -> bool:
    """True unless the company's own verified website plainly self-describes
    as a manufacturer/OEM/producer. Deliberately lenient (False negatives —
    missing the phrase — are cheaper than wrongly excluding a real
    distributor): this is a safety net on top of the research draft's
    "verdict" field, not the sole gate, so a site with no such language
    passes through to the LLM verdict check instead of being auto-accepted
    here."""
    if not website:
        return False
    try:
        from vendor_intel.scraping.httpx_fetch import fetch_page_httpx

        result = fetch_page_httpx(website)
    except Exception:
        return True
    if not result.alive:
        return True
    return not _MANUFACTURER_SELF_DESC.search(result.text or "")


async def _fetch_contact_fields(row: dict[str, Any]) -> dict[str, Any]:
    """Contact page first, leadership page only if contact page didn't
    yield an acceptable sales/procurement person — both awaited on the SAME
    event loop (single _run_async call from the sync caller) rather than
    two separate asyncio.run()-style loop lifecycles per company."""
    contact = await extract_contact_from_website(row)
    person = contact.get("contact_person", "")
    role = contact.get("contact_role", "")
    if not (person and role and _ACCEPTABLE_ROLE.search(role)):
        leadership = await extract_leadership_from_website(row)
        lp, lr = leadership.get("contact_person", ""), leadership.get("contact_role", "")
        if lp and lr and _ACCEPTABLE_ROLE.search(lr):
            contact = {**contact, "contact_person": lp, "contact_role": lr}
        else:
            contact = {**contact, "contact_person": "", "contact_role": ""}
    return contact


def verify_contact_fields(website: str) -> dict[str, str]:
    """Fetches the verified website's real Contact/About pages and returns
    only fields actually found there: email, phone (-> "Office No."),
    contact_person, role. Missing keys mean "not verifiable", not empty."""
    if not website:
        return {}
    domain = website.removeprefix("https://").removeprefix("http://").removeprefix("www.").split("/")[0]
    row = {"website": domain}

    contact = _run_async(_fetch_contact_fields(row))
    out: dict[str, str] = {}
    if contact.get("email"):
        out["Email"] = contact["email"]
    if contact.get("phone"):
        out["Office No."] = contact["phone"]
    if contact.get("contact_person") and contact.get("contact_role"):
        out["Contact Person"] = contact["contact_person"]
        out["Role"] = contact["contact_role"]

    return out


def verify_linkedin(draft_linkedin: str, contact_person: str) -> str:
    """Only accepts a real linkedin.com/in/... personal profile URL
    (never a /company/ page), fetched and identity-checked against the
    CONTACT PERSON's name — never a search-result URL, per spec section 9."""
    url = (draft_linkedin or "").strip()
    if not url or not _LINKEDIN_RE.match(url):
        return ""
    try:
        from vendor_intel.scraping.httpx_fetch import fetch_page_httpx

        result = fetch_page_httpx(url)
    except Exception:
        return ""
    if not result.alive:
        return ""
    from vendor_intel.pipeline.website_verification import _looks_like_identity_match

    if _looks_like_identity_match(contact_person, url, result.title, result.text):
        return url
    return ""


def verify_office_website(website: str) -> str:
    """This pipeline researches one global entity per row (no distinct
    regional office pages discovered), so Office Website is the verified
    company website itself when one exists — per spec section 9 ("if there
    is no separate office website, use the appropriate official company
    website")."""
    return website


def verify_office_email(email: str, website: str) -> str:
    """Reuses the same verified contact email as Office Email when the
    email's domain matches the verified company website (i.e. it is
    genuinely that office's email, not a generic freemail address)."""
    if not email or not website:
        return ""
    domain = website.removeprefix("https://").removeprefix("http://").removeprefix("www.").split("/")[0]
    email_domain = email.split("@")[-1].lower()
    if domain and (email_domain == domain.lower() or email_domain.endswith("." + domain.lower())):
        return email
    return ""
