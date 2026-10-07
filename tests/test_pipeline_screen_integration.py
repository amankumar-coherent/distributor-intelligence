"""The main pipeline's screening path uses the market gate and block guard.

These fixes live in _verdict_from_overview, which the pipeline calls via
screen_roles_via_google_ai. This pins that wiring so the pipeline can never
silently fall back to the old "is it a distributor of anything?" behaviour.
"""
from __future__ import annotations

import pytest

from vendor_intel.discovery import chatgpt_company_screen as cs

MARKET = "Wearable Medical Devices"


def fake_backend(monkeypatch, answers: dict[str, str]):
    """Stub google_ai_ask so no network/browser is touched."""
    async def fake_ask(query, **kw):
        for name, text in answers.items():
            if name.lower() in query.lower():
                return {"markdown": text}
        return {"markdown": ""}

    monkeypatch.setattr(
        "vendor_intel.clients.google_ai_scraper.google_ai_ask", fake_ask
    )


def test_wrong_market_company_is_not_accepted(monkeypatch):
    """Fritz Minke: a real distributor, of the wrong industry."""
    fake_backend(monkeypatch, {
        "Minke": "No, Fritz Minke is not a distributor of wearable medical devices. "
                 "They are a B2B wholesaler for orthopedic shoemaking.",
        "Acme": "Acme is a distributor of wearable medical devices in Italy.",
    })
    companies = [cs.ScreenedCompany(name="Fritz Minke"), cs.ScreenedCompany(name="Acme")]
    cs.screen_roles_via_google_ai(companies, market=MARKET, concurrency=1)

    by_name = {c.name: c for c in companies}
    assert by_name["Fritz Minke"].is_distributor is False
    assert by_name["Acme"].is_distributor is True


def test_bot_block_page_never_accepts_a_company(monkeypatch):
    """A CAPTCHA page echoes the query, which contains the word
    'distributor'. Parsed as prose it accepted every blocked company."""
    fake_backend(monkeypatch, {
        "Acme": "Our systems have detected unusual traffic from your computer network. "
                "Why did this happen? .../search?q=Is%20Acme%20a%20distributor",
    })
    companies = [cs.ScreenedCompany(name="Acme")]
    cs.screen_roles_via_google_ai(companies, market=MARKET, concurrency=1)

    assert companies[0].verdict == "unknown"
    assert companies[0].is_distributor is False


def test_genuine_distributor_still_accepted(monkeypatch):
    fake_backend(monkeypatch, {
        "Acme": "Acme S.p.A. is a distributor and reseller of wearable medical devices.",
    })
    companies = [cs.ScreenedCompany(name="Acme")]
    cs.screen_roles_via_google_ai(companies, market=MARKET, concurrency=1)
    assert companies[0].is_distributor is True


def test_manufacturer_still_rejected(monkeypatch):
    fake_backend(monkeypatch, {
        "Acme": "Acme manufactures wearable medical devices at its own plant.",
    })
    companies = [cs.ScreenedCompany(name="Acme")]
    cs.screen_roles_via_google_ai(companies, market=MARKET, concurrency=1)
    assert companies[0].is_distributor is False
    assert companies[0].verdict == "manufacturer"
