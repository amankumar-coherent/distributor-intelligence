"""KOMPASS_PRODUCT_ID lets an operator supply the category id directly.

Automatic discovery legitimately fails for analyst-style market names:
Kompass indexes by product classification, so "Semiconductor Market" has no
page even though "Semiconductors NES" (id 40210) exists and lists 81
companies on page 1. Rather than lose those markets, the id can be passed in.

The safety rule: this skips DISCOVERY, never VERIFICATION. A supplied id is
still fetched and confirmed, so a wrong one is rejected exactly like an
AI-invented URL.
"""
from __future__ import annotations

import pytest

from vendor_intel.discovery import kompass_distributors as kd


class FakePage:
    def __init__(self, alive, final_url="", error=""):
        self.alive = alive
        self.final_url = final_url
        self.error = error
        self.html = ""


@pytest.fixture
def no_ai(monkeypatch):
    """AI discovery must not be consulted when a valid id is supplied."""
    calls = {"n": 0}

    def fake_discover(market):
        calls["n"] += 1
        return None

    monkeypatch.setattr(kd, "discover_product_page", fake_discover)
    return calls


def test_supplied_id_is_used_when_it_verifies(monkeypatch, no_ai):
    monkeypatch.setenv("KOMPASS_PRODUCT_ID", "40210")
    monkeypatch.setattr(
        kd.kompass_client, "fetch_kompass_page",
        lambda url: FakePage(True, "https://www.kompass.com/z/ww/a/semiconductors-nes/40210/"),
    )
    url, pid = kd.find_market_product_page("Semiconductor Market")
    assert pid == "40210"
    assert "semiconductors-nes" in url
    assert no_ai["n"] == 0, "AI discovery should not run when the id verifies"


def test_id_is_re_extracted_from_the_fetched_page(monkeypatch, no_ai):
    """The id recorded is the one the real page resolves to, not the string
    that was passed in."""
    monkeypatch.setenv("KOMPASS_PRODUCT_ID", "0746005")
    monkeypatch.setattr(
        kd.kompass_client, "fetch_kompass_page",
        lambda url: FakePage(True, "https://www.kompass.com/z/ww/a/natural-gas/0746005/"),
    )
    _, pid = kd.find_market_product_page("LNG Market")
    assert pid == "0746005"


def test_unreachable_id_falls_back_to_discovery(monkeypatch, no_ai):
    """A bad id must not be trusted just because a human typed it."""
    monkeypatch.setenv("KOMPASS_PRODUCT_ID", "99999999")
    monkeypatch.setattr(
        kd.kompass_client, "fetch_kompass_page",
        lambda url: FakePage(False, "", "404"),
    )
    assert kd.find_market_product_page("Nonsense Market") is None
    assert no_ai["n"] == 1, "must fall back to AI discovery when the id fails"


def test_unset_env_uses_normal_discovery(monkeypatch, no_ai):
    monkeypatch.delenv("KOMPASS_PRODUCT_ID", raising=False)
    called = {"fetch": 0}
    monkeypatch.setattr(
        kd.kompass_client, "fetch_kompass_page",
        lambda url: called.__setitem__("fetch", called["fetch"] + 1) or FakePage(True, url),
    )
    assert kd.find_market_product_page("Wearable Medical Devices") is None
    assert no_ai["n"] == 1
    assert called["fetch"] == 0, "must not fetch anything when no id is supplied"


def test_blank_env_is_ignored(monkeypatch, no_ai):
    monkeypatch.setenv("KOMPASS_PRODUCT_ID", "   ")
    assert kd.find_market_product_page("Some Market") is None
    assert no_ai["n"] == 1
