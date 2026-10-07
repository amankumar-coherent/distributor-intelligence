"""Candidate contact/leadership paths are fetched concurrently.

Most of these paths are guesses that 404. Fetching them one at a time
meant a company's whole gap-fill budget was spent waiting on pages that do
not exist — a dead domain burned ~17 sequential timeouts, so contact_page
blew its 45s cap and returned nothing. These tests pin the speed fix AND
the selection semantics it had to preserve.
"""
from __future__ import annotations

import asyncio

import pytest

import vendor_intel.enrichment.gap_fill.contact_page as cp


@pytest.fixture
def pages(monkeypatch):
    store: dict[str, tuple[str, str]] = {}

    async def fake(url):
        await asyncio.sleep(0.01)
        return store.get(url, ("", ""))

    monkeypatch.setattr(cp, "_fetch_page_text", fake)
    return store


def test_earlier_path_with_contact_details_wins(pages):
    """/contact must beat /contact-us even when /contact-us is far longer —
    concurrency must not turn priority into a race."""
    pages["https://x.com/contact"] = ("mail us at a@x.com", "https://x.com/contact")
    pages["https://x.com/contact-us"] = ("y" * 9000 + " b@x.com", "https://x.com/contact-us")
    _, url = asyncio.run(cp._find_contact_url("x.com"))
    assert url == "https://x.com/contact"


def test_falls_back_to_the_longest_page_when_no_details_found(pages):
    pages["https://x.com/contact"] = ("short", "https://x.com/contact")
    pages["https://x.com/en/contact"] = ("much longer body here", "https://x.com/en/contact")
    _, url = asyncio.run(cp._find_contact_url("x.com"))
    assert url == "https://x.com/en/contact"


def test_homepage_fallback_when_every_path_is_dead(pages):
    pages["https://x.com/"] = ("home a@x.com", "https://x.com/")
    text, _ = asyncio.run(cp._find_contact_url("x.com"))
    assert "a@x.com" in text


def test_leadership_prefers_a_page_naming_a_leader(pages):
    pages["https://x.com/about"] = ("nothing here", "https://x.com/about")
    pages["https://x.com/team"] = ("Jane Doe, Managing Director", "https://x.com/team")
    _, url = asyncio.run(cp._find_leadership_text("x.com"))
    assert url == "https://x.com/team"


def test_one_failing_fetch_does_not_abort_the_others(monkeypatch):
    """gather(return_exceptions=True): a DNS failure on one guessed path
    must not lose the good result from another."""
    async def boom(url):
        if "contact-us" in url:
            raise RuntimeError("dns")
        if url.endswith("/contactus"):
            return ("c@x.com", "https://x.com/contactus")
        return ("", "")

    monkeypatch.setattr(cp, "_fetch_page_text", boom)
    _, url = asyncio.run(cp._find_contact_url("x.com"))
    assert url == "https://x.com/contactus"


def test_paths_are_fetched_concurrently_not_serially(monkeypatch):
    async def slow(url):
        await asyncio.sleep(0.3)
        return ("", "")

    monkeypatch.setattr(cp, "_fetch_page_text", slow)

    async def timed():
        loop = asyncio.get_event_loop()
        start = loop.time()
        await cp._find_leadership_text("x.com")
        return loop.time() - start

    elapsed = asyncio.run(timed())
    # 9 paths x 0.3s would be 2.7s if serial.
    assert elapsed < 1.5, f"still sequential: {elapsed:.2f}s"
