"""Session-wide test safety nets.

Tests must never depend on the real .env for whether external, real-world
side effects (like launching a live browser session) can happen — a test
that forgets to mock one stage of a pipeline should fail loudly on a bad
assertion, not silently succeed by actually calling out to chatgpt.com via
a real Selenium-controlled Chrome window. See the incident that motivated
this: two tests in test_company_enrichment.py mocked every enrich_company()
stage except verify_website_stage, which relied on CHATGPT_SCRAPER_ENABLED
being False; once that flag was legitimately turned on in .env for real
runs, those tests started launching a real browser.
"""
from __future__ import annotations

import pytest


@pytest.fixture(autouse=True)
def _chatgpt_scraper_disabled_by_default():
    """Force CHATGPT_SCRAPER_ENABLED off for every test, regardless of what
    the real .env says, so a missed mock can never launch a real browser
    session. Tests that specifically want to exercise the enabled path
    already do `patch(...CHATGPT_SCRAPER_ENABLED, True)` themselves, which
    still works fine layered on top of this default."""
    import vendor_intel.scraping.chatgpt_scraper as cs

    original = cs.CHATGPT_SCRAPER_ENABLED
    cs.CHATGPT_SCRAPER_ENABLED = False
    try:
        yield
    finally:
        cs.CHATGPT_SCRAPER_ENABLED = original


@pytest.fixture(autouse=True)
def _browser_backends_disabled_by_default(monkeypatch):
    """Force every browser-driving backend off for all tests.

    Same hazard as the ChatGPT scraper above, and it bit again: once
    GOOGLE_AI_MODE_ENABLED=true landed in the real .env, the suite began
    launching real Playwright Chrome windows and hung indefinitely instead
    of finishing in seconds. Tests must never depend on .env for whether a
    real browser or network call can happen — a test that forgets a mock
    should fail on an assertion, not silently drive a browser.

    Tests that specifically exercise an enabled path set these themselves
    via monkeypatch, which still layers correctly on top of this default.
    """
    monkeypatch.setenv("GOOGLE_AI_MODE_ENABLED", "false")
    monkeypatch.setenv("KOMPASS_GOOGLE_AI_SCREEN", "false")
    monkeypatch.setenv("KOMPASS_CHATGPT_SCREEN", "false")
    # Also the AI Overview relay: it talks to a local backend that drives a
    # real Chrome tab, so a test that reaches it hangs the suite waiting on
    # a live query rather than failing fast.
    monkeypatch.setenv("GOOGLE_AI_SCRAPER_ENABLED", "false")


@pytest.fixture(autouse=True)
def _never_write_to_the_real_output_dir(tmp_path_factory, monkeypatch):
    """Redirect the market-research export directory into a temp dir.

    Tests call export_xlsx() to exercise the real writer, and several patch
    STATE_DIR but not OUTPUT_DIR — so the export read temp state and wrote
    the workbook straight into output/market_research/, next to live market
    data.

    That is not merely untidy. A full test run left a one-row "Acme Inc"
    workbook where a 325-company Silicon Carbide export belonged, the
    verifier ran against it, and a 0-company verified file was written over
    a good one. The suite has to be unable to reach that directory at all,
    rather than each test remembering to patch it.
    """
    target = tmp_path_factory.mktemp("market_research_output")
    for module in (
        "vendor_intel.pipeline.market_research.export_xlsx",
    ):
        try:
            monkeypatch.setattr(f"{module}.OUTPUT_DIR", target, raising=False)
        except Exception:  # noqa: BLE001 - never block a test on this
            pass
    yield
