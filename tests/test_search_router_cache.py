"""Search-result caching in src/vendor_intel/clients/search_router.py (Phase 6)
— the router aggregates several engines into one filtered result per
(query, geo, mode) and caches THAT, so a repeated/near-duplicate query
(common across a large batch's regional-variant markets) skips all real
network calls on the second ask.
"""
from __future__ import annotations

import asyncio

import pytest

from vendor_intel.batch import db
from vendor_intel.clients import search_router as sr
from vendor_intel.config import Settings


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    monkeypatch.setattr(db, "_DB_PATH", tmp_path / "batch.db")
    monkeypatch.setattr(db, "_CACHE_DIR", tmp_path / "company_cache")
    db.reset_for_tests()
    yield
    db.reset_for_tests()


def test_cache_namespace_is_mode_specific():
    assert sr._cache_namespace(discovery_mode=False, validation_mode=False) == "router:default"
    assert sr._cache_namespace(discovery_mode=True, validation_mode=False) == "router:discovery"
    assert sr._cache_namespace(discovery_mode=False, validation_mode=True) == "router:validation"
    # validation_mode wins if (incorrectly) both are set — matches search()'s own precedence
    assert sr._cache_namespace(discovery_mode=True, validation_mode=True) == "router:validation"


def test_cached_search_results_roundtrip():
    assert sr._cached_search_results("widget market", "Brazil", "router:default") is None

    results = [sr.SearchResult(title="Acme", link="https://acme.com", snippet="x", backend="ddgs")]
    sr._put_cached_search_results("widget market", "Brazil", "router:default", results)

    got = sr._cached_search_results("widget market", "Brazil", "router:default")
    assert got == results


def test_cached_search_results_scoped_by_namespace():
    results = [sr.SearchResult(title="A", link="https://a.com", snippet="", backend="ddgs")]
    sr._put_cached_search_results("widget market", "Brazil", "router:discovery", results)
    assert sr._cached_search_results("widget market", "Brazil", "router:default") is None
    assert sr._cached_search_results("widget market", "Brazil", "router:discovery") == results


@pytest.fixture
def router(monkeypatch):
    """A FreeSearchRouter with SearXNG/host-probing short-circuited to False
    (no real sockets) and ddgs as the only live path, for controlled testing."""
    monkeypatch.setattr(sr, "searxng_local_reachable", lambda: False)
    monkeypatch.setattr(sr, "searxng_search_usable", lambda: False)
    monkeypatch.setattr(sr, "print_connectivity_report", lambda: None)  # avoid real TCP probes
    settings = Settings.load().model_copy(
        update={"use_mock_data": False, "mock_mode": False, "results_per_query": 5}
    )
    return sr.FreeSearchRouter(settings)


def _mock_all_engines(monkeypatch, calls: dict) -> None:
    """Mock every real network path search() can fall back to, so the test
    only measures cache behavior, not which specific fallback fired."""
    from vendor_intel.clients.duckduckgo import DuckResult

    async def fake_duckduckgo_search(query, max_results=15, *, geo=""):
        calls["n"] += 1
        return [
            DuckResult(title="Acme", link="https://acme.com", snippet="x", engine="ddgs"),
            DuckResult(title="Beta", link="https://beta.com", snippet="x", engine="ddgs"),
            DuckResult(title="Gamma", link="https://gamma.com", snippet="x", engine="ddgs"),
            DuckResult(title="Delta", link="https://delta.com", snippet="x", engine="ddgs"),
            DuckResult(title="Epsilon", link="https://epsilon.com", snippet="x", engine="ddgs"),
        ]

    def fake_search_bing_html_try(query, max_results=12, *, geo="", force=False):
        calls["n"] += 1
        return []

    async def fake_wikipedia_search(query, max_results=15, *, geo=""):
        return []

    monkeypatch.setattr(sr, "duckduckgo_available", lambda: True)
    monkeypatch.setattr(sr, "duckduckgo_search", fake_duckduckgo_search)
    monkeypatch.setattr(
        "vendor_intel.clients.bing_html.search_bing_html_try", fake_search_bing_html_try
    )
    monkeypatch.setattr(sr, "wikipedia_search", fake_wikipedia_search)
    monkeypatch.setattr(sr, "is_mock_run", lambda settings: False)


def test_search_second_identical_call_hits_cache_not_network(monkeypatch, router):
    calls = {"n": 0}
    _mock_all_engines(monkeypatch, calls)

    first = asyncio.run(router.search("rupture disc manufacturers", geo="Brazil"))
    calls_after_first = calls["n"]
    assert calls_after_first >= 1
    assert len(first) >= 1

    second = asyncio.run(router.search("rupture disc manufacturers", geo="Brazil"))
    assert calls["n"] == calls_after_first, (
        "second identical search should be served from cache, not the network"
    )
    assert [r.link for r in second] == [r.link for r in first]


def test_search_different_query_is_not_a_cache_hit(monkeypatch, router):
    calls = {"n": 0}
    _mock_all_engines(monkeypatch, calls)

    asyncio.run(router.search("rupture disc manufacturers", geo="Brazil"))
    calls_after_first = calls["n"]
    asyncio.run(router.search("a completely different query", geo="Brazil"))
    assert calls["n"] > calls_after_first


def test_search_mock_mode_bypasses_cache_entirely(monkeypatch, router):
    monkeypatch.setattr(sr, "is_mock_run", lambda settings: True)
    result = asyncio.run(router.search("anything", geo="Brazil"))
    assert result == []
    # Nothing should have been written to the cache in mock mode.
    assert db.search_cache_stats()["cached_queries"] == 0
