"""Search stability — engine order, cache, rate limits."""
from __future__ import annotations

from vendor_intel.clients.ddgs_engines import (
    normalize_text_backends,
    preferred_text_backend_order,
)
from vendor_intel.clients.search_cache_disk import (
    get_disk_cached_search,
    put_disk_cached_search,
)
from vendor_intel.clients.search_stats import reset_for_tests, snapshot


def test_preferred_backend_order_google_last():
    names = ["google", "bing", "brave", "duckduckgo", "yandex"]
    ordered = preferred_text_backend_order(names)
    assert ordered[0] == "bing"
    assert ordered[-1] == "google"
    assert "duckduckgo" in ordered
    assert ordered.index("brave") < ordered.index("google")


def test_normalize_text_backends_default_order():
    param, names, _ = normalize_text_backends("")
    assert names[0] == "bing"
    assert names[-1] == "google"
    assert "google" in param


def test_disk_search_cache_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setenv("SEARCH_DISK_CACHE_DIR", str(tmp_path))
    monkeypatch.setenv("SEARCH_DISK_CACHE", "true")
    put_disk_cached_search(
        "acme corp official site",
        "global",
        "router:test",
        [{"title": "Acme", "link": "https://acme.com", "snippet": "x", "backend": "ddgs_bing"}],
    )
    rows = get_disk_cached_search("acme corp official site", "global", "router:test")
    assert rows and rows[0]["link"] == "https://acme.com"


def test_search_stats_counters():
    reset_for_tests()
    from vendor_intel.clients import search_stats

    search_stats.record_search_start()
    search_stats.record_cache_hit()
    search_stats.record_rate_limit_skip("ddgs:google")
    snap = snapshot()
    assert snap["total_searches"] == 1
    assert snap["cache_hits"] == 1
    assert snap["rate_limit_skips"] == 1
