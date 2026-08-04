"""Multi-source discovery (src/vendor_intel/pipeline/multi_source_discovery.py)
— topic-driven B2B directory/LinkedIn/Owler search templates, generalized
from discover_more_companies.py's hardcoded per-market query lists.
"""
from __future__ import annotations

import asyncio

from vendor_intel.clients.search_router import SearchResult
from vendor_intel.pipeline.multi_source_discovery import (
    _SOURCE_TEMPLATES,
    discover_additional_sources,
)


class _FakeRouter:
    def __init__(self, results_by_query: dict[str, list[SearchResult]]):
        self.results_by_query = results_by_query
        self.calls: list[str] = []

    async def search(self, query, **kw):
        self.calls.append(query)
        return self.results_by_query.get(query, [])


def test_empty_topic_returns_no_candidates():
    router = _FakeRouter({})
    result = asyncio.run(
        discover_additional_sources({}, None, settings=object(), router=router)
    )
    assert result == []
    assert router.calls == []


def test_extracts_candidates_from_search_results():
    topic = "rupture disc"
    alibaba_query = _SOURCE_TEMPLATES[0][1].format(topic=topic)
    router = _FakeRouter(
        {
            alibaba_query: [
                SearchResult(
                    title="Acme Rupture Discs | Official Site",
                    link="https://acme.com/products",
                    snippet="x",
                    backend="ddgs",
                ),
            ],
        }
    )
    result = asyncio.run(
        discover_additional_sources(
            {"industry": topic, "discovery_mode": "landscape"},
            None,
            settings=object(),
            router=router,
            max_sources=1,
        )
    )
    assert len(result) == 1
    assert result[0]["name"] == "Acme Rupture Discs"
    assert result[0]["domain"] == "acme.com"
    assert result[0]["discovery_source"] == "multi_source:alibaba"


def test_filters_out_junk_and_blocked_domains():
    topic = "rupture disc"
    query = _SOURCE_TEMPLATES[0][1].format(topic=topic)
    router = _FakeRouter(
        {
            query: [
                SearchResult(title="Real Co", link="https://real.com", snippet="", backend="ddgs"),
                SearchResult(
                    title="Directory Listing", link="https://crunchbase.com/x", snippet="", backend="ddgs"
                ),
                SearchResult(
                    title="Video", link="https://www.youtube.com/watch?v=1", snippet="", backend="ddgs"
                ),
                SearchResult(
                    title="Top 10 Rupture Disc Manufacturers 2024",
                    link="https://blog.example.com/list",
                    snippet="",
                    backend="ddgs",
                ),
            ],
        }
    )
    result = asyncio.run(
        discover_additional_sources(
            {"industry": topic, "discovery_mode": "landscape"},
            None,
            settings=object(),
            router=router,
            max_sources=1,
        )
    )
    assert [c["domain"] for c in result] == ["real.com"]


def test_dedupes_across_sources_by_domain():
    topic = "rupture disc"
    q1 = _SOURCE_TEMPLATES[0][1].format(topic=topic)
    q2 = _SOURCE_TEMPLATES[1][1].format(topic=topic)
    same_result = SearchResult(title="Acme Co", link="https://acme.com", snippet="", backend="ddgs")
    router = _FakeRouter({q1: [same_result], q2: [same_result]})
    result = asyncio.run(
        discover_additional_sources(
            {"industry": topic, "discovery_mode": "landscape"},
            None,
            settings=object(),
            router=router,
            max_sources=2,
        )
    )
    assert len(result) == 1


def test_max_sources_limits_query_count():
    topic = "rupture disc"
    router = _FakeRouter({})
    asyncio.run(
        discover_additional_sources(
            {"industry": topic, "discovery_mode": "landscape"},
            None,
            settings=object(),
            router=router,
            max_sources=3,
        )
    )
    assert len(router.calls) == 3


def test_one_source_failing_does_not_abort_others():
    topic = "rupture disc"
    q1 = _SOURCE_TEMPLATES[0][1].format(topic=topic)
    q2 = _SOURCE_TEMPLATES[1][1].format(topic=topic)

    class _FlakyRouter(_FakeRouter):
        async def search(self, query, **kw):
            self.calls.append(query)
            if query == q1:
                raise RuntimeError("network exploded")
            return self.results_by_query.get(query, [])

    router = _FlakyRouter(
        {q2: [SearchResult(title="Beta Co", link="https://beta.com", snippet="", backend="ddgs")]}
    )
    result = asyncio.run(
        discover_additional_sources(
            {"industry": topic, "discovery_mode": "landscape"},
            None,
            settings=object(),
            router=router,
            max_sources=2,
        )
    )
    assert len(result) == 1
    assert result[0]["domain"] == "beta.com"


def test_topic_prefers_scope_market_over_query_context_industry():
    topic_query_side = "wrong topic"
    scope = {"market": "rupture disc"}
    q1 = _SOURCE_TEMPLATES[0][1].format(topic="rupture disc")
    router = _FakeRouter({})
    asyncio.run(
        discover_additional_sources(
            {"industry": topic_query_side, "discovery_mode": "landscape"},
            scope,
            settings=object(),
            router=router,
            max_sources=1,
        )
    )
    assert router.calls == [q1]


def test_non_cyber_channel_mode_skips_mssp_templates():
    router = _FakeRouter({})
    asyncio.run(
        discover_additional_sources(
            {"industry": "avocado oil"},
            {"market": "avocado oil"},
            settings=object(),
            router=router,
            max_sources=12,
        )
    )
    joined = " | ".join(router.calls).lower()
    assert "mssp" not in joined
    assert "security integrator" not in joined
    assert "avocado oil" in joined


def test_cyber_channel_mode_includes_mssp_templates():
    router = _FakeRouter({})
    asyncio.run(
        discover_additional_sources(
            {"industry": "firewall"},
            {"market": "firewall"},
            settings=object(),
            router=router,
            max_sources=12,
        )
    )
    joined = " | ".join(router.calls).lower()
    assert "mssp" in joined or "security integrator" in joined
