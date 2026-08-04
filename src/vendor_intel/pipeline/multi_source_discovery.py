"""Multi-source company discovery — B2B trade directories + LinkedIn/Owler.

Generalizes discover_more_companies.py's hand-written per-market query lists
(32 queries hardcoded for each of 6 specific markets — couldn't work for any
other market) into topic-driven templates that work for any market, using
the query plan's own `scope`/`query_context` instead of a hardcoded dict.

Candidates get NO special treatment: they're returned in the same
{name, domain, discovery_source} shape as phase2.discovery_fast's output and
flow through the exact same entity-gate/dedup/off-market-filter/enrichment/
classification gates as every other candidate. This is deliberate — the
original script's own GPT "is this a real company?" pre-validation step is
dropped here because the real gates already do this more rigorously (see the
batch-scaling plan's Phase 7 design).
"""
from __future__ import annotations

import re
from typing import Any

# (source label, query template) — {topic} is filled from the market's own
# scope/query text. Covers the sources discover_more_companies.py targeted
# (Alibaba, ThomasNet, Kompass, TradeIndia, ImportYeti, LinkedIn, Owler),
# minus the redundant geo-variant queries that script hardcoded per market.
# Channel-partner discovery templates (default CHANNEL_PARTNER_MODE).
# Keep role language market-agnostic — do not hardcode firewall/MSSP terms;
# those belong in queries/discovery/global_firewall_market.txt only.
_CHANNEL_SOURCE_TEMPLATES: list[tuple[str, str]] = [
    ("kompass", "site:kompass.com {topic} distributor"),
    ("thomasnet", "site:thomasnet.com {topic} distributor"),
    ("tradeindia", "site:tradeindia.com {topic} distributor"),
    ("linkedin_dist", "{topic} distributor OR reseller OR wholesaler site:linkedin.com/company"),
    ("linkedin_partner", "{topic} channel partner OR distributor site:linkedin.com/company"),
    ("owler", "{topic} distributor OR channel partner site:owler.com/company"),
    ("partner", "{topic} authorized distributor OR wholesale supplier"),
    ("reseller", "{topic} reseller OR importer OR exporter"),
]

# Extra channel templates only for cyber / network-security markets.
_CYBER_CHANNEL_SOURCE_TEMPLATES: list[tuple[str, str]] = [
    ("linkedin_si", "{topic} system integrator OR MSSP site:linkedin.com/company"),
    ("integrator", "{topic} system integrator OR security integrator"),
    ("var", "{topic} authorized distributor OR value added reseller"),
]

_CYBER_TOPIC_RE = re.compile(
    r"\b(?:firewall|cyber|cybersecurity|infosec|endpoint|ngfw|network\s+security|"
    r"zero[\s-]?trust|xdr|siem|mssp)\b",
    re.I,
)

# Landscape / manufacturer census templates (opt-out mode only).
_SOURCE_TEMPLATES: list[tuple[str, str]] = [
    ("alibaba", "site:alibaba.com {topic} manufacturer"),
    ("globalsources", "site:globalsources.com {topic} manufacturer"),
    ("kompass", "site:kompass.com {topic} manufacturer"),
    ("thomasnet", "site:thomasnet.com {topic} supplier"),
    ("tradeindia", "site:tradeindia.com {topic} manufacturer"),
    ("importyeti", "site:importyeti.com {topic}"),
    ("linkedin", "{topic} company site:linkedin.com/company"),
    ("owler", "{topic} company site:owler.com/company"),
]


def _discovery_topic(query_context: dict[str, Any], scope: dict[str, Any] | None) -> str:
    scope = scope or {}
    return str(scope.get("market") or query_context.get("industry") or "").strip()


async def discover_additional_sources(
    query_context: dict[str, Any],
    scope: dict[str, Any] | None,
    settings: Any,
    *,
    router: Any = None,
    geo: str = "",
    max_sources: int = 8,
) -> list[dict[str, str]]:
    """Run B2B-directory / LinkedIn / Owler searches for this market.

    Returns raw {name, domain, discovery_source} candidates, deduped by
    domain within this call — unvalidated. The caller's entity-gate and
    classification stages do the real filtering, same as any other
    discovery source.
    """
    topic = _discovery_topic(query_context, scope)
    if not topic:
        return []

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.discovery.entity_extract import (
        clean_title_as_name,
        is_blocked_domain,
        is_junk_url,
    )
    from vendor_intel.utils.domains import domain_from_url

    router = router or FreeSearchRouter(settings)
    candidates: dict[str, dict[str, str]] = {}

    from vendor_intel.funnel.channel_partner_types import channel_partner_mode_enabled

    if channel_partner_mode_enabled(query_context, settings=settings):
        templates = list(_CHANNEL_SOURCE_TEMPLATES)
        if _CYBER_TOPIC_RE.search(topic):
            templates.extend(_CYBER_CHANNEL_SOURCE_TEMPLATES)
    else:
        templates = list(_SOURCE_TEMPLATES)

    # Analyst-supplied discovery queries (from --queries-file) run in addition to
    # the built-in templates. Full phrases run as-is; bare role terms get the topic
    # appended so the search stays market-scoped.
    extra_queries = [
        str(q).strip()
        for q in (query_context.get("extra_discovery_queries") or [])
        if str(q).strip()
    ]
    runnable: list[tuple[str, str]] = [
        (source_name, template.format(topic=topic))
        for source_name, template in templates[: max(0, max_sources)]
    ]
    for q in extra_queries:
        full = q if topic.lower() in q.lower() else f"{q} {topic}"
        runnable.append(("custom", full))

    for source_name, query in runnable:
        try:
            results = await router.search(
                query, geo=geo, discovery_mode=True, search_topic=topic,
            )
        except Exception as exc:
            print(f"  [multi_source] {source_name} search failed: {exc}", flush=True)
            continue
        for r in results:
            if is_junk_url(r.link):
                continue
            dom = domain_from_url(r.link)
            if not dom or is_blocked_domain(dom):
                continue
            name = clean_title_as_name(r.title)
            if not name or dom in candidates:
                continue
            candidates[dom] = {
                "name": name,
                "domain": dom,
                "discovery_source": f"multi_source:{source_name}",
            }

    if candidates:
        print(
            f"  [multi_source] {len(candidates)} candidates from "
            f"{min(max_sources, len(_SOURCE_TEMPLATES))} directory/social sources",
            flush=True,
        )
    return list(candidates.values())
