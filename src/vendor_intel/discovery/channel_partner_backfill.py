"""Web backfill: keep finding distributors/channel partners until export min is met.

Used when CHANNEL_PARTNER_MODE is on and filter_for_export yields fewer than
PIPELINE_GLOBAL_EXPORT_MIN_ROWS (default 300). Searches the open web (SearXNG /
DDGS / Google AI via the search router) with distributor-only prompts and appends
vetted channel-partner rows. No practical upper cap (caller enforces export_max).
"""
from __future__ import annotations

import os
import re
from typing import Any
from urllib.parse import quote_plus

from vendor_intel.discovery.country_sweep import PRIORITY_COUNTRIES, world_countries
from vendor_intel.discovery.entity_extract import (
    clean_title_as_name,
    is_blocked_domain,
    is_junk_url,
    is_valid_company,
)
from vendor_intel.funnel.channel_partner_types import (
    channel_partner_mode_enabled,
    is_accepted_channel_type,
)
from vendor_intel.utils.domains import domain_from_url


def backfill_enabled() -> bool:
    raw = (os.getenv("CHANNEL_PARTNER_BACKFILL") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def backfill_min_rows(default: int = 300) -> int:
    try:
        return max(0, int(os.getenv("CHANNEL_PARTNER_BACKFILL_MIN") or default or 300))
    except ValueError:
        return default


def _market_label(market: str) -> str:
    m = (market or "").strip()
    if m.lower().endswith(" market"):
        m = m[: -len(" market")].strip()
    return m or "market"


def _name_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


_JUNK_TITLE = re.compile(
    r"(?:market\s+(?:size|share|report|outlook|forecast)|list\s+of|top\s+\d+|"
    r"imarc|mordor|grand\s+view|wikipedia|buyers?\s+guide)",
    re.I,
)


def build_backfill_queries(market: str, *, start: int = 0, batch: int = 40) -> list[str]:
    """Distributor-only web queries (no manufacturer wording)."""
    m = _market_label(market)
    templates: list[str] = [
        f"{m} distributors wholesalers worldwide company list",
        f"{m} authorized distributors channel partners companies",
        f"{m} importers distributors B2B wholesale companies",
        f"{m} bulk oil distributors foodservice suppliers",
        f"{m} ingredient distributors wholesalers companies",
        f"site:thomasnet.com {m} distributor",
        f"site:kompass.com {m} distributor",
        f"site:europages.com {m} distributor",
        f"site:tradeindia.com {m} distributor wholesale",
    ]
    countries = list(PRIORITY_COUNTRIES) + [
        c for c in world_countries(max_countries=120) if c not in PRIORITY_COUNTRIES
    ]
    for c in countries:
        templates.append(f"{m} distributors wholesalers importers in {c}")
        templates.append(f"{m} authorized distributor companies {c}")

    # Rotate through the list so successive backfill rounds get fresh queries.
    if start >= len(templates):
        start = start % len(templates)
    out = templates[start : start + batch]
    if len(out) < batch:
        out.extend(templates[: batch - len(out)])
    return out


def _row_from_hit(
    *,
    name: str,
    domain: str,
    link: str,
    snippet: str,
    market: str,
    query: str,
) -> dict[str, Any]:
    m = _market_label(market)
    return {
        "company": name,
        "brand": name,
        "domain": domain,
        "website": link if link.startswith("http") else f"https://{domain}/",
        "role": "Distributor",
        "company_function": "distributor",
        "discovered_via_function": "distributor",
        "is_relevant": True,
        "confidence": 0.72,
        "quality_score": 0.72,
        "market_relevance_score": 2,
        "company_summary": (
            f"Web backfill channel partner for {m}. "
            f"Found via distributor search: {query[:120]}. "
            f"{(snippet or '')[:240]}"
        ),
        "summary": f"{name} — {m} distributor / channel partner (web backfill).",
        "signals": {
            "is_distributor": True,
            "industry_match": True,
            "channel_backfill": True,
        },
        "discovery_source": "channel_partner_backfill",
        "enrichment_source": "web_backfill",
        "data_sources": ["web_search_backfill"],
        "value_chain_section": "Distributors",
        "is_seed": False,
        "is_analyst_seed": False,
        "_backfill": True,
    }


async def backfill_channel_partners_until_min(
    *,
    market: str,
    query_context: dict[str, Any],
    settings: Any,
    final_rows: list[dict[str, Any]],
    rejected_rows: list[dict[str, Any]] | None = None,
    min_rows: int | None = None,
    router: Any = None,
    scope: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Search the web and return NEW channel-partner rows until min_rows is reached.

    Does not mutate ``final_rows``; caller extends and dedupes.
    """
    if not backfill_enabled():
        return []
    if not channel_partner_mode_enabled(query_context, settings=settings):
        return []

    target = min_rows if min_rows is not None else backfill_min_rows(
        int(getattr(settings, "pipeline_global_export_min_rows", 300) or 300)
    )
    if target <= 0:
        return []

    have = len(final_rows or [])
    if have >= target:
        return []

    from vendor_intel.clients.search_router import FreeSearchRouter

    router = router or FreeSearchRouter(settings)
    m = _market_label(market or str((scope or {}).get("market") or query_context.get("industry") or ""))

    seen_dom: set[str] = set()
    seen_name: set[str] = set()
    for r in list(final_rows or []) + list(rejected_rows or []):
        d = str(r.get("domain") or "").strip().lower()
        n = _name_key(str(r.get("company") or r.get("brand") or ""))
        if d:
            seen_dom.add(d)
        if n:
            seen_name.add(n)

    need = target - have
    print(
        f"  [channel-backfill] Need {need} more genuine distributors "
        f"(have {have}, target {target}) — searching web…",
        flush=True,
    )

    added: list[dict[str, Any]] = []
    cursor = 0
    round_i = 0
    stagnant = 0
    max_rounds = int(os.getenv("CHANNEL_PARTNER_BACKFILL_MAX_ROUNDS", "80") or "80")
    batch_q = int(os.getenv("CHANNEL_PARTNER_BACKFILL_QUERIES_PER_ROUND", "12") or "12")

    while len(added) + have < target and round_i < max_rounds:
        round_i += 1
        queries = build_backfill_queries(m, start=cursor, batch=batch_q)
        cursor += batch_q
        before = len(added)

        for q in queries:
            if len(added) + have >= target:
                break
            try:
                results = await router.search(
                    q, geo="global", discovery_mode=True, search_topic=m
                )
            except Exception as exc:
                print(f"  [channel-backfill] search failed: {exc}", flush=True)
                continue

            for hit in results or []:
                if len(added) + have >= target:
                    break
                link = str(getattr(hit, "link", None) or getattr(hit, "url", None) or "")
                title = str(getattr(hit, "title", None) or "")
                snippet = str(getattr(hit, "snippet", None) or "")
                if not link or is_junk_url(link):
                    continue
                if _JUNK_TITLE.search(title):
                    continue
                dom = domain_from_url(link)
                if not dom or is_blocked_domain(dom) or dom in seen_dom:
                    continue
                name = clean_title_as_name(title) or ""
                if not name or _JUNK_TITLE.search(name):
                    continue
                nk = _name_key(name)
                if not nk or nk in seen_name:
                    continue
                if not is_valid_company(name, dom, scope=scope or {}):
                    continue
                # Must look channel-ish from query context (all our queries are distributor).
                role = "Distributor"
                if not is_accepted_channel_type(role):
                    continue

                row = _row_from_hit(
                    name=name,
                    domain=dom,
                    link=link,
                    snippet=snippet,
                    market=m,
                    query=q,
                )
                seen_dom.add(dom)
                seen_name.add(nk)
                added.append(row)

        gained = len(added) - before
        print(
            f"  [channel-backfill] round {round_i}/{max_rounds}: +{gained} "
            f"(total new {len(added)}, combined {have + len(added)}/{target})",
            flush=True,
        )
        if gained == 0:
            stagnant += 1
        else:
            stagnant = 0
        if stagnant >= 8:
            print(
                "  [channel-backfill] No new domains for several rounds — stopping",
                flush=True,
            )
            break

    print(
        f"  [channel-backfill] Done: added {len(added)} distributor rows "
        f"(will combine to {have + len(added)})",
        flush=True,
    )
    return added
