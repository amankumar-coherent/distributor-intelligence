"""Web search router — primary: ddgs library; optional fallbacks in search_router only."""
from __future__ import annotations

import asyncio
import os
from dataclasses import dataclass

from vendor_intel.clients.duckduckgo import (
    ddgs_backend_param,
    duckduckgo_available,
    duckduckgo_search,
    duckduckgo_search_many,
)
from vendor_intel.clients.host_reachability import (
    print_connectivity_report,
    searxng_local_reachable,
    searxng_search_usable,
)
from vendor_intel.clients.searxng import searxng_search_any, searxng_urls
from vendor_intel.clients.wikipedia_search import wikipedia_search
from vendor_intel.config import Settings
from vendor_intel.mock.fixtures import is_mock_run

try:
    from vendor_intel.clients.google_ai_scraper import (
        google_ai_discovery_enabled,
        google_ai_enabled,
        google_ai_search,
    )
except Exception:  # pragma: no cover
    def google_ai_enabled() -> bool:  # type: ignore[misc]
        return False

    def google_ai_discovery_enabled() -> bool:  # type: ignore[misc]
        return False

    async def google_ai_search(query: str, *, max_results: int = 12):  # type: ignore[misc]
        return []

_probe_printed = False


def _geo_suffix(geo: str) -> str:
    """Space-prefixed geography for appending to queries (never strip the leading space)."""
    g = (geo or "").strip()
    if not g or g.lower() == "global":
        return ""
    return f" {g}"


def _backup_hit_threshold(*, discovery_mode: bool, validation_mode: bool, default_min: int) -> int:
    if validation_mode:
        return 1
    if discovery_mode:
        try:
            return max(1, int(os.getenv("DISCOVERY_MIN_BEFORE_SEARXNG", "2") or "2"))
        except ValueError:
            return 2
    return default_min


@dataclass
class SearchResult:
    title: str
    link: str
    snippet: str
    backend: str


def _cache_namespace(*, discovery_mode: bool, validation_mode: bool) -> str:
    """search_cache's `engine` column doubles as a cache namespace here — the
    router aggregates several real engines into one filtered result per
    (query, geo, mode), so that's what gets cached, not a per-engine result."""
    if validation_mode:
        return "router:validation"
    if discovery_mode:
        return "router:discovery"
    return "router:default"


def _cached_search_results(query: str, geo: str, namespace: str) -> list["SearchResult"] | None:
    """Cross-market search-result cache — disk + SQLite batch cache."""
    try:
        from vendor_intel.clients.search_cache_disk import get_disk_cached_search

        disk = get_disk_cached_search(query, geo, namespace)
        if disk is not None:
            from vendor_intel.clients import search_stats

            search_stats.record_cache_hit()
            return [SearchResult(**row) for row in disk]
    except Exception:
        pass
    try:
        from vendor_intel.batch.db import get_cached_search

        cached = get_cached_search(query, geo, namespace)
    except Exception:
        return None
    if cached is None:
        return None
    try:
        from vendor_intel.clients import search_stats

        search_stats.record_cache_hit()
        return [SearchResult(**row) for row in cached]
    except (TypeError, KeyError):
        return None


def _put_cached_search_results(query: str, geo: str, namespace: str, results: list["SearchResult"]) -> None:
    payload = [
        {"title": r.title, "link": r.link, "snippet": r.snippet, "backend": r.backend}
        for r in results
    ]
    try:
        from vendor_intel.clients.search_cache_disk import put_disk_cached_search

        put_disk_cached_search(query, geo, namespace, payload)
    except Exception:
        pass
    try:
        from vendor_intel.batch.db import put_cached_search

        put_cached_search(query, geo, namespace, payload)
    except Exception:
        pass  # search cache is a performance optimization, never fail a search over it


def _query_variants(
    query: str,
    market: str,
    geo: str,
    *,
    validation_mode: bool = False,
    discovery_mode: bool = False,
) -> list[str]:
    base = query.strip()
    if validation_mode or discovery_mode:
        return [base] if base else []

    variants = [base]
    tail = _geo_suffix(geo)
    if market and market.lower() not in base.lower():
        variants.append(f"{market}{tail}".strip())
    variants.extend([f"list of {base}", f"top {base}"])
    seen: set[str] = set()
    out: list[str] = []
    for v in variants:
        key = v.lower()
        if key in seen or not v:
            continue
        seen.add(key)
        out.append(v)
    return out[:4]


def _log(msg: str) -> None:
    print(f"  [search] {msg}", flush=True)


def _env_true(name: str, default: bool = False) -> bool:
    raw = os.getenv(name, "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "on")


def skip_ddgs(*, discovery_mode: bool = False) -> bool:
    """Skip ddgs package (use router fallbacks only)."""
    if _env_true("SKIP_DDGS") or _env_true("SKIP_DUCKDUCKGO"):
        return True
    if discovery_mode and _env_true("DISCOVERY_SKIP_DDGS"):
        return True
    return False


def search_fallbacks_enabled() -> bool:
    """When false, only ddgs (+ Wikipedia thin-result helper) — no Bing HTML / SearXNG."""
    return _env_true("SEARCH_USE_FALLBACKS", default=True)


def search_stack_only() -> bool:
    """Strict Step-5 engines: DDGS + SearXNG + Google AI scraper (no Bing/Exa/Wikipedia)."""
    return _env_true("SEARCH_STACK_ONLY")


def search_stack_description(settings: Settings) -> str:
    try:
        from vendor_intel.integrations.agent_reach_client import exa_search_enabled
    except Exception:
        def exa_search_enabled() -> bool:  # type: ignore[misc]
            return False

    if skip_ddgs():
        parts = []
        if search_fallbacks_enabled():
            if not _env_true("SKIP_BING_HTML"):
                parts.append("Bing HTML")
            if settings.searxng_base_url:
                parts.append("SearXNG")
        return "SKIP_DDGS -> " + (" -> ".join(parts) if parts else "none")
    order = [f"ddgs (backend={ddgs_backend_param()})"]
    if search_fallbacks_enabled():
        if not _env_true("SKIP_BING_HTML"):
            order.append("Bing HTML")
        if settings.searxng_base_url:
            order.append("SearXNG")
    if exa_search_enabled():
        order.append("Agent Reach Exa (optional)")
    if google_ai_enabled():
        order.append("Google AI scraper (:15551, optional)")
    order.append("Wikipedia (thin results)")
    return " -> ".join(order)


class FreeSearchRouter:
    def __init__(self, settings: Settings):
        self._settings = settings
        self._num = settings.results_per_query
        self._min = settings.min_results_before_backup
        self._searx_urls = searxng_urls(settings.searxng_base_url)
        self._searxng_tcp = searxng_local_reachable() and bool(self._searx_urls)
        self._searxng_up = self._searxng_tcp and searxng_search_usable()
        prefer = os.getenv("PREFER_SEARXNG", "").strip().lower()
        self._prefer_searxng = prefer in ("1", "true", "yes") or (
            prefer not in ("0", "false", "no") and self._searxng_up
        )
        self._skip_ddgs = skip_ddgs(discovery_mode=False)

    async def _try_searxng(
        self,
        queries: list[str],
        out: list[SearchResult],
        seen: set[str],
        target: int,
        fetch_n: int,
    ) -> bool:
        if not self._searx_urls:
            return False

        def add(backend: str, title: str, link: str, snippet: str) -> None:
            key = link.split("#")[0].rstrip("/").lower()
            if key in seen or not link:
                return
            seen.add(key)
            out.append(SearchResult(title=title, link=link, snippet=snippet, backend=backend))

        for q in queries[:2]:
            if len(out) >= target:
                break
            rows, used = await searxng_search_any(q, self._searx_urls, max_results=fetch_n)
            for row in rows:
                add("searxng", row.title, row.link, row.snippet)
            if rows:
                _log(f"SearXNG OK via {used} ({len(rows)} hits) for {q[:45]!r}")
                return True
        return False

    async def search(
        self,
        query: str,
        *,
        after_days: int | None = None,
        market: str = "",
        geo: str = "",
        search_topic: str = "",
        discovery_mode: bool = False,
        validation_mode: bool = False,
    ) -> list[SearchResult]:
        global _probe_printed
        del after_days

        mock = is_mock_run(self._settings)
        cache_ns = _cache_namespace(discovery_mode=discovery_mode, validation_mode=validation_mode)
        if not mock:
            cached = _cached_search_results(query, geo, cache_ns)
            if cached is not None:
                return cached[: self._num]

        try:
            from vendor_intel.clients import search_stats

            search_stats.record_search_start()
        except Exception:
            pass

        if not mock and not _probe_printed:
            print_connectivity_report()
            _log(f"Search stack: {search_stack_description(self._settings)}")
            if self._skip_ddgs:
                _log("SKIP_DDGS=true — ddgs library not used")
            else:
                _log(
                    f"Primary: ddgs.text() backend={ddgs_backend_param()!r} "
                    f"(see https://github.com/deedy5/ddgs)"
                )
            if search_fallbacks_enabled() and not self._skip_ddgs:
                _log("Fallbacks (if ddgs < min hits): Bing HTML, SearXNG")
            elif not search_fallbacks_enabled():
                _log("SEARCH_USE_FALLBACKS=false — ddgs only (+ Wikipedia)")
            if self._searxng_up:
                _log("SearXNG local is up")
            elif self._searxng_tcp and self._searx_urls:
                _log(
                    "SearXNG port open but 0 test hits — docker compose down && docker compose up -d"
                )
            _probe_printed = True

        out: list[SearchResult] = []
        seen: set[str] = set()

        def add(backend: str, title: str, link: str, snippet: str) -> None:
            key = link.split("#")[0].rstrip("/").lower()
            if key in seen or not link:
                return
            seen.add(key)
            out.append(SearchResult(title=title, link=link, snippet=snippet, backend=backend))

        if validation_mode:
            fetch_n = max(self._num, 10)
            target = fetch_n
        elif discovery_mode:
            fetch_n = max(self._num * 2, 30)
            target = fetch_n
        else:
            fetch_n = max(self._num * 2, 20)
            target = fetch_n
        skip_ddgs_this = skip_ddgs(discovery_mode=discovery_mode)
        backup_min = _backup_hit_threshold(
            discovery_mode=discovery_mode,
            validation_mode=validation_mode,
            default_min=self._min,
        )
        q_variants = _query_variants(
            query,
            market,
            geo,
            validation_mode=validation_mode,
            discovery_mode=discovery_mode,
        )
        skip_bing = _env_true("SKIP_BING_HTML") or search_stack_only()
        stack_only = search_stack_only()

        async def _try_bing_html(force: bool = False) -> None:
            if mock or len(out) >= target:
                return
            if stack_only:
                return
            if skip_bing and not force:
                return
            from vendor_intel.clients.bing_html import search_bing_html_try

            bing_variants = q_variants[:3] if discovery_mode else q_variants[:2]
            for bq in bing_variants:
                if len(out) >= target:
                    break
                # search_bing_html_try does blocking network I/O — run it on a
                # worker thread so it doesn't stall the event loop (and every
                # other concurrently-running market's searches along with it).
                rows = await asyncio.to_thread(
                    search_bing_html_try,
                    bq,
                    max_results=min(fetch_n, 15),
                    geo=geo,
                    force=force,
                )
                for r in rows:
                    add("bing_html", r.title, r.link, r.snippet)

        async def _try_ddgs() -> None:
            if mock or skip_ddgs_this or len(out) >= target or not duckduckgo_available():
                return
            max_ddgs_variants = 2 if discovery_mode else (1 if validation_mode else 2)
            for i, q in enumerate(q_variants[:max_ddgs_variants]):
                if len(out) >= target:
                    break
                if i > 0 and len(out) >= self._min:
                    break
                need = min(target - len(out), self._num + 5)
                batch = await duckduckgo_search(q, max_results=need, geo=geo)
                for row in batch:
                    add(getattr(row, "engine", None) or "ddgs", row.title, row.link, row.snippet)
                if not batch and i == 0 and len(out) < self._min:
                    _log(f"DDGS returned 0 for {q[:50]!r}")

        use_fallbacks = search_fallbacks_enabled() and not mock

        if skip_ddgs_this:
            if use_fallbacks:
                await _try_bing_html()
                if (
                    len(out) < backup_min
                    and self._searx_urls
                    and (self._searxng_up or self._searxng_tcp)
                ):
                    _log("Fallback: SearXNG")
                    await self._try_searxng(q_variants, out, seen, target, fetch_n)
                if len(out) < backup_min and skip_bing:
                    await _try_bing_html(force=True)
        else:
            if use_fallbacks and self._prefer_searxng and self._searxng_up and len(out) < backup_min:
                await self._try_searxng(q_variants, out, seen, target, fetch_n)
            await _try_ddgs()
            if use_fallbacks and len(out) < backup_min:
                if skip_bing:
                    _log("Fallback: Bing HTML (ddgs returned few results)")
                    try:
                        from vendor_intel.clients import search_stats

                        search_stats.record_fallback()
                    except Exception:
                        pass
                    await _try_bing_html(force=True)
                else:
                    await _try_bing_html()
            if (
                use_fallbacks
                and len(out) < backup_min
                and self._searx_urls
                and (self._searxng_up or self._searxng_tcp)
                and not self._prefer_searxng
            ):
                _log("Fallback: SearXNG")
                await self._try_searxng(q_variants, out, seen, target, fetch_n)

        wiki_threshold = self._min if discovery_mode else max(2, self._min // 2)
        if not mock and not stack_only and len(out) < wiki_threshold:
            for row in await wikipedia_search(query, max_results=self._num, geo=geo):
                add("wikipedia", row.title, row.link, row.snippet)

        # Agent Reach Exa (optional — mcporter + exa MCP, AGENT_REACH_EXA_SEARCH=true)
        if not mock and not stack_only and len(out) < backup_min:
            try:
                from vendor_intel.integrations.agent_reach_client import exa_search_enabled, exa_web_search

                if exa_search_enabled():
                    exa_hits = await exa_web_search(query, num_results=min(self._num, 8))
                    for hit in exa_hits:
                        add(
                            "agent_reach_exa",
                            hit.get("title") or "",
                            hit.get("link") or "",
                            hit.get("snippet") or "",
                        )
                    if exa_hits:
                        _log(f"Agent Reach Exa OK ({len(exa_hits)} hits) for {query[:45]!r}")
            except Exception:
                pass

        # Google AI Overview scraper (browser extension relay) — boost discovery recall
        if not mock and google_ai_discovery_enabled() and discovery_mode:
            try:
                gai_hits = await google_ai_search(query, max_results=min(12, max(self._num, 8)))
                before = len(out)
                for hit in gai_hits:
                    add("google_ai_scraper", hit.title, hit.link, hit.snippet)
                added = len(out) - before
                if added:
                    _log(f"Google AI scraper OK (+{added} hits) for {query[:45]!r}")
            except Exception as exc:
                _log(f"Google AI scraper skipped: {type(exc).__name__}")

        if not mock and len(out) < backup_min:
            tips: list[str] = []
            if not skip_ddgs_this:
                tips.append(
                    f"try DDGS_BACKENDS=bing,mojeek and DDGS_TIMEOUT=30 "
                    f"(current backend={ddgs_backend_param()!r})"
                )
            if skip_ddgs_this:
                tips.append("discovery uses Bing HTML first (DISCOVERY_SKIP_DDGS or SKIP_DDGS)")
            if use_fallbacks:
                if self._searxng_tcp and not self._searxng_up:
                    tips.append("restart SearXNG: docker compose down && docker compose up -d")
                elif not self._searxng_tcp:
                    tips.append("start SearXNG: docker compose up -d")
                if skip_bing:
                    tips.append("set SKIP_BING_HTML=false for Bing HTML fallback")
            _log("Few results — " + "; ".join(tips))

        from vendor_intel.clients.search_relevance import filter_search_results

        if validation_mode:
            min_keep = 1
        elif discovery_mode:
            min_keep = 1
        else:
            min_keep = 5
        if len(out) < min_keep:
            min_keep = max(1, len(out))

        final = filter_search_results(
            query,
            market or query,
            geo,
            out,
            search_topic=search_topic,
            min_keep=min_keep,
            require_geo_match=discovery_mode and bool(geo) and geo.lower() != "global",
        )[: self._num]
        if not mock and final:
            _put_cached_search_results(query, geo, cache_ns, final)
            try:
                from vendor_intel.clients import search_stats

                search_stats.record_success(final[0].backend if final else "")
                search_stats.persist_summary()
            except Exception:
                pass
        elif not mock and not final:
            try:
                from vendor_intel.clients import search_stats

                search_stats.record_failure()
            except Exception:
                pass
        return final

    async def search_batch(
        self,
        queries: list[str],
        *,
        market: str = "",
        geo: str = "",
        search_topic: str = "",
        discovery_mode: bool = False,
    ) -> dict[str, list[SearchResult]]:
        """Parallel discovery searches — same filters as search(), Phase 2 only."""
        clean = [q.strip() for q in queries if (q or "").strip()]
        if not clean:
            return {}

        mock = is_mock_run(self._settings)
        skip_ddgs_this = skip_ddgs(discovery_mode=discovery_mode)
        use_fallbacks = search_fallbacks_enabled() and not mock
        backup_min = _backup_hit_threshold(
            discovery_mode=discovery_mode,
            validation_mode=False,
            default_min=self._min,
        )

        if discovery_mode:
            fetch_n = max(self._num * 2, 30)
        else:
            fetch_n = max(self._num * 2, 20)

        out_map: dict[str, list[SearchResult]] = {q: [] for q in clean}

        if not mock and not skip_ddgs_this and duckduckgo_available():
            try:
                worker_count = int(os.getenv("DDG_WORKER_COUNT", "0") or "0")
            except ValueError:
                worker_count = 0
            if worker_count > 0 and len(clean) > 1:
                raw_batches = await duckduckgo_search_many(
                    clean, max_results=fetch_n, geo=geo
                )
                for q in clean:
                    seen: set[str] = set()
                    bucket: list[SearchResult] = []
                    for row in raw_batches.get(q) or []:
                        key = row.link.split("#")[0].rstrip("/").lower()
                        if not key or key in seen:
                            continue
                        seen.add(key)
                        bucket.append(
                            SearchResult(
                                title=row.title,
                                link=row.link,
                                snippet=row.snippet,
                                backend=row.engine or "ddgs",
                            )
                        )
                    out_map[q] = bucket

        from vendor_intel.clients.search_relevance import filter_search_results

        min_keep = 1 if discovery_mode else 5

        async def _finalize_one(q: str) -> None:
            out = out_map.get(q) or []
            if len(out) < backup_min and use_fallbacks:
                out = await self.search(
                    q,
                    market=market,
                    geo=geo,
                    search_topic=search_topic,
                    discovery_mode=discovery_mode,
                )
            else:
                keep = max(1, min_keep) if len(out) < min_keep else min_keep
                out = filter_search_results(
                    q,
                    market or q,
                    geo,
                    out,
                    search_topic=search_topic,
                    min_keep=keep,
                    require_geo_match=discovery_mode and bool(geo) and geo.lower() != "global",
                )[: self._num]
            out_map[q] = out

        await asyncio.gather(*[_finalize_one(q) for q in clean])
        return out_map
