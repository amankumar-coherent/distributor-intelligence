"""Phase 3 — parallel smart_crawl enrichment + domain fix + scrape fallback."""
from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path
from typing import Any, Awaitable, Callable

# Project root: vendor_intel/enrichment → src → project
_PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

_cache: dict[str, dict[str, Any]] = {}
_smart_crawl: Callable[..., Awaitable[dict[str, Any]]] | None = None
# Cache-hit prints used to flood Cursor terminals (OSError 22 / process kill).
_cache_hit_logs = 0
_CACHE_HIT_LOG_EVERY = 25


def _log_cache_hit(kind: str, dom: str) -> None:
    global _cache_hit_logs
    _cache_hit_logs += 1
    if _cache_hit_logs == 1 or _cache_hit_logs % _CACHE_HIT_LOG_EVERY == 0:
        print(
            f"  [enrich] {kind}: {dom} (hit #{_cache_hit_logs})",
            flush=True,
        )


def _sync_llm_env() -> None:
    """One key in .env: OPENCODE_API_KEY also powers smart_crawl via OpenAI-compatible API."""
    from backend.config import sync_opencode_to_openai_env

    sync_opencode_to_openai_env()


def _load_smart_crawl() -> Callable[..., Awaitable[dict[str, Any]]]:
    global _smart_crawl
    if _smart_crawl is not None:
        return _smart_crawl
    _sync_llm_env()
    try:
        from crawler.smart_crawl import smart_crawl as fn

        _smart_crawl = fn
        return fn
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            f"{exc}. Install pipeline crawl deps: "
            ".venv\\Scripts\\python.exe -m pip install openai scrapling"
        ) from exc


def clear_enrichment_cache() -> None:
    _cache.clear()


def _crawl_has_content(result: dict[str, Any] | None) -> bool:
    if not result or result.get("error"):
        return False
    data = result.get("data")
    if isinstance(data, dict) and data:
        return len(str(data)) > 80
    pages = result.get("pages") or []
    return bool(pages)


def _summary_len(result: dict[str, Any] | None) -> int:
    if not result:
        return 0
    pages = result.get("pages") or []
    if pages and isinstance(pages[0], dict):
        return len(str(pages[0].get("text") or ""))
    data = result.get("data") or {}
    if isinstance(data, dict):
        intel = data.get("intel") or {}
        if isinstance(intel, dict) and intel.get("summary"):
            return len(str(intel["summary"]))
        return len(json.dumps(data))
    return 0


async def supplement_crawl(
    domain: str,
    existing: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Extra web scrape for weak classify rows — ddgs extract then smart_crawl retry."""
    dom = (domain or "").strip().lower().removeprefix("www.")
    if not dom:
        return existing or {"error": "no_domain", "data": {}}
    if existing and _crawl_has_content(existing) and _summary_len(existing) >= 400:
        return existing

    print(f"  [enrich] supplement crawl: {dom}", flush=True)
    fb = await _ddgs_fallback(dom)
    if fb and _summary_len(fb) >= 120:
        return fb

    try:
        smart_crawl = _load_smart_crawl()
        result = await smart_crawl(dom, mode="company")
        if _crawl_has_content(result) and _summary_len(result) >= 120:
            print(f"  [enrich] supplement ok (smart_crawl): {dom}", flush=True)
            return result
    except Exception as exc:
        print(f"  [enrich] supplement smart_crawl failed: {dom} - {exc}", flush=True)

    return fb or existing or {"error": "thin_content", "data": {}, "domain": dom}


async def _ddgs_fallback(domain: str) -> dict[str, Any] | None:
    """When local DNS/HTTP crawl fails, ddgs.extract often still returns markdown."""
    try:
        from vendor_intel.scraping.ddgs_extract import fetch_page_via_ddgs_extract

        url = f"https://{domain}/"
        page = await asyncio.to_thread(fetch_page_via_ddgs_extract, url)
        if not page.alive or len((page.text or "").strip()) < 80:
            return None
        text = (page.text or "")[:12000]
        print(f"  [enrich] ddgs.extract fallback OK: {domain} ({len(text)} chars)", flush=True)
        return {
            "domain": domain,
            "data": {
                "company": {"name": domain},
                "business": {"products": [], "services": []},
                "intel": {"summary": text[:4000]},
                "source": "ddgs_extract_fallback",
            },
            "pages": [{"url": url, "text": text}],
        }
    except Exception as exc:
        print(f"  [enrich] ddgs fallback failed: {domain} - {exc}", flush=True)
        return None


async def _ssc_only(dom: str) -> dict[str, Any] | None:
    from vendor_intel.enrichment.ssc_fetch import fetch_server_side_content

    result = await fetch_server_side_content(dom)
    if result.get("error") and not result.get("pages"):
        return None
    print(
        f"  [enrich] SSC OK: {dom} via {result.get('source', 'ssc')} "
        f"({len(str((result.get('pages') or [{}])[0].get('text', '')))} chars)",
        flush=True,
    )
    return result


def _cross_market_cache_get(dom: str, settings: Any) -> dict[str, Any] | None:
    if not bool(getattr(settings, "company_cache_enabled", True)):
        return None
    try:
        from vendor_intel.batch.db import get_cached_company

        return get_cached_company(
            dom, max_age_days=float(getattr(settings, "company_cache_max_age_days", 30.0))
        )
    except Exception as exc:
        print(f"  [enrich] company cache read skipped: {exc}", flush=True)
        return None


def _cross_market_cache_put(dom: str, name: str, source: str, result: dict[str, Any], settings: Any) -> None:
    if not bool(getattr(settings, "company_cache_enabled", True)):
        return
    try:
        from vendor_intel.batch.db import upsert_company_cache

        upsert_company_cache(dom, result, canonical_name=name, crawl_source=source, success=True)
    except Exception as exc:
        print(f"  [enrich] company cache write skipped: {exc}", flush=True)


async def _crawl_one(
    name: str,
    domain: str,
    *,
    country: str = "",
    use_ssc: bool = False,
) -> tuple[str, dict[str, Any] | None, str | None]:
    from vendor_intel.utils.domain_corrections import fix_company_domain

    from vendor_intel.config import Settings as _Settings

    settings = _Settings.load()
    _recall = bool(getattr(settings, "pipeline_recall_mode", False))
    fixed = fix_company_domain(
        {"name": name, "domain": domain}, country=country, recall_mode=_recall
    )
    from vendor_intel.utils.domain_corrections import crawl_host_for_domain

    dom = crawl_host_for_domain((fixed.get("domain") or "").strip().lower())
    if not dom:
        print(f"  [enrich] skip (no domain): {name[:50]}", flush=True)
        return name, None, "no_domain"

    if dom in _cache:
        _log_cache_hit("cache hit", dom)
        return name, _cache[dom], None

    cross_market = _cross_market_cache_get(dom, settings)
    if cross_market is not None:
        _log_cache_hit("cross-market cache hit", dom)
        _cache[dom] = cross_market
        return name, cross_market, None

    print(f"  [enrich] crawl start: {name[:40]} ({dom})", flush=True)
    if use_ssc:
        result = await _ssc_only(dom)
        if result:
            _cache[dom] = result
            _cross_market_cache_put(dom, name, "ssc", result, settings)
            return name, result, None
        fb = await _ddgs_fallback(dom)
        if fb:
            _cache[dom] = fb
            _cross_market_cache_put(dom, name, "ddgs_extract_fallback", fb, settings)
            return name, fb, None
        return name, None, "ssc_failed"

    try:
        smart_crawl = _load_smart_crawl()
        result = await smart_crawl(dom, mode="company")
        source = "smart_crawl"
        if not _crawl_has_content(result):
            fb = await _ddgs_fallback(dom)
            if fb:
                result = fb
                source = "ddgs_extract_fallback"
        _cache[dom] = result
        _cross_market_cache_put(dom, name, source, result, settings)
        print(f"  [enrich] crawl ok: {dom}", flush=True)
        return name, result, None
    except Exception as exc:
        err = f"{type(exc).__name__}: {exc}"[:200]
        fb = await _ddgs_fallback(dom)
        if fb:
            _cache[dom] = fb
            _cross_market_cache_put(dom, name, "ddgs_extract_fallback", fb, settings)
            print(f"  [enrich] crawl ok (fallback): {dom}", flush=True)
            return name, fb, None
        print(f"  [enrich] crawl failed: {dom} - {err}", flush=True)
        return name, None, err


async def enrich_companies(
    companies: list[dict[str, str]],
    limit: int = 80,
    *,
    max_concurrent: int = 4,
    country: str = "",
    use_ssc: bool | None = None,
    checkpoint_path: Path | str | None = None,
    resume: bool = True,
    query: str = "",
    market_slug: str | None = None,
) -> dict[str, Any]:
    """
    Run smart_crawl in parallel for up to `limit` companies.

    Input: [{"name": str, "domain": str}, ...]
    Returns: {company_name: smart_crawl_output}

    When ``checkpoint_path`` is set, already-enriched companies are skipped and
    each new result is written so a crash mid-enrich can resume.
    """
    from vendor_intel.config import Settings
    from vendor_intel.pipeline.checkpoint_resume import (
        company_already_enriched,
        load_phase3_checkpoint,
        save_phase3_checkpoint,
    )

    settings = Settings.load()
    if use_ssc is None:
        use_ssc = bool(getattr(settings, "pipeline_use_ssc", True))
    if not use_ssc:
        _load_smart_crawl()
    else:
        print("  [enrich] mode=SSC (server-side content, fast)", flush=True)
    batch = list(companies or [])[: max(0, limit)]
    if not batch:
        print("  [enrich] no companies to enrich", flush=True)
        return {}

    ck_path = Path(checkpoint_path) if checkpoint_path else None
    out: dict[str, Any] = {}
    if resume and ck_path is not None:
        prev = load_phase3_checkpoint(ck_path)
        if prev and isinstance(prev.get("enriched"), dict):
            out.update(prev["enriched"])
            if out:
                print(
                    f"  [enrich] RESUME checkpoint — {len(out)} companies already enriched "
                    f"({ck_path.name})",
                    flush=True,
                )

    todo: list[dict[str, str]] = []
    for c in batch:
        if company_already_enriched(c, out):
            # Ensure both name and domain keys exist for classify lookup
            name = str(c.get("name") or "").strip()
            dom = str(c.get("domain") or "").strip().lower()
            payload = None
            for k in (name, dom, dom[4:] if dom.startswith("www.") else ""):
                if k and k in out:
                    payload = out[k]
                    break
            if payload is not None:
                if name:
                    out[name] = payload
                if dom:
                    out[dom] = payload
            continue
        todo.append(c)

    skipped = len(batch) - len(todo)
    print(
        f"  [enrich] enriching {len(todo)} companies "
        f"(skip already done={skipped}, limit={limit}, concurrent={max_concurrent})",
        flush=True,
    )
    if not todo:
        print(f"  [enrich] done: {len(out)} entries (all from checkpoint)", flush=True)
        return out

    sem = asyncio.Semaphore(max(1, max_concurrent))
    done_new = 0
    lock = asyncio.Lock()

    def _persist() -> None:
        if ck_path is None:
            return
        save_phase3_checkpoint(
            ck_path,
            query=query,
            market_slug=market_slug,
            companies=None,  # keep existing companies snapshot
            enriched=out,
            merge=True,
        )

    async def _run_one(c: dict[str, str]) -> None:
        nonlocal done_new
        async with sem:
            name, data, err = await _crawl_one(
                str(c.get("name") or "").strip(),
                str(c.get("domain") or "").strip(),
                country=country,
                use_ssc=use_ssc,
            )
            key_name = name or "unknown"
            dom = str(c.get("domain") or "").strip().lower()
            payload: dict[str, Any]
            if data is not None:
                payload = data
            elif err:
                payload = {"error": err, "domain": dom, "data": {}}
            else:
                payload = {"error": "empty", "domain": dom, "data": {}}
            async with lock:
                out[key_name] = payload
                if dom:
                    out[dom] = payload
                done_new += 1
                if ck_path is not None and (done_new % 5 == 0 or done_new == len(todo)):
                    _persist()
                    print(
                        f"  [enrich] checkpoint {done_new}/{len(todo)} "
                        f"(total keys={len(out)})",
                        flush=True,
                    )

    await asyncio.gather(*[_run_one(c) for c in todo])
    if ck_path is not None:
        _persist()

    print(f"  [enrich] done: {len(out)} entries", flush=True)
    return out
