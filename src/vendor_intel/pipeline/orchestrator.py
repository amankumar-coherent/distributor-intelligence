"""Pipeline: Plan → Discovery → SSC enrich → Quality classify → CSV (~20 min)."""
from __future__ import annotations

import asyncio
import json
import time
from datetime import datetime, timezone
from typing import Any

from vendor_intel.config import Settings
from vendor_intel.enrichment.smart_enrichment import clear_enrichment_cache, enrich_companies
from vendor_intel.intelligence.classifier import classify_company
from vendor_intel.intelligence.signal_extractor import extract_signals
from vendor_intel.phase1.runner import run_phase1
from vendor_intel.phase2.discovery_fast import run_phase2_fast
from vendor_intel.pipeline.entity_gate import filter_companies, reject_reason
from vendor_intel.pipeline.llm_meter import get_meter, reset_meter
from vendor_intel.pipeline.plan_seeds import load_scope_from_plan, merge_seeds_first
from vendor_intel.pipeline.csv_fields import (
    extract_company_summary,
    extract_operational_presence,
    format_data_sources,
)
from vendor_intel.pipeline.quality_export import filter_for_export


def _build_query(query_context: dict[str, Any]) -> str:
    industry = str(query_context.get("industry") or "").strip()
    country = str(query_context.get("country") or "global").strip()
    functions = query_context.get("functions") or []
    fn = ", ".join(str(f) for f in functions[:5] if str(f).strip())
    parts = [p for p in [industry, f"in {country}" if country and country != "global" else "", fn] if p]
    return " ".join(parts).strip() or industry or "market companies"


def _industry_keywords(query_context: dict[str, Any], scope: dict[str, Any]) -> list[str]:
    from vendor_intel.pipeline.participant_domains import filter_industry_match_terms

    kws: list[str] = []
    kws.append(str(query_context.get("industry") or ""))
    kws.extend(str(f) for f in query_context.get("functions") or [])
    for field in ("relevance_keywords", "industry_terms"):
        kws.extend(str(x) for x in (scope.get(field) or []))
    out: list[str] = []
    seen: set[str] = set()
    for k in filter_industry_match_terms(
        [x.strip().lower() for x in kws if x and str(x).strip()]
    ):
        if k not in seen:
            seen.add(k)
            out.append(k)
    return out


def _profile_limits(
    settings: Settings, recall: bool, country: str | None = None
) -> tuple[int, int, float]:
    from vendor_intel.pipeline.geo_limits import pipeline_limits

    lim = pipeline_limits(settings, recall=recall, country=country)
    return int(lim["discover"]), int(lim["enrich"]), float(lim["min_conf"])


def _enrichment_source(smart_data: dict[str, Any]) -> str:
    if not smart_data:
        return "none"
    if smart_data.get("source"):
        return str(smart_data["source"])
    if smart_data.get("error"):
        return f"error:{smart_data.get('error')}"[:40]
    return "smart_crawl"


def _website_display(domain: str, smart_data: dict[str, Any]) -> str:
    dom = (domain or "").strip()
    if not dom:
        return ""
    data = smart_data.get("data") or {}
    if isinstance(data, dict):
        co = data.get("company") or {}
        if isinstance(co, dict) and co.get("website"):
            return str(co["website"])
    return dom if dom.startswith("http") else f"https://{dom}"


_REJECT_FRIENDLY: dict[str, str] = {
    "not_relevant": "the site did not clearly mention this market - relevance could not be confirmed",
    "non_product_site": "not a company/product site (directory, blog, marketplace, or aggregator)",
    "market_research_site": "a market-research / news / data site, not a market participant",
    "consulting_not_participant": "appears to be consulting/advisory, not a market participant",
    "below_quality_threshold": "too little evidence on the site to confirm",
    "weak_product_fit": "products shown on the site did not match the market closely enough",
    "off_topic_pharma": "site content pointed to a different (pharma) domain",
    "off_topic_pharma_cdmo": "site content pointed to a different (pharma) domain",
    "duplicate_domain": "duplicate of another listed company",
}


def _friendly_reject(code: str) -> str:
    code = (code or "").strip()
    if not code:
        return "could not confirm this company operates in the market"
    if code.startswith("geo_mismatch"):
        return (
            f"headquartered in {code.split(':', 1)[-1]} with no detected operations in the "
            f"target geography"
        )
    return _REJECT_FRIENDLY.get(code, f"could not confirm ({code})")


def _verification_status(v: dict[str, Any]) -> tuple[bool, str]:
    """Is an exported row actually substantiated? Reason if not.

    Seeds / known players (we already resolved their official domain) are NEVER demoted to
    'Not Verified' for a mere crawl failure — a transient fetch error must not banish a
    recognised major. They only leave the list if classified clearly off-market.
    """
    src = str(v.get("enrichment_source") or "").lower()
    crawl_failed = (not src) or src in ("none", "no_crawl") or src.startswith("error") or "fail" in src
    if not v.get("is_relevant"):
        return False, _friendly_reject(str(v.get("reject_reason") or ""))
    if v.get("is_seed"):
        return True, ""  # known/pinned player with a resolved domain — keep it
    if crawl_failed:
        return (
            False,
            "website could not be retrieved - listed by name only; its site did not confirm "
            "activity in this market",
        )
    return True, ""


async def run_pipeline(
    query_context: dict[str, Any],
    settings: Settings | None = None,
    *,
    enrich_limit: int | None = None,
    classify_limit: int | None = None,
) -> dict[str, Any]:
    settings = settings or Settings.load()
    recall = bool(getattr(settings, "pipeline_recall_mode", False))
    profile = str(getattr(settings, "pipeline_profile", "quality") or "quality")
    quality = not recall and profile in ("quality", "balanced")
    country = str(query_context.get("country") or "global")
    from vendor_intel.pipeline.geo_limits import is_global_geography, pipeline_limits

    lim = pipeline_limits(settings, recall=recall, country=country)
    discover_cap, enrich_cap, min_conf = (
        int(lim["discover"]),
        int(lim["enrich"]),
        float(lim["min_conf"]),
    )
    classify_cap = classify_limit if classify_limit is not None else discover_cap
    enrich_cap = enrich_limit if enrich_limit is not None else enrich_cap
    export_max = int(lim["export_max"])
    export_min = int(lim["export_min"])
    min_quality = float(lim["min_quality"])
    geo_label = "global (higher volume)" if is_global_geography(country) else country

    reset_meter()
    t0 = time.perf_counter()
    query = _build_query(query_context)
    print(f"\n=== Pipeline start ===\n  Query: {query}", flush=True)
    print(
        f"  Profile: {profile}" + (" (recall/noisy)" if recall else " (quality)"),
        flush=True,
    )
    print(
        f"  Geography: {geo_label} - discover<={discover_cap}, enrich<={enrich_cap}, "
        f"export up to {export_max} rows (no padding)",
        flush=True,
    )
    use_ssc = bool(getattr(settings, "pipeline_use_ssc", True)) and profile != "deep"
    print(f"  Enrichment: {'SSC server-side' if use_ssc else 'smart_crawl'}", flush=True)

    clear_enrichment_cache()

    # When the user specifies functional categories, steer discovery toward them:
    # the market-understanding LLM generates prompts for these company types only.
    focus_sections = [
        str(s).strip()
        for s in (query_context.get("sections") or [])
        if str(s).strip() and str(s).strip().lower() != "other"
    ]

    # Channel Partner Intelligence — infer market channel types before Phase 1.
    from vendor_intel.funnel.channel_partner_types import (
        apply_channel_types_to_scope,
        channel_partner_mode_enabled,
        infer_channel_partner_types,
    )

    if channel_partner_mode_enabled(query_context, settings=settings):
        market_label = str(query_context.get("industry") or query or "")
        types = list(
            query_context.get("channel_partner_types")
            or infer_channel_partner_types(market_label, query)
        )
        query_context["discovery_mode"] = "channel_partners"
        query_context["channel_partner_types"] = types
        if not focus_sections:
            focus_sections = [str(t).strip() for t in types if str(t).strip()]
        print(
            f"  [pipeline] channel-partner mode — types: {', '.join(types)}",
            flush=True,
        )

    plan_query = query
    if focus_sections:
        label = (
            "channel-partner types"
            if channel_partner_mode_enabled(query_context, settings=settings)
            else "company functions / types"
        )
        plan_query = (
            f"{query} -- only include these {label}: "
            f"{'; '.join(focus_sections)}"
        )
        print(
            f"  [pipeline] discovery focused on {len(focus_sections)} requested {label}",
            flush=True,
        )

    resume_checkpoint = bool(query_context.get("resume_checkpoint"))
    market_slug_hint = str(query_context.get("market_slug") or "").strip() or None
    skip_phase1 = False
    skip_phase2 = False
    plan_path: str | None = None
    companies: list[dict[str, str]] = []
    # Always defined: resume path used to leave `p2` unset and crash at return
    # ("cannot access local variable 'p2'") after Phase 3/4 completed.
    p2: dict[str, Any] = {}

    if resume_checkpoint:
        from vendor_intel.pipeline.checkpoint_resume import (
            find_phase1_plan,
            find_phase2_companies,
            load_phase2_companies,
        )

        ck_p1 = find_phase1_plan(query=query, market_slug=market_slug_hint)
        ck_p2 = find_phase2_companies(
            query=query, market_slug=market_slug_hint, phase1_plan=ck_p1
        )
        if ck_p1 is not None:
            plan_path = str(ck_p1.resolve())
            skip_phase1 = True
            print(
                f"  [pipeline] RESUME checkpoint — reuse Phase 1 plan: {ck_p1.name}",
                flush=True,
            )
        if ck_p1 is not None and ck_p2 is not None:
            companies = load_phase2_companies(ck_p2)
            if companies:
                skip_phase2 = True
                p2 = {
                    "output_path": str(ck_p2.resolve()),
                    "companies": companies,
                }
                print(
                    f"  [pipeline] RESUME checkpoint — reuse Phase 2 list: {ck_p2.name} "
                    f"({len(companies)} companies) → skip discovery, continue enrich/classify",
                    flush=True,
                )
            else:
                print(
                    f"  [pipeline] RESUME checkpoint — Phase 2 file empty ({ck_p2.name}); "
                    f"will re-run discovery",
                    flush=True,
                )

    if skip_phase1 and plan_path:
        scope = load_scope_from_plan(plan_path)
    else:
        print("  [pipeline] Phase 1 — query plan", flush=True)
        plan = await run_phase1(plan_query, settings)
        get_meter().add_phase1()
        plan_path = plan.get("plan_path") or plan.get("output_path")
        scope = load_scope_from_plan(str(plan_path) if plan_path else None)

    if channel_partner_mode_enabled(query_context, settings=settings):
        market_label = str(
            scope.get("market") or query_context.get("industry") or query or ""
        )
        scope = apply_channel_types_to_scope(
            scope, market_label, query, query_context=query_context
        )
        query_context["channel_partner_types"] = list(
            scope.get("channel_partner_types")
            or query_context.get("channel_partner_types")
            or []
        )
    industry_kws = _industry_keywords(query_context, scope)

    # Analyst-provided seed companies: resolve domains and guarantee them into the run.
    # Skipped when resuming from a Phase 2 company list (already baked into that list).
    user_seed_names = [
        str(s).strip() for s in (query_context.get("seed_companies") or []) if str(s).strip()
    ]
    unresolved_seed_names: list[str] = []
    if user_seed_names and not skip_phase2:
        from vendor_intel.pipeline.plan_seeds import (
            merge_user_seeds_into_scope,
            resolve_user_seeds,
        )

        seed_pins = {
            str(k): str(v)
            for k, v in (query_context.get("seed_domains") or {}).items()
            if str(k).strip() and str(v).strip()
        }
        resolved_seeds, unresolved_seed_names = await resolve_user_seeds(
            user_seed_names, settings, country, pinned_domains=seed_pins
        )
        scope = merge_user_seeds_into_scope(scope, resolved_seeds, analyst=True)
        print(
            f"  [pipeline] {len(resolved_seeds)}/{len(user_seed_names)} analyst seeds "
            f"resolved and pinned into the run",
            flush=True,
        )
    elif user_seed_names and skip_phase2:
        print(
            f"  [pipeline] RESUME — skip seed resolve ({len(user_seed_names)} seeds already in Phase 2 list)",
            flush=True,
        )

    # Widen the REAL-company base from the LLM's own knowledge (what GPT/Google would
    # list). These are named, real companies — NOT web junk — so it boosts recall without
    # loosening the relevance gate. They still go through normal enrichment + classification.
    if (
        not skip_phase2
        and not recall
        and getattr(settings, "pipeline_enumerate_players", True)
    ):
        try:
            from vendor_intel.clients.claude import ClaudeClient
            from vendor_intel.funnel.market_understanding import enumerate_market_players
            from vendor_intel.pipeline.plan_seeds import merge_user_seeds_into_scope

            section_hints = [
                str(s).strip() for s in (query_context.get("sections") or []) if str(s).strip()
            ]
            players = await asyncio.to_thread(
                enumerate_market_players,
                str(scope.get("market") or query_context.get("industry") or ""),
                country,
                section_hints,
                ClaudeClient(settings),
                settings,
                market_definition=str(scope.get("market_definition") or ""),
            )
            if players:
                before = len(scope.get("seed_companies") or [])
                scope = merge_user_seeds_into_scope(scope, players)
                get_meter().add_market_understanding()
                print(
                    f"  [pipeline] enumerated {len(players)} known market players (LLM) "
                    f"— seed pool {before} -> {len(scope.get('seed_companies') or [])}",
                    flush=True,
                )
        except Exception as exc:
            print(f"  [pipeline] player enumeration skipped: {exc}", flush=True)

    if skip_phase2:
        print(
            f"  [pipeline] Phase 2 — skipped (checkpoint) — {len(companies)} companies loaded",
            flush=True,
        )
    else:
        print("  [pipeline] Phase 2 — discovery + entity gate", flush=True)
        p2 = await run_phase2_fast(query, settings, phase1_plan_path=plan_path)
        companies = list(p2.get("companies") or [])

        if not recall and getattr(settings, "pipeline_multi_source_discovery_enabled", True):
            try:
                from vendor_intel.pipeline.multi_source_discovery import discover_additional_sources

                existing_doms = {(c.get("domain") or "").strip().lower() for c in companies}
                extra = await discover_additional_sources(
                    query_context, scope, settings, geo=country
                )
                new_extra = [
                    c for c in extra if c.get("domain", "").lower() not in existing_doms
                ]
                if new_extra:
                    companies.extend(new_extra)
                    print(
                        f"  [pipeline] +{len(new_extra)} candidates from multi-source discovery "
                        f"(directories/LinkedIn/Owler)",
                        flush=True,
                    )
            except Exception as exc:
                print(f"  [pipeline] multi-source discovery skipped: {exc}", flush=True)

    companies, seed_doms = merge_seeds_first(companies, scope)
    print(f"  [pipeline] {len(companies)} after seeds merge ({len(seed_doms)} seeds)", flush=True)

    if not recall:
        companies, rejected_p2 = filter_companies(companies)
        if rejected_p2:
            print(f"  [pipeline] entity_gate removed {len(rejected_p2)} non-participants", flush=True)

    cleaned: list[dict[str, str]] = []
    seen_dom: set[str] = set()
    from vendor_intel.pipeline.quality_export import _export_domain

    for c in companies:
        dom = _export_domain(c)
        name = (c.get("name") or "").strip()
        if not dom or not name or dom in seen_dom:
            continue
        seen_dom.add(dom)
        cleaned.append(c)

    from vendor_intel.utils.domain_corrections import fix_company_list

    p3_path = None
    p3_ckpt = None
    if resume_checkpoint:
        from vendor_intel.pipeline.checkpoint_resume import (
            load_phase3_checkpoint,
            phase3_enrich_path,
            save_phase3_checkpoint,
        )

        p3_path = phase3_enrich_path(query=query, market_slug=market_slug_hint)
        p3_ckpt = load_phase3_checkpoint(p3_path)

    # Reuse domain-fixed list from a prior crash so we don't re-print hundreds of [domain] lines.
    reused_domain_list = False
    if (
        resume_checkpoint
        and p3_ckpt
        and isinstance(p3_ckpt.get("companies"), list)
        and p3_ckpt["companies"]
    ):
        top = []
        for c in p3_ckpt["companies"]:
            if not isinstance(c, dict):
                continue
            name = str(c.get("name") or "").strip()
            dom = str(c.get("domain") or "").strip()
            if name and dom:
                top.append(c)
        if top:
            reused_domain_list = True
            print(
                f"  [pipeline] RESUME checkpoint — reuse domain-fixed list "
                f"({len(top)} companies) → skip DNS re-check",
                flush=True,
            )

    if not reused_domain_list:
        top = fix_company_list(cleaned[:classify_cap], country=country, recall_mode=recall)
        if resume_checkpoint and p3_path is not None:
            from vendor_intel.pipeline.checkpoint_resume import save_phase3_checkpoint

            save_phase3_checkpoint(
                p3_path,
                query=query,
                market_slug=market_slug_hint,
                companies=top,
                enriched=None,
                merge=True,
            )
    # DNS/alias fix can change host (e.g. braskem.com → www.braskem.com.br) — refresh seed_doms for export
    # Accept both is_seed / is_analyst_seed and legacy analyst_seed flag from curated Phase 2 JSON.
    for _c in top:
        if _c.get("analyst_seed") and not _c.get("is_analyst_seed"):
            _c["is_analyst_seed"] = True
        if _c.get("is_analyst_seed") or _c.get("analyst_seed"):
            _c["is_seed"] = True
    seed_doms = {
        (c.get("domain") or "").strip().lower()
        for c in top
        if (c.get("is_seed") or c.get("is_analyst_seed") or c.get("analyst_seed"))
        and (c.get("domain") or "").strip()
    }
    # Off-market pre-filter: drop candidates whose name/domain clearly signals a
    # different industry (scope-driven exclude/negative keywords) BEFORE enrichment,
    # so we don't waste crawls on obvious junk. Seeds are exempt.
    if not recall and not reused_domain_list:
        from vendor_intel.pipeline.market_relevance import exclude_keyword_hit, keyword_profile

        prof = keyword_profile(scope, query_context)
        kept_top: list[dict[str, str]] = []
        off_market = 0
        for c in top:
            nm = (c.get("name") or "").strip()
            dom = (c.get("domain") or "").strip()
            if c.get("is_seed") or dom in seed_doms:
                kept_top.append(c)
                continue
            if exclude_keyword_hit(f"{nm} {dom}", dom, prof, name=nm):
                off_market += 1
                continue
            kept_top.append(c)
        if off_market:
            top = kept_top
            print(
                f"  [pipeline] off-market pre-filter dropped {off_market} before enrichment",
                flush=True,
            )
            if resume_checkpoint and p3_path is not None:
                from vendor_intel.pipeline.checkpoint_resume import save_phase3_checkpoint

                save_phase3_checkpoint(
                    p3_path,
                    query=query,
                    market_slug=market_slug_hint,
                    companies=top,
                    enriched=None,
                    merge=True,
                )
    print(f"  [pipeline] {len(top)} companies to process", flush=True)

    print("  [pipeline] Phase 3 — enrichment", flush=True)
    effective_enrich = min(len(top), max(enrich_cap, discover_cap))
    enrich_batch = top[:effective_enrich]
    enrich_concurrent = int(getattr(settings, "pipeline_enrich_concurrent", 8) or 8)
    enriched = await enrich_companies(
        enrich_batch,
        limit=effective_enrich,
        max_concurrent=enrich_concurrent,
        country=country,
        use_ssc=use_ssc,
        checkpoint_path=p3_path if resume_checkpoint else None,
        resume=resume_checkpoint,
        query=query,
        market_slug=market_slug_hint,
    )

    print("  [pipeline] Phase 4 — classification (quality)", flush=True)
    from vendor_intel.clients.claude import ClaudeClient

    client = ClaudeClient(settings)
    from vendor_intel.pipeline.sections import build_section_taxonomy, main_product_label

    main_product = main_product_label(query_context, scope)
    custom_sections = [
        str(s).strip() for s in (query_context.get("sections") or []) if str(s).strip()
    ]
    if custom_sections:
        if not any(s.lower() == "other" for s in custom_sections):
            custom_sections = custom_sections + ["Other"]
        section_taxonomy = custom_sections
        print(f"  [pipeline] using {len(custom_sections)} custom CEO sections", flush=True)
    else:
        section_taxonomy = build_section_taxonomy(main_product)
    classify_ctx = {
        **query_context,
        "plan_keywords": industry_kws,
        "scope": scope,
        "market": str(scope.get("market") or query_context.get("industry") or ""),
        "main_product": main_product,
        "value_chain_sections": section_taxonomy,
    }
    classify_concurrent = int(getattr(settings, "pipeline_classify_concurrent", 8) or 8)

    from vendor_intel.pipeline.checkpoint_resume import (
        company_already_classified,
        load_phase4_checkpoint,
        lookup_classified,
        phase4_classify_path,
        save_phase4_checkpoint,
    )

    p4_path = (
        phase4_classify_path(query=query, market_slug=market_slug_hint)
        if resume_checkpoint
        else None
    )
    classified_map: dict[str, Any] = {}
    if resume_checkpoint and p4_path is not None:
        p4_ckpt = load_phase4_checkpoint(p4_path)
        if p4_ckpt and isinstance(p4_ckpt.get("classified"), dict):
            classified_map.update(p4_ckpt["classified"])
            if classified_map:
                print(
                    f"  [pipeline] RESUME checkpoint — "
                    f"{int(p4_ckpt.get('classified_count') or 0)} companies already classified "
                    f"({p4_path.name})",
                    flush=True,
                )

    classify_todo = [
        c for c in top if not company_already_classified(c, classified_map)
    ]
    skipped_cls = len(top) - len(classify_todo)
    print(
        f"  [pipeline] classifying {len(classify_todo)} companies "
        f"(skip already done={skipped_cls}, concurrent={classify_concurrent})",
        flush=True,
    )
    sem = asyncio.Semaphore(max(1, classify_concurrent))
    done = 0
    cls_lock = asyncio.Lock()

    def _persist_classify() -> None:
        if p4_path is None:
            return
        save_phase4_checkpoint(
            p4_path,
            query=query,
            market_slug=market_slug_hint,
            classified=classified_map,
            merge=True,
        )

    async def _classify_one(c: dict[str, str]) -> None:
        nonlocal done
        name = c["name"]
        dom = c["domain"]
        smart_data = enriched.get(name) or enriched.get(dom) or {"error": "no_crawl", "data": {}}
        is_seed = (
            dom in seed_doms
            or bool(c.get("is_seed"))
            or bool(c.get("analyst_seed"))
            or bool(c.get("is_analyst_seed"))
        )
        is_analyst_seed = bool(c.get("is_analyst_seed") or c.get("analyst_seed"))
        async with sem:
            signals = extract_signals(smart_data, industry_keywords=industry_kws)
            verdict = await classify_company(
                c,
                smart_data,
                classify_ctx,
                settings=settings,
                client=client,
                recall_mode=recall,
                quality_mode=quality,
                is_seed=is_seed,
            )
        supplemented = verdict.pop("_enriched", None)
        classify_signals = verdict.pop("_classify_signals", None)
        if supplemented is not None:
            smart_data = supplemented
            enriched[name] = supplemented
            enriched[dom] = supplemented
        verdict["signals"] = classify_signals or extract_signals(
            smart_data, industry_keywords=industry_kws
        )
        verdict["discovery_source"] = c.get("discovery_source", "search")
        verdict["enrichment_source"] = _enrichment_source(smart_data)
        verdict["website"] = _website_display(dom, smart_data)
        verdict["is_seed"] = is_seed
        verdict["is_analyst_seed"] = is_analyst_seed
        verdict["company_function"] = (
            c.get("company_function") or verdict.get("company_function") or ""
        )
        # Summary comes from the LLM's clean prose field; the CSV layer composes a
        # clean fallback from role/products. We deliberately do NOT use raw scraped
        # page text here (it leaks nav menus, cookie banners, markdown, mojibake).
        verdict["company_summary"] = str(verdict.get("summary") or "")
        verdict["operational_presence"] = extract_operational_presence(
            smart_data,
            str(query_context.get("country") or ""),
            signals=signals,
        )
        verdict["data_sources"] = format_data_sources(
            str(c.get("discovery_source") or verdict.get("discovery_source") or ""),
            verdict["enrichment_source"],
            smart_data,
        )
        from vendor_intel.export.fmcg_landscape import crawl_snapshot

        verdict["_crawl"] = crawl_snapshot(smart_data)
        # Stable identity for resume lookup
        verdict.setdefault("company", name)
        verdict.setdefault("domain", dom)
        async with cls_lock:
            classified_map[name] = verdict
            if dom:
                classified_map[dom] = verdict
            done += 1
            if p4_path is not None and (done % 5 == 0 or done == len(classify_todo)):
                _persist_classify()
                print(
                    f"  [pipeline] classified checkpoint {done}/{len(classify_todo)} "
                    f"(skip already done={skipped_cls})",
                    flush=True,
                )
            elif done % 10 == 0 or done == len(classify_todo):
                print(
                    f"  [pipeline] classified {done}/{len(classify_todo)} "
                    f"(skip already done={skipped_cls})",
                    flush=True,
                )

    if classify_todo:
        await asyncio.gather(*[_classify_one(c) for c in classify_todo])
        if p4_path is not None:
            _persist_classify()

    # Rebuild in `top` order (checkpoint + newly classified)
    classified: list[dict[str, Any]] = []
    missing = 0
    for c in top:
        v = lookup_classified(c, classified_map)
        if v is None:
            missing += 1
            # Should not happen; keep a stub so export does not crash
            classified.append(
                {
                    "company": c.get("name"),
                    "domain": c.get("domain"),
                    "is_relevant": False,
                    "confidence": 0.0,
                    "error": "classify_checkpoint_missing",
                }
            )
        else:
            classified.append(v)
    if missing:
        print(
            f"  [pipeline] WARN classify resume: {missing} companies missing from checkpoint",
            flush=True,
        )
    # Honor analyst-assigned sections for must-haves ("Name | Section" in the seed file):
    # the pinned company is placed in that exact section regardless of keyword routing.
    seed_sections = query_context.get("seed_sections") or {}
    if seed_sections:
        import re as _re

        def _ck(s: str) -> str:
            return _re.sub(r"[^a-z0-9]", "", str(s or "").lower())

        smap = {_ck(k): v for k, v in seed_sections.items() if v}
        for v in classified:
            cn = _ck(v.get("company") or v.get("brand") or "")
            if not cn:
                continue
            for sk, sec in smap.items():
                if sk and (sk in cn or cn in sk):
                    v["_forced_section"] = sec
                    break

    if recall:
        final = [v for v in classified if float(v.get("confidence") or 0) >= min_conf]
        if not final:
            final = list(classified)
        export_rejected: list[dict[str, Any]] = []
    else:
        from collections import Counter

        final, export_rejected = filter_for_export(
            classified,
            enriched,
            query_context,
            scope=scope,
            seed_domains=seed_doms,
            min_rows=export_min,
            max_rows=export_max,
            min_quality=min_quality,
            # Honor PIPELINE_GLOBAL_EXPORT_MIN_ROWS (≥300) by soft-padding
            # borderline real companies instead of shipping ~60–70 rows.
            pad_to_min=bool(export_min and export_min > 0),
        )
        if export_rejected:
            reasons = Counter(
                str(r.get("export_reject") or "?") for r in export_rejected
            )
            top = ", ".join(f"{k}={v}" for k, v in reasons.most_common(4))
            print(f"  [pipeline] export rejects ({len(export_rejected)}): {top}", flush=True)
        roles = Counter(str(r.get("role") or "?") for r in final) if final else Counter()
        if roles:
            print(
                f"  [pipeline] export roles: "
                + ", ".join(f"{k}={v}" for k, v in roles.most_common()),
                flush=True,
            )

        # If channel-partner export is still below min (e.g. 300), keep searching
        # the web for distributors until the floor is met (no practical max).
        from vendor_intel.funnel.channel_partner_types import channel_partner_mode_enabled

        if (
            channel_partner_mode_enabled(query_context, settings=settings)
            and export_min > 0
            and len(final) < export_min
        ):
            try:
                from vendor_intel.discovery.channel_partner_backfill import (
                    backfill_channel_partners_until_min,
                )

                market_label = str(
                    (scope or {}).get("market")
                    or query_context.get("industry")
                    or query
                )
                backfilled = await backfill_channel_partners_until_min(
                    market=market_label,
                    query_context=query_context,
                    settings=settings,
                    final_rows=final,
                    rejected_rows=export_rejected,
                    min_rows=export_min,
                    scope=scope if isinstance(scope, dict) else None,
                )
                if backfilled:
                    final.extend(backfilled)
                    classified.extend(backfilled)
                    print(
                        f"  [pipeline] channel backfill added {len(backfilled)} → "
                        f"{len(final)} export rows (target {export_min})",
                        flush=True,
                    )
            except Exception as exc:
                print(
                    f"  [pipeline] channel backfill failed: {type(exc).__name__}: {exc}",
                    flush=True,
                )

    if getattr(settings, "pipeline_shuffle_export", False):
        # Stable pseudo-random order (no quality clustering / seeds-first tell).
        import hashlib

        final.sort(
            key=lambda x: hashlib.md5(
                str(x.get("domain") or x.get("website") or x.get("company") or "").encode()
            ).hexdigest()
        )
    else:
        final.sort(
            key=lambda x: float(x.get("quality_score") or x.get("confidence") or 0), reverse=True
        )

    from vendor_intel.pipeline.quality_export import dedupe_export_rows

    before_dedupe = len(final)
    final = dedupe_export_rows(final)
    if before_dedupe != len(final):
        print(
            f"  [pipeline] deduped export rows: {before_dedupe} -> {len(final)}",
            flush=True,
        )

    # Strict geography: drop companies clearly headquartered outside the target region
    # (seeds exempt; companies with no geo signal are kept). Disabled via PIPELINE_STRICT_GEO
    # / --no-geo-filter to keep global players that operate in the target geography.
    if not recall and getattr(settings, "pipeline_strict_geo", True):
        from vendor_intel.pipeline.geo_filter import geo_mismatch_reason, resolve_target_countries

        if resolve_target_countries(country):
            kept: list[dict[str, Any]] = []
            geo_dropped: list[dict[str, Any]] = []
            for r in final:
                reason = geo_mismatch_reason(r, r.get("signals"), country)
                if reason:
                    r["export_reject"] = reason
                    geo_dropped.append(r)
                else:
                    kept.append(r)
            if geo_dropped:
                final = kept
                export_rejected = list(export_rejected) + geo_dropped
                from collections import Counter

                by_country = Counter(
                    str(r.get("export_reject") or "").split(":", 1)[-1] for r in geo_dropped
                )
                print(
                    f"  [pipeline] strict geo ({country}) dropped {len(geo_dropped)}: "
                    + ", ".join(f"{k}={v}" for k, v in by_country.most_common(5)),
                    flush=True,
                )

    # Confirmation pass: only keep companies the system actually substantiated in the
    # main list. Named/known companies it could NOT confirm are reported separately with
    # a reason (no website found, site unreachable, or relevance not confirmed).
    unverified_companies: list[dict[str, Any]] = []
    if not recall:
        from vendor_intel.utils.domains import normalize_name

        def _vdom(v: dict[str, Any]) -> str:
            d = str(v.get("domain") or v.get("website") or "").strip().lower()
            return d.removeprefix("https://").removeprefix("http://").removeprefix("www.").split("/")[0]

        confirmed: list[dict[str, Any]] = []
        for v in final:
            ok, reason = _verification_status(v)
            if ok:
                confirmed.append(v)
            else:
                unverified_companies.append(
                    {
                        "company": v.get("company") or v.get("brand") or "",
                        "domain": _vdom(v),
                        "role": v.get("role") or "",
                        "reason": reason,
                    }
                )
        final = confirmed

        seed_name_set = {normalize_name(n).lower() for n in user_seed_names}
        final_doms = {_vdom(v) for v in final}
        listed = {(u["company"].lower(), u["domain"]) for u in unverified_companies}
        for r in export_rejected:
            nm = str(r.get("company") or "")
            if normalize_name(nm).lower() in seed_name_set and _vdom(r) not in final_doms:
                key = (nm.lower(), _vdom(r))
                if key not in listed:
                    unverified_companies.append(
                        {
                            "company": nm,
                            "domain": _vdom(r),
                            "role": r.get("role") or "",
                            "reason": _friendly_reject(
                                str(r.get("export_reject") or r.get("reject_reason") or "")
                            ),
                        }
                    )
                    listed.add(key)
        for nm in unresolved_seed_names:
            unverified_companies.append(
                {
                    "company": nm,
                    "domain": "",
                    "role": "",
                    "reason": "could not find an official website for this company",
                }
            )
        if unverified_companies:
            print(
                f"  [pipeline] {len(unverified_companies)} named/known companies could NOT be "
                f"confirmed — listed separately with reasons",
                flush=True,
            )

        if getattr(settings, "pipeline_post_audit_enabled", False) and final:
            try:
                from vendor_intel.pipeline.post_classification_audit import audit_final_companies

                kept, removed, reviewed = await audit_final_companies(final, query_context, scope, settings)
                if removed or reviewed:
                    final = kept
                    export_rejected = list(export_rejected) + removed
                    unverified_companies = unverified_companies + reviewed
            except Exception as exc:
                print(f"  [pipeline] post-classification audit skipped: {exc}", flush=True)

    elapsed = time.perf_counter() - t0
    meter = get_meter()
    llm_summary = meter.summary(
        provider=str(getattr(settings, "llm_provider", "opencode")),
        model=str(getattr(settings, "opencode_model", "")),
    )

    print(
        f"  [pipeline] {len(final)} exported (quality) of {len(classified)} classified "
        f"— no row padding",
        flush=True,
    )
    if not recall and len(final) < max(8, export_min or 0):
        print(
            f"  [pipeline] Thin result set ({len(final)} rows): discovery found limited "
            f"vetted companies; not inflating with low-quality fillers",
            flush=True,
        )
    if not recall and export_rejected:
        print(f"  [pipeline] {len(export_rejected)} rows dropped by quality_export", flush=True)
    print(f"  [pipeline] Total time: {elapsed / 60:.1f} min ({elapsed:.0f}s)", flush=True)
    print(
        f"  [pipeline] LLM calls: {llm_summary['llm_calls_total']} "
        f"(phase1={llm_summary['phase1_compile_calls']}, classify={llm_summary['classify_calls']}) "
        f"est. ${llm_summary['estimated_cost_usd']}",
        flush=True,
    )

    return {
        "query": query,
        "query_context": query_context,
        "pipeline_profile": profile,
        "recall_mode": recall,
        "use_ssc": use_ssc,
        "scope": scope,
        "completed_at": datetime.now(timezone.utc).isoformat(),
        "elapsed_seconds": round(elapsed, 1),
        "elapsed_minutes": round(elapsed / 60, 2),
        "llm_usage": llm_summary,
        "phase1_plan_path": plan_path,
        "phase2_path": (p2 or {}).get("output_path"),
        "phase2_company_count": len((p2 or {}).get("companies") or []) or len(cleaned),
        "after_entity_gate": len(cleaned),
        "enriched_count": len(enriched),
        "classified_count": len(classified),
        "export_rejected_count": len(export_rejected) if not recall else 0,
        "relevant_companies": final,
        "all_classified": classified,
        "export_rejected": export_rejected if not recall else [],
        "unverified_companies": unverified_companies,
    }


def _alpha_name(r: dict) -> str:
    from vendor_intel.utils.domains import fix_mojibake

    return fix_mojibake(str(r.get("company") or r.get("brand") or "")).strip().lower()


def _alpha_sort_sections(grouped: list) -> list:
    """Order companies A->Z within each section (sections keep their given order)."""
    return [(name, sorted(rows, key=_alpha_name)) for name, rows in grouped]


def save_pipeline_csv(result: dict[str, Any], path: str) -> str:
    """CSV with sources, website, quality score."""
    import csv
    from pathlib import Path

    ctx = result.get("query_context") or {}
    industry = str(ctx.get("industry") or result.get("query") or "")
    scope = ctx.get("scope") if isinstance(ctx.get("scope"), dict) else result.get("scope")
    country = str(ctx.get("country") or "global")
    rows = [
        r for r in (result.get("relevant_companies") or []) if r.get("is_relevant")
    ]
    from vendor_intel.pipeline.quality_export import dedupe_export_rows

    rows = dedupe_export_rows(rows)

    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)

    from vendor_intel.pipeline.sections import (
        build_section_taxonomy,
        group_into_sections,
        main_product_label,
    )

    _scope = scope if isinstance(scope, dict) else None
    main_product = main_product_label(ctx, _scope)
    custom_sections = [
        str(s).strip() for s in (ctx.get("sections") or []) if str(s).strip()
    ]
    if custom_sections:
        taxonomy = custom_sections
        grouped = group_into_sections(rows, taxonomy, main_product, custom=True)
    else:
        taxonomy = build_section_taxonomy(main_product)
        grouped = group_into_sections(rows, taxonomy, main_product)
    grouped = _alpha_sort_sections(grouped)

    _HEADER = [
        "#",
        "Company",
        "Brand",
        "Parent_or_Independent",
        "Website",
        "Industry",
        "Role",
        "Functionality",
        "Key_Products",
        "Geography",
        "Confidence",
        "Quality_Score",
        "Is_Relevant",
        "Data_Sources",
        "Summary",
    ]

    def _row_cells(i: int, r: dict) -> list:
        from vendor_intel.pipeline.csv_fields import (
            company_summary_cell,
            format_csv_functionality,
            sanitize_key_products,
            truncate_key_products,
        )

        from vendor_intel.utils.domains import fix_mojibake

        kp_raw = r.get("key_products", "") or r.get("product_system_types", "")
        company = fix_mojibake(str(r.get("company", "")))
        brand = fix_mojibake(str(r.get("brand", "") or "")) or company
        kp_raw = sanitize_key_products(str(kp_raw or ""), company=company, brand=brand)
        key_products = truncate_key_products(kp_raw, max_items=3, max_len=48)
        role = str(r.get("role") or "Other")
        role_desc = str(r.get("role_description", "") or r.get("role_detail", ""))
        return [
            i,
            company,
            brand,
            r.get("parent", "") or r.get("parent_or_independent", ""),
            r.get("website", "") or r.get("domain", ""),
            industry,
            role,
            format_csv_functionality(
                role,
                role_desc,
                kp_raw,
                market=industry,
                company_function=str(r.get("company_function") or ""),
                scope=_scope,
            ),
            key_products,
            country,
            round(float(r.get("confidence") or 0), 3),
            round(float(r.get("quality_score") or r.get("confidence") or 0), 3),
            "yes" if r.get("is_relevant") else "no",
            r.get("data_sources", ""),
            company_summary_cell(
                str(r.get("company_summary") or ""),
                company=str(r.get("company") or ""),
                role=role,
                role_description=role_desc,
                key_products=kp_raw,
                max_sentences=4,
                max_len=460,
            ),
        ]

    unverified = sorted(
        result.get("unverified_companies") or [],
        key=lambda u: str(u.get("company") or "").lower(),
    )

    def _write_csv(target: Path) -> None:
        from vendor_intel.utils.domains import fix_mojibake

        with target.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(_HEADER)
            for sect_idx, (section_name, section_rows) in enumerate(grouped):
                if sect_idx > 0:
                    w.writerow([])  # blank spacer row between sections
                w.writerow([f"=== {section_name} ({len(section_rows)}) ==="])
                for i, r in enumerate(section_rows, 1):
                    w.writerow(_row_cells(i, r))
            if unverified:
                w.writerow([])
                w.writerow([f"=== Not Verified - Could Not Confirm ({len(unverified)}) ==="])
                for i, u in enumerate(unverified, 1):
                    w.writerow(
                        [
                            i,
                            fix_mojibake(str(u.get("company") or "")),
                            "", "",
                            u.get("domain", ""),
                            industry,
                            "(unverified)",
                            "", "", country, "", "", "no",
                            "",
                            str(u.get("reason") or ""),
                        ]
                    )

    try:
        _write_csv(out)
    except PermissionError:
        alt = out.with_name(
            f"{out.stem}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}{out.suffix}"
        )
        _write_csv(alt)
        print(
            f"  [pipeline] CSV locked ({out.name}) — saved: {alt.resolve()}",
            flush=True,
        )
        return str(alt.resolve())

    print(f"  [pipeline] CSV saved: {out.resolve()}", flush=True)
    return str(out.resolve())


def save_pipeline_docx(result: dict[str, Any], path: str) -> str:
    """CEO Word doc: market heading → section sub-headings → numbered companies
    with each company's functionality in brackets. Mirrors the CSV sectioning."""
    from pathlib import Path

    from docx import Document
    from docx.shared import Pt

    from vendor_intel.pipeline.csv_fields import format_csv_functionality
    from vendor_intel.pipeline.quality_export import dedupe_export_rows
    from vendor_intel.pipeline.sections import (
        build_section_taxonomy,
        group_into_sections,
        main_product_label,
    )
    from vendor_intel.utils.domains import fix_mojibake

    ctx = result.get("query_context") or {}
    industry = str(ctx.get("industry") or result.get("query") or "")
    scope = ctx.get("scope") if isinstance(ctx.get("scope"), dict) else result.get("scope")
    _scope = scope if isinstance(scope, dict) else None
    country = str(ctx.get("country") or "global")
    rows = [r for r in (result.get("relevant_companies") or []) if r.get("is_relevant")]
    rows = dedupe_export_rows(rows)

    main_product = main_product_label(ctx, _scope)
    custom_sections = [str(s).strip() for s in (ctx.get("sections") or []) if str(s).strip()]
    if custom_sections:
        grouped = group_into_sections(rows, custom_sections, main_product, custom=True)
    else:
        grouped = group_into_sections(rows, build_section_taxonomy(main_product), main_product)
    grouped = _alpha_sort_sections(grouped)

    doc = Document()
    doc.add_heading(industry or main_product or "Market Landscape", level=0)
    meta = doc.add_paragraph()
    meta.add_run(f"Geography: {country}     Companies: {len(rows)}").italic = True

    for section_name, section_rows in grouped:
        doc.add_heading(f"{section_name} ({len(section_rows)})", level=1)
        for i, r in enumerate(section_rows, 1):
            company = fix_mojibake(str(r.get("company") or r.get("brand") or "")).strip()
            role = str(r.get("role") or "Other")
            kp_raw = r.get("key_products", "") or r.get("product_system_types", "")
            func = format_csv_functionality(
                role,
                str(r.get("role_description", "") or r.get("role_detail", "")),
                kp_raw,
                market=industry,
                company_function=str(r.get("company_function") or ""),
                scope=_scope,
            )
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(18)
            p.add_run(f"{i}. ")
            p.add_run(company).bold = True
            if func:
                p.add_run(f" ({func})")

    unverified = sorted(
        result.get("unverified_companies") or [],
        key=lambda u: str(u.get("company") or "").lower(),
    )
    if unverified:
        doc.add_heading(f"Not Verified - Could Not Confirm ({len(unverified)})", level=1)
        note = doc.add_paragraph()
        note.add_run(
            "Named/known companies the system could not substantiate. Reason given per company."
        ).italic = True
        for i, u in enumerate(unverified, 1):
            company = fix_mojibake(str(u.get("company") or "")).strip()
            reason = str(u.get("reason") or "")
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(18)
            p.add_run(f"{i}. ")
            p.add_run(company).bold = True
            if reason:
                p.add_run(f" - {reason}")

    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(str(out))
    except PermissionError:
        alt = out.with_name(
            f"{out.stem}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}{out.suffix}"
        )
        doc.save(str(alt))
        print(f"  [pipeline] DOCX locked ({out.name}) — saved: {alt.resolve()}", flush=True)
        return str(alt.resolve())
    print(f"  [pipeline] DOCX saved: {out.resolve()}", flush=True)
    return str(out.resolve())


def run_pipeline_sync(
    query_context: dict[str, Any],
    settings: Settings | None = None,
    **kwargs: Any,
) -> dict[str, Any]:
    return asyncio.run(run_pipeline(query_context, settings, **kwargs))
