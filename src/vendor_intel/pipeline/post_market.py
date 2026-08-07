"""Post-export pipeline: gap-fill (Google AI → DDGS/SearXNG → … → OpenAI residual) → LinkedIn verify → …

Called automatically from run_query.py after a market export (unless --no-post-enrich).

Enrichment phases for Channel Partner Intelligence:
  - **Google AI scraper FIRST** (primary) — fill Founded, HQ, cities, turnover ranges,
    contact person, core categories, specialty, key brands via scrape + LLM
  - Then LinkedIn MCP, Owler, Wikidata, Wikipedia, web search (backfill) for leftovers
  - Key Brands Represented uses the same key_brands_only resolve as Founded+HQ enrich scripts
  - Excel columns stay the same (Business Type holds Channel Type; Key Brands column unchanged)
"""
from __future__ import annotations

import asyncio
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from vendor_intel.config import _project_root

ROOT = _project_root()
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Any market: Google AI → DDGS/SearXNG (backfill) → other free sources → OpenAI residual last.
GAP_FILL_SOURCES = (
    "google_ai,backfill,wikipedia,wikidata,owler,linkedin,crunchbase,sec,"
    "locations,contact_page,leadership,openai_residual"
)


def _json_skip_names() -> set[str]:
    return {
        "audit_cache.json",
        "gap_fill_cache.json",
        "contact_fill_cache.json",
        "apollo_contact_cache.json",
        "apollo_contact_progress.json",
        "discovery_new_companies.json",
        "brand_resolve_cache.json",
        "market_relevance_audit_cache.json",
        "post_market_progress.json",
        "_neon_filtered_export.json",
    }


def _pick_export_json(candidates: list[Path]) -> Path | None:
    if not candidates:
        return None
    return next((p for p in candidates if "global" in p.name), candidates[0])


def _find_main_json(market_dir: Path) -> Path | None:
    """Locate the market export JSON (flat or nested slug/slug.json layouts)."""
    skip = _json_skip_names()

    def _ok(p: Path) -> bool:
        return (
            p.name not in skip
            and "market_research" not in p.name
            and not p.name.startswith("_")
        )

    candidates = [p for p in sorted(market_dir.glob("*.json")) if _ok(p)]
    picked = _pick_export_json(candidates)
    if picked:
        return picked

    # Nested layout: output/<slug>/<slug>/<slug>.json (older runs)
    nested = market_dir / market_dir.name
    if nested.is_dir():
        nested_cands = [p for p in sorted(nested.glob("*.json")) if _ok(p)]
        picked = _pick_export_json(nested_cands)
        if picked:
            return picked

    # Last resort: any relevant_companies JSON one level down
    for sub in sorted(market_dir.iterdir()):
        if not sub.is_dir():
            continue
        for p in sorted(sub.glob("*.json")):
            if not _ok(p):
                continue
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                continue
            if isinstance(data.get("relevant_companies"), list):
                return p
    return None


def _refresh_primary_landscape_exports(market_dir: Path, data: dict[str, Any]) -> dict[str, Any]:
    """Rewrite the main landscape Excel/CSV after LinkedIn/Wikidata/etc. fills.

    Phase 4 writes the primary ``{slug}.xlsx`` before post-market. Without this
    refresh, Founded / Headquarters / Key Brands stay empty in the file users open
    even when JSON was updated by LinkedIn MCP.
    """
    from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx
    from vendor_intel.pipeline.orchestrator import save_pipeline_csv

    slug = market_dir.name
    market_type = str(data.get("market_type") or "general")
    xlsx_path = Path(str(data.get("_xlsx_path") or "").strip() or (market_dir / f"{slug}.xlsx"))
    if not xlsx_path.is_absolute():
        xlsx_path = (market_dir / xlsx_path).resolve() if not Path(data.get("_xlsx_path") or "").is_absolute() else xlsx_path
        if not xlsx_path.is_file():
            xlsx_path = market_dir / Path(str(data.get("_xlsx_path") or f"{slug}.xlsx")).name
    # Prefer existing top-level landscape xlsx if present
    if not xlsx_path.is_file():
        top = [
            p for p in market_dir.glob("*.xlsx")
            if "founded_and_hq" not in p.name.lower()
        ]
        if top:
            xlsx_path = top[0]
        else:
            xlsx_path = market_dir / f"{slug}.xlsx"

    csv_path = Path(str(data.get("_csv_path") or "").strip() or xlsx_path.with_suffix(".csv"))
    if not csv_path.is_absolute():
        csv_path = market_dir / csv_path.name
    if not csv_path.parent.is_dir():
        csv_path = market_dir / f"{slug}.csv"

    out_xlsx = export_fmcg_xlsx(data, xlsx_path, market_type=market_type)
    save_pipeline_csv(data, str(csv_path))

    data["_xlsx_path"] = str(Path(out_xlsx).resolve())
    data["_csv_path"] = str(Path(csv_path).resolve())

    return {
        "primary_xlsx": str(Path(out_xlsx).resolve()),
        "primary_csv": str(Path(csv_path).resolve()),
        "market_type": market_type,
    }


def _norm_seed_key(s: str) -> str:
    """Lowercase, drop legal suffixes/punctuation for tolerant seed-name matching."""
    s = str(s or "").lower()
    for suf in (
        " inc", " inc.", " llc", " ltd", " ltd.", " limited", " corp", " corp.",
        " corporation", " co.", " co", " group", " gmbh", " ag", " se", " s.a.",
        " plc", " pvt", " private",
    ):
        if s.endswith(suf):
            s = s[: -len(suf)]
    return "".join(ch for ch in s if ch.isalnum())


def _stamp_analyst_seeds(data: dict) -> int:
    """Mark rows as analyst seeds by matching the persisted query_context seed list.

    query_context['seed_companies'] holds ONLY analyst seed names (auto-enumerated
    market players are never added there), and 'seed_domains' holds pinned domains.
    Matching on either the normalized name or the domain guarantees the analyst seed
    bypass survives even if the discovery stage didn't stamp is_analyst_seed.
    Returns the number of rows newly stamped. Idempotent.
    """
    qc = data.get("query_context") or {}
    seed_names = {
        _norm_seed_key(n) for n in (qc.get("seed_companies") or []) if str(n).strip()
    }
    seed_doms = {
        str(d).lower().removeprefix("www.").strip()
        for d in (qc.get("seed_domains") or {}).values()
        if str(d).strip()
    }
    if not seed_names and not seed_doms:
        return 0
    stamped = 0
    for r in data.get("relevant_companies") or []:
        if r.get("is_analyst_seed"):
            continue
        dom = str(r.get("domain") or "").lower().removeprefix("www.").strip()
        nm = _norm_seed_key(r.get("company") or "")
        if (dom and dom in seed_doms) or (nm and nm in seed_names):
            r["is_analyst_seed"] = True
            stamped += 1
    return stamped


def _has_founded(row: dict) -> bool:
    v = str((row.get("_crawl") or {}).get("founded_year") or row.get("founded_year") or "").strip()
    return bool(v) and v.lower() not in {"n/a", "unknown", "-", ""}


def _has_headquarters(row: dict) -> bool:
    crawl = row.get("_crawl") or {}
    for key in ("headquarters", "contact_address"):
        v = str(crawl.get(key) or "").strip()
        if len(v) > 2:
            return True
    return False


def _progress_path(market_dir: Path) -> Path:
    return market_dir / "post_market_progress.json"


def _load_progress(market_dir: Path) -> dict[str, Any]:
    path = _progress_path(market_dir)
    if not path.is_file():
        return {"steps_done": [], "updated_at": None}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"steps_done": [], "updated_at": None}
        data.setdefault("steps_done", [])
        return data
    except Exception:
        return {"steps_done": [], "updated_at": None}


def _save_progress(market_dir: Path, progress: dict[str, Any]) -> None:
    progress["updated_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    _progress_path(market_dir).write_text(
        json.dumps(progress, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def _mark_step_done(market_dir: Path, progress: dict[str, Any], step: str) -> None:
    done = list(progress.get("steps_done") or [])
    if step not in done:
        done.append(step)
    progress["steps_done"] = done
    progress["current_step"] = None
    progress["last_completed_step"] = step
    _save_progress(market_dir, progress)


def _begin_step(market_dir: Path, progress: dict[str, Any], step: str) -> None:
    """Checkpoint that this step is in progress (resume can see where we stopped)."""
    progress["current_step"] = step
    _save_progress(market_dir, progress)


def _step_done(progress: dict[str, Any], step: str, *, resume: bool) -> bool:
    return resume and step in (progress.get("steps_done") or [])


def _sync_neon_filtered(market_dir: Path, data: dict, filtered: list[dict]) -> Path:
    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    out = {
        "query": market,
        "query_context": ctx,
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ + market-audit passed — post-market export for Neon",
    }
    path = market_dir / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


async def run_post_market_pipeline(
    market_dir: Path,
    *,
    gap_fill: bool = True,
    linkedin_firmographic_verify: bool = True,
    wikidata_firmographic: bool = True,
    wikipedia_firmographic: bool = True,
    apollo_names: bool = True,
    resolve_partials: bool = True,
    brand_domain_resolve: bool = True,
    market_relevance_audit: bool = True,
    channel_evidence_verify: bool = True,
    export_excel: bool = True,
    push_neon: bool = True,
    gap_force: bool = False,
    gap_concurrent: int = 4,
    linkedin_verify_concurrent: int = 1,
    apollo_delay: float = 2.0,
    name_resolve_concurrent: int = 3,
    resume: bool = True,
) -> dict[str, Any]:
    """Run all post-export enrichment steps for one market folder.

    resume=True (default): skip finished steps and already-processed companies.
    Ctrl+C mid-step is safe — each step writes post_market_progress.json between
    steps; LinkedIn/gap-fill/name/brand/audit also checkpoint per company.
    """
    market_dir = Path(market_dir).resolve()
    json_path = _find_main_json(market_dir)
    if not json_path:
        return {"error": f"no export JSON in {market_dir}", "market_dir": str(market_dir)}

    slug = market_dir.name
    stats: dict[str, Any] = {
        "slug": slug,
        "market_dir": str(market_dir),
        "json": str(json_path),
        "resume": resume,
    }
    t0 = time.perf_counter()
    progress = _load_progress(market_dir)

    print(f"\n{'=' * 60}", flush=True)
    print(f"  POST-MARKET PIPELINE: {slug}", flush=True)
    print(f"  Folder: {market_dir}", flush=True)
    if resume and progress.get("steps_done"):
        print(f"  Resume: skipping done steps {progress.get('steps_done')}", flush=True)
    if resume and progress.get("current_step"):
        print(f"  Resume: last in-progress step was {progress.get('current_step')}", flush=True)
    print(f"{'=' * 60}", flush=True)

    def _persist(data_obj: dict) -> None:
        json_path.write_text(json.dumps(data_obj, indent=2, ensure_ascii=False), encoding="utf-8")

    # ── Step 1: Gap-fill (Founded, HQ, leadership, contact pages) ──────────
    if gap_fill and not _step_done(progress, "gap_fill", resume=resume):
        print("\n  [1/8] Gap-fill (Google AI → DDGS/SearXNG → … → OpenAI residual)…", flush=True)
        _begin_step(market_dir, progress, "gap_fill")
        from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_market

        data = json.loads(json_path.read_text(encoding="utf-8"))
        gap_stats = await gap_fill_market(
            data,
            json_path=json_path,
            output_dir=market_dir,
            market_type=data.get("market_type") or "fmcg",
            fmcg_xlsx=False,
            also_csv=True,
            push_neon=False,
            concurrent=gap_concurrent,
            sources=GAP_FILL_SOURCES,
            force=gap_force,
        )
        stats["gap_fill"] = gap_stats
        data = json.loads(json_path.read_text(encoding="utf-8"))
        _mark_step_done(market_dir, progress, "gap_fill")
        print("  [checkpoint] gap_fill done → post_market_progress.json", flush=True)
    else:
        if gap_fill:
            print("\n  [1/8] Gap-fill — SKIPPED (resume: already done)", flush=True)
            stats["gap_fill"] = {"skipped": True, "reason": "resume_step_done"}
        data = json.loads(json_path.read_text(encoding="utf-8"))

    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]

    # ── Step 2: LinkedIn MCP verify Founded + HQ ───────────────────────────
    if linkedin_firmographic_verify and not _step_done(
        progress, "linkedin_firmographic_verify", resume=resume
    ):
        from vendor_intel.integrations.linkedin_mcp import linkedin_mcp_available, linkedin_mcp_delay

        if linkedin_mcp_available():
            print(
                f"\n  [2/8] LinkedIn MCP — fill/verify Founded + Headquarters "
                f"(also cities/about/key brands when present; "
                f"delay ~{linkedin_mcp_delay():.1f}s/call)...",
                flush=True,
            )
            _begin_step(market_dir, progress, "linkedin_firmographic_verify")
            from vendor_intel.enrichment.linkedin_firmographic_verify import verify_firmographics_for_market

            li_verify_stats = await verify_firmographics_for_market(
                rows,
                concurrent=linkedin_verify_concurrent,
                verify_all=True,
                persist_fn=lambda: _persist(data),
            )
            stats["linkedin_firmographic_verify"] = li_verify_stats
            _persist(data)
            # Always mark done after a full pass; per-company linkedin_mcp_attempted
            # still skips already-processed rows if Ctrl+C mid-step.
            _mark_step_done(market_dir, progress, "linkedin_firmographic_verify")
            print("  [checkpoint] linkedin_firmographic_verify done", flush=True)
            rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
        else:
            print("\n  [2/8] LinkedIn verify skipped — MCP not available", flush=True)
            stats["linkedin_firmographic_verify"] = {"skipped": True, "reason": "mcp_unavailable"}
    elif linkedin_firmographic_verify:
        print("\n  [2/8] LinkedIn verify — SKIPPED (resume: already done)", flush=True)
        stats["linkedin_firmographic_verify"] = {"skipped": True, "reason": "resume_step_done"}
    else:
        print("\n  [2/8] LinkedIn verify — SKIPPED (disabled)", flush=True)
        stats["linkedin_firmographic_verify"] = {"skipped": True, "reason": "disabled"}
        if not _step_done(progress, "linkedin_firmographic_verify", resume=False):
            _mark_step_done(market_dir, progress, "linkedin_firmographic_verify")

    # ── Step 2b: Wikidata fallback — fill Founded + HQ LinkedIn couldn't cover ─
    if wikidata_firmographic and not _step_done(
        progress, "wikidata_firmographic", resume=resume
    ):
        print(
            "\n  [2b/8] Wikidata fallback — fill Founded + Headquarters "
            "(free; recovers brands LinkedIn omits, e.g. Nokia/Sony/LG)...",
            flush=True,
        )
        _begin_step(market_dir, progress, "wikidata_firmographic")
        from vendor_intel.enrichment.wikidata_firmographic_fill import (
            fill_firmographics_via_wikidata_for_market,
        )

        wd_stats = await fill_firmographics_via_wikidata_for_market(
            rows,
            persist_fn=lambda: _persist(data),
        )
        stats["wikidata_firmographic"] = wd_stats
        _persist(data)
        _mark_step_done(market_dir, progress, "wikidata_firmographic")
        print("  [checkpoint] wikidata_firmographic done", flush=True)
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    elif wikidata_firmographic:
        print("\n  [2b/8] Wikidata fallback — SKIPPED (resume: already done)", flush=True)
        stats["wikidata_firmographic"] = {"skipped": True, "reason": "resume_step_done"}
    else:
        print("\n  [2b/8] Wikidata fallback — SKIPPED (disabled)", flush=True)
        stats["wikidata_firmographic"] = {"skipped": True, "reason": "disabled"}
        if not _step_done(progress, "wikidata_firmographic", resume=False):
            _mark_step_done(market_dir, progress, "wikidata_firmographic")

    # ── Step 2c: Wikipedia infobox fallback — fill Founded + HQ still missing ─
    if wikipedia_firmographic and not _step_done(
        progress, "wikipedia_firmographic", resume=resume
    ):
        print(
            "\n  [2c/8] Wikipedia infobox fallback — fill Founded + Headquarters "
            "(free; recovers brands missing from Wikidata, e.g. Amazfit/Pine64/Fastrack)...",
            flush=True,
        )
        _begin_step(market_dir, progress, "wikipedia_firmographic")
        from vendor_intel.enrichment.wikipedia_firmographic_fill import (
            fill_firmographics_via_wikipedia_for_market,
        )

        wp_stats = await fill_firmographics_via_wikipedia_for_market(
            rows,
            persist_fn=lambda: _persist(data),
        )
        stats["wikipedia_firmographic"] = wp_stats
        _persist(data)
        _mark_step_done(market_dir, progress, "wikipedia_firmographic")
        print("  [checkpoint] wikipedia_firmographic done", flush=True)
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    elif wikipedia_firmographic:
        print("\n  [2c/8] Wikipedia fallback — SKIPPED (resume: already done)", flush=True)
        stats["wikipedia_firmographic"] = {"skipped": True, "reason": "resume_step_done"}
    else:
        print("\n  [2c/8] Wikipedia fallback — SKIPPED (disabled)", flush=True)
        stats["wikipedia_firmographic"] = {"skipped": True, "reason": "disabled"}
        if not _step_done(progress, "wikipedia_firmographic", resume=False):
            _mark_step_done(market_dir, progress, "wikipedia_firmographic")

    # ── Step 3: Apollo free search — contact_person only (no enrich/credits) ─
    if apollo_names and not _step_done(progress, "apollo_names", resume=resume):
        print("\n  [3/8] Apollo search — contact names only (no email, no credits)...", flush=True)
        _begin_step(market_dir, progress, "apollo_names")
        # Import from project root script.
        sys.path.insert(0, str(ROOT))
        from apollo_contact_fill import run as apollo_run

        apollo_stats = await apollo_run(
            slug,
            founded_hq_only=True,
            replace_all=False,
            limit=None,
            concurrent=1,
            enrich_match=False,
            merge_fields=frozenset({"contact_person"}),
            dry_run=False,
            push_neon=False,
            request_delay=apollo_delay,
            resume=True,
            skip_apollo_filled=False,
            companies=None,
            market_dir=market_dir,
        )
        stats["apollo_names"] = apollo_stats
        data = json.loads(json_path.read_text(encoding="utf-8"))
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
        _mark_step_done(market_dir, progress, "apollo_names")
        print("  [checkpoint] apollo_names done", flush=True)
    elif apollo_names:
        print("\n  [3/8] Apollo — SKIPPED (resume: already done)", flush=True)
        stats["apollo_names"] = {"skipped": True, "reason": "resume_step_done"}

    # ── Step 4: Upgrade partial / missing names (leadership + LinkedIn search) ─
    # Allow CLI/env to skip: leave Apollo partial names as-is (no RocketReach/Google waterfall).
    _skip_nr = (os.getenv("SKIP_NAME_RESOLVE") or os.getenv("RESOLVE_PARTIALS") or "").strip().lower()
    if _skip_nr in ("1", "true", "yes", "on") or _skip_nr in ("0", "false", "no", "off"):
        # SKIP_NAME_RESOLVE=true OR RESOLVE_PARTIALS=false
        if (os.getenv("SKIP_NAME_RESOLVE") or "").strip().lower() in ("1", "true", "yes", "on"):
            resolve_partials = False
        if (os.getenv("RESOLVE_PARTIALS") or "").strip().lower() in ("0", "false", "no", "off"):
            resolve_partials = False
    if resolve_partials and not _step_done(progress, "name_resolve", resume=resume):
        print("\n  [4/8] Resolve full names (LinkedIn → leadership → Google → RocketReach)...", flush=True)
        _begin_step(market_dir, progress, "name_resolve")
        from vendor_intel.integrations.linkedin_mcp import linkedin_mcp_available

        if linkedin_mcp_available():
            name_resolve_concurrent = min(name_resolve_concurrent, 1)

        from vendor_intel.enrichment.contact_name_resolve import resolve_contact_names_for_market

        name_stats = await resolve_contact_names_for_market(
            rows,
            concurrent=name_resolve_concurrent,
            persist_fn=lambda: _persist(data),
        )
        stats["name_resolve"] = name_stats
        _persist(data)
        _mark_step_done(market_dir, progress, "name_resolve")
        print("  [checkpoint] name_resolve done", flush=True)
    elif resolve_partials:
        print("\n  [4/8] Name resolve — SKIPPED (resume: already done)", flush=True)
        stats["name_resolve"] = {"skipped": True, "reason": "resume_step_done"}
    else:
        print("\n  [4/8] Name resolve — SKIPPED (Apollo partials kept as-is)", flush=True)
        stats["name_resolve"] = {"skipped": True, "reason": "resolve_partials_off"}
        if not _step_done(progress, "name_resolve", resume=False):
            _mark_step_done(market_dir, progress, "name_resolve")

    # Always reload after early steps so resume mid-file sees latest checkpoint.
    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]

    # ── Step 5: Correct company / brand / domain + Key Brands Represented ─
    if brand_domain_resolve and not _step_done(progress, "brand_resolve", resume=resume):
        print("\n  [5/8] Correct company identity (web + LinkedIn + OpenAI)...", flush=True)
        _begin_step(market_dir, progress, "brand_resolve")
        from vendor_intel.enrichment.brand_domain_resolve import resolve_brand_domains_for_market

        ctx = data.get("query_context") or {}
        market_query = str(ctx.get("industry") or data.get("query") or slug.replace("_", " "))
        fh_rows = [r for r in rows if _has_founded(r) and _has_headquarters(r)]
        brand_stats = await resolve_brand_domains_for_market(
            fh_rows,
            market_query=market_query,
            market_dir=market_dir,
            concurrent=name_resolve_concurrent,
            founded_hq_only=False,
            key_brands_only=False,
        )
        stats["brand_resolve"] = brand_stats
        _persist(data)
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]

        # Same Key Brands fill as Founded+HQ enrich scripts (technical foam, etc.):
        # LinkedIn MCP may already have set key_brands; this fills remaining gaps.
        print(
            "\n  [5b/8] Key Brands Represented (web + LinkedIn evidence, same as before)...",
            flush=True,
        )
        fh_rows = [r for r in rows if _has_founded(r) and _has_headquarters(r)]
        key_brand_stats = await resolve_brand_domains_for_market(
            fh_rows,
            market_query=market_query,
            market_dir=market_dir,
            concurrent=1,
            founded_hq_only=False,
            key_brands_only=True,
            force=False,
        )
        stats["key_brands_resolve"] = key_brand_stats
        _persist(data)
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
        _mark_step_done(market_dir, progress, "brand_resolve")
        print("  [checkpoint] brand_resolve done", flush=True)
    elif brand_domain_resolve:
        print("\n  [5/8] Brand / key brands — SKIPPED (resume: already done)", flush=True)
        stats["brand_resolve"] = {"skipped": True, "reason": "resume_step_done"}

    # ── Step 6: Market relevance audit (web search + OpenAI keep/review/remove) ─
    if market_relevance_audit and not _step_done(progress, "market_audit", resume=resume):
        print("\n  [6/8] Audit market relevance (web search + OpenAI)...", flush=True)
        _begin_step(market_dir, progress, "market_audit")
        from vendor_intel.enrichment.market_relevance_audit import audit_market_relevance_for_rows

        ctx = data.get("query_context") or {}
        scope = data.get("scope") if isinstance(data.get("scope"), dict) else ctx.get("scope")
        fh_rows = [r for r in rows if _has_founded(r) and _has_headquarters(r)]
        audit_stats = await audit_market_relevance_for_rows(
            fh_rows,
            query_context=ctx,
            scope=scope if isinstance(scope, dict) else None,
            market_dir=market_dir,
            concurrent=name_resolve_concurrent,
        )
        stats["market_audit"] = audit_stats
        _persist(data)
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
        _mark_step_done(market_dir, progress, "market_audit")
        print("  [checkpoint] market_audit done", flush=True)
    elif market_relevance_audit:
        print("\n  [6/8] Market audit — SKIPPED (resume: already done)", flush=True)
        stats["market_audit"] = {"skipped": True, "reason": "resume_step_done"}

    # ── Channel Partner Intelligence gates: manufacturers, confidence, evidence ──
    from vendor_intel.funnel.channel_partner_types import (
        apply_channel_gate_to_rows,
        apply_confidence_tier_to_rows,
        channel_partner_mode_enabled,
    )

    # Robust seed-guarantee: re-derive is_analyst_seed from the persisted
    # query_context (analyst seed names + pinned domains) so seeds are never dropped
    # even if the discovery stage failed to stamp the flag. Idempotent.
    _stamped = _stamp_analyst_seeds(data)
    if _stamped:
        print(f"  [seeds] marked {_stamped} analyst seed row(s) as guaranteed", flush=True)

    if channel_partner_mode_enabled(data.get("query_context") or {}):
        # (a) Drop pure manufacturers/brands (strict channel-only).
        gated = apply_channel_gate_to_rows(data.get("relevant_companies") or [])
        if gated:
            print(
                f"\n  [channel-gate] Removed {gated} pure manufacturer/brand rows",
                flush=True,
            )

    # Always-on channel purity (OEM blocklist + heuristics) — discovery stays wide
    try:
        from vendor_intel.intelligence.channel_purity import (
            channel_purity_enabled,
            filter_channel_pure_rows,
        )
        from vendor_intel.pipeline.web_expand import market_family as _mf

        if channel_purity_enabled():
            qctx = data.get("query_context") or {}
            q = str(qctx.get("industry") or data.get("query") or "")
            before = list(data.get("relevant_companies") or [])
            kept, dropped = filter_channel_pure_rows(
                before, family=_mf(q), query=q
            )
            if len(kept) != len(before):
                data["relevant_companies"] = kept
                _persist(data)
                print(
                    f"\n  [channel-purity] kept {len(kept)}/{len(before)} "
                    f"(dropped {len(dropped)} OEM/utility/non-channel)",
                    flush=True,
                )
                (market_dir / "channel_purity_dropped.json").write_text(
                    json.dumps(dropped, indent=2, ensure_ascii=False), encoding="utf-8"
                )
    except Exception as exc:
        print(f"  [channel-purity] skipped: {exc}", flush=True)

    if channel_partner_mode_enabled(data.get("query_context") or {}):
        # (b) Verify channel-role evidence from each company's own website.
        if channel_evidence_verify and not _step_done(
            progress, "channel_evidence", resume=resume
        ):
            print(
                "\n  [6b/8] Verify channel-role evidence (crawl official site: "
                "'Authorized Distributor', 'Certified Integrator', ...)...",
                flush=True,
            )
            _begin_step(market_dir, progress, "channel_evidence")
            from vendor_intel.enrichment.channel_evidence_verify import (
                verify_channel_evidence_for_market,
            )

            ev_rows = [
                r for r in (data.get("relevant_companies") or [])
                if r.get("is_relevant", True)
                and _has_founded(r) and _has_headquarters(r)
            ]
            ev_stats = await verify_channel_evidence_for_market(
                ev_rows, persist_fn=lambda: _persist(data)
            )
            stats["channel_evidence"] = ev_stats
            _persist(data)
            _mark_step_done(market_dir, progress, "channel_evidence")
            print("  [checkpoint] channel_evidence done", flush=True)
        elif channel_evidence_verify:
            print("\n  [6b/8] Channel evidence — SKIPPED (resume: already done)", flush=True)
            stats["channel_evidence"] = {"skipped": True, "reason": "resume_step_done"}

        # (c) Confidence tiers: >=.90 auto, .70-.89 high, .50-.69 review, <.50 reject.
        tiers = apply_confidence_tier_to_rows(data.get("relevant_companies") or [])
        print(
            f"  [confidence] auto(>=.90)={tiers['auto']} high(.70-.89)={tiers['high']} "
            f"review(.50-.69)={tiers['review']} reject(<.50)={tiers['reject']}",
            flush=True,
        )
        _persist(data)
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]

    # ── Step 7: Refresh primary landscape Excel + Founded+HQ filtered Excel ─
    if export_excel and not _step_done(progress, "excel", resume=resume):
        print(
            "\n  [7/8] Export Excel — refresh primary landscape (Founded/HQ/Key Brands) "
            "+ Founded+HQ filtered sheet...",
            flush=True,
        )
        _begin_step(market_dir, progress, "excel")
        sys.path.insert(0, str(ROOT))
        from export_founded_hq_v2 import export_one

        from vendor_intel.enrichment.market_relevance_audit import exportable_after_audit
        from vendor_intel.funnel.channel_partner_types import (
            channel_partner_mode_enabled,
            is_production_exportable,
        )

        # Reload so Excel sees LinkedIn / Wikidata / key-brands writes.
        data = json.loads(json_path.read_text(encoding="utf-8"))
        rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]

        primary_stats = _refresh_primary_landscape_exports(market_dir, data)
        _persist(data)
        print(
            f"  [excel] primary landscape refreshed → {primary_stats.get('primary_xlsx')}",
            flush=True,
        )

        export_stats = export_one(market_dir, suffix="founded_and_hq_v2")
        export_stats["primary_landscape"] = primary_stats
        stats["excel"] = export_stats

        _channel_mode = channel_partner_mode_enabled(data.get("query_context") or {})
        # Only ANALYST seeds are guaranteed in channel mode; auto-enumerated players
        # must clear Founded+HQ + audit like any discovered company. Landscape mode
        # keeps its original "all seeds guaranteed" behavior.
        _keep_flag = "is_analyst_seed" if _channel_mode else "is_seed"
        filtered = [
            r for r in rows
            if r.get(_keep_flag)
            or (
                _has_founded(r) and _has_headquarters(r) and exportable_after_audit(r)
                and (not _channel_mode or is_production_exportable(r))
            )
        ]
        neon_path = _sync_neon_filtered(market_dir, data, filtered)
        stats["neon_json"] = str(neon_path)
        stats["filtered_companies"] = len(filtered)
        _mark_step_done(market_dir, progress, "excel")
        print("  [checkpoint] excel done", flush=True)
    elif export_excel:
        print("\n  [7/8] Excel export — SKIPPED (resume: already done)", flush=True)
        stats["excel"] = {"skipped": True, "reason": "resume_step_done"}
        neon_path = market_dir / "_neon_filtered_export.json"
        if neon_path.is_file():
            stats["neon_json"] = str(neon_path)

    # ── Step 8: Push to Neon ────────────────────────────────────────────────
    if push_neon and not _step_done(progress, "neon_push", resume=resume):
        print("\n  [8/8] Push to Neon...", flush=True)
        _begin_step(market_dir, progress, "neon_push")
        try:
            from vendor_intel.neon_push import push_json_to_neon, should_push_neon

            if should_push_neon():
                neon_path = Path(stats.get("neon_json") or market_dir / "_neon_filtered_export.json")
                if neon_path.is_file():
                    neon_stats = push_json_to_neon(neon_path, merge_companies=False)
                    stats["neon_push"] = neon_stats
                    print(
                        f"  [neon] Pushed {neon_stats.get('vendor_intel_companies', 0)} companies",
                        flush=True,
                    )
                else:
                    neon_stats = push_json_to_neon(json_path, merge_companies=False)
                    stats["neon_push"] = neon_stats
                _mark_step_done(market_dir, progress, "neon_push")
                print("  [checkpoint] neon_push done", flush=True)
            else:
                print("  [neon] Skipped — DATABASE_URL not set", flush=True)
                stats["neon_push"] = {"skipped": True, "reason": "no DATABASE_URL"}
                progress["current_step"] = None
                _save_progress(market_dir, progress)
        except Exception as exc:
            stats["neon_push_error"] = str(exc)
            print(f"  [neon] Push failed: {exc}", flush=True)
    elif push_neon:
        print("\n  [8/8] Neon push — SKIPPED (resume: already done)", flush=True)
        stats["neon_push"] = {"skipped": True, "reason": "resume_step_done"}

    stats["elapsed_sec"] = round(time.perf_counter() - t0, 1)
    stats["completed_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    data["post_market_pipeline"] = stats
    _persist(data)

    print(f"\n  POST-MARKET DONE ({stats['elapsed_sec']}s)", flush=True)
    if stats.get("excel", {}).get("file"):
        print(f"  Excel: {stats['excel']['file']}", flush=True)
    return stats


def run_post_market_pipeline_sync(market_dir: Path, **kwargs: Any) -> dict[str, Any]:
    return asyncio.run(run_post_market_pipeline(market_dir, **kwargs))
