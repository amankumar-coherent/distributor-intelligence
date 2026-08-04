#!/usr/bin/env python3
"""Gap-fill (crawl + LLM merge-only) then export Founded+HQ filtered Excel.

Fills missing Cities/Regions, employees, contacts, etc. via website crawl and
search snippets. Never overwrites existing Founded / Headquarters (merge-only).

Usage:
  .venv\\Scripts\\python.exe export_founded_hq_v2.py cyber_physical_systems_security
  .venv\\Scripts\\python.exe export_founded_hq_v2.py cyber_physical_systems_security --force
  .venv\\Scripts\\python.exe export_founded_hq_v2.py cyber_physical_systems_security --export-only
  .venv\\Scripts\\python.exe export_founded_hq_v2.py --all --export-only
"""
from __future__ import annotations

import argparse
import asyncio
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx  # noqa: E402

SKIP_JSON = {
    "audit_cache.json",
    "gap_fill_cache.json",
    "contact_fill_cache.json",
    "discovery_new_companies.json",
}

# Crawl + search + contact pages; merge-only (see gap_fill/gaps.py merge_extracted).
ENRICH_SOURCES = (
    "backfill,linkedin,owler,crunchbase,wikipedia,sec,locations,contact_page,leadership"
)


def _filled(v: object) -> bool:
    s = str(v or "").strip()
    return bool(s) and s.lower() not in {
        "null",
        "none",
        "n/a",
        "not publicly disclosed",
        "not specified",
        "unknown",
        "-",
    }


def has_founded(row: dict) -> bool:
    crawl = row.get("_crawl") or {}
    if _filled(crawl.get("founded_year")) or _filled(row.get("founded_year")):
        return True
    text = str(crawl.get("crawl_text") or "")
    return bool(re.search(r"\bfounded(?:\s+in)?\s+\d{4}\b", text, re.I))


def has_headquarters(row: dict) -> bool:
    crawl = row.get("_crawl") or {}
    return _filled(crawl.get("headquarters")) or _filled(crawl.get("contact_address"))


def _slug(name: str) -> str:
    s = re.sub(r"[^\w\s-]", "", name.lower())
    s = re.sub(r"\s+", "_", s.strip())
    return s[:80] or "market"


def find_main_json(market_dir: Path) -> Path | None:
    if not market_dir.is_dir():
        return None
    candidates = [
        p
        for p in sorted(market_dir.glob("*.json"))
        if p.name not in SKIP_JSON and "market_research" not in p.name
        and not p.name.startswith("_")
    ]
    if candidates:
        return next((p for p in candidates if "global" in p.name), candidates[0])
    # Nested layout: output/<slug>/<slug>/<slug>.json
    nested = market_dir / market_dir.name
    if nested.is_dir():
        nested_cands = [
            p
            for p in sorted(nested.glob("*.json"))
            if p.name not in SKIP_JSON and "market_research" not in p.name
            and not p.name.startswith("_")
        ]
        if nested_cands:
            return next((p for p in nested_cands if "global" in p.name), nested_cands[0])
    return None


async def enrich_market(
    slug: str,
    *,
    force: bool = False,
    concurrent: int = 4,
    push_neon: bool = True,
) -> dict:
    from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_market_from_slug

    print(f"\n=== STEP 1/2: GAP-FILL + CRAWL ({slug}) ===", flush=True)
    print(f"  Sources: {ENRICH_SOURCES}", flush=True)
    print("  Merge-only: existing Founded/HQ are kept; only empty fields filled.", flush=True)

    stats = await gap_fill_market_from_slug(
        slug,
        sources=ENRICH_SOURCES,
        force=force,
        fmcg_xlsx=True,
        also_csv=True,
        push_neon=push_neon,
        concurrent=concurrent,
    )
    print(f"\n  Gap-fill stats: {json.dumps(stats, indent=2)}", flush=True)
    return stats


def export_one(market_dir: Path, *, suffix: str = "founded_and_hq_v2") -> dict:
    json_path = find_main_json(market_dir)
    if not json_path:
        return {"market_dir": str(market_dir), "error": "no export JSON found"}

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]

    # In channel mode only ANALYST seeds are guaranteed (auto-enumerated market
    # players are gated normally); landscape mode keeps all seeds guaranteed.
    _channel_mode = False
    try:
        from vendor_intel.funnel.channel_partner_types import channel_partner_mode_enabled

        _channel_mode = channel_partner_mode_enabled(data.get("query_context") or {})
    except Exception:
        pass
    _keep_flag = "is_analyst_seed" if _channel_mode else "is_seed"
    filtered = [
        r for r in rows
        if (has_founded(r) and has_headquarters(r)) or r.get(_keep_flag)
    ]

    # Channel Partner Intelligence production gate: non-reject confidence tier.
    # Review-tier rows stay in the JSON but analyst seeds are always exportable.
    try:
        from vendor_intel.funnel.channel_partner_types import is_production_exportable

        if _channel_mode:
            filtered = [r for r in filtered if is_production_exportable(r)]
    except Exception:
        pass

    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    country = str(ctx.get("country") or "global")

    out_name = f"{_slug(market)}_{suffix}.xlsx"
    out_path = market_dir / out_name

    result = {
        "query": market,
        "query_context": {"industry": market, "country": country},
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + Headquarters both filled (after gap-fill enrich)",
    }
    export_fmcg_xlsx(result, out_path, market_type="fmcg")

    return {
        "market": market,
        "total": len(rows),
        "filtered": len(filtered),
        "file": str(out_path.resolve()),
    }


def discover_market_dirs(output_root: Path) -> list[Path]:
    dirs: list[Path] = []
    seen: set[str] = set()
    for json_path in sorted(output_root.rglob("*.json")):
        if json_path.name in SKIP_JSON or "market_research" in json_path.name:
            continue
        try:
            data = json.loads(json_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(data.get("relevant_companies"), list):
            continue
        slug = json_path.parent.name
        if slug not in seen:
            seen.add(slug)
            dirs.append(json_path.parent)
    return dirs


MARKET_DESCRIPTIONS = {
    "cyber_physical_systems_security": (
        "cybersecurity for cyber-physical systems (CPS), OT/ICS security, SCADA protection, "
        "industrial control system security, critical infrastructure protection"
    ),
}


async def verify_market(slug: str, *, concurrent: int = 3, force: bool = False) -> dict:
    from vendor_intel.enrichment.company_data_verify import verify_companies

    market_dir = ROOT / "output" / slug
    json_path = find_main_json(market_dir)
    if not json_path:
        return {"skipped": True, "reason": "no json"}

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    filtered = [r for r in rows if has_founded(r) and has_headquarters(r)]
    ctx = data.get("query_context") or {}
    market_query = str(ctx.get("industry") or data.get("query") or slug.replace("_", " "))
    market_desc = MARKET_DESCRIPTIONS.get(slug, market_query)

    print(f"\n=== STEP 3: OPENAI VERIFY ({slug}) — {len(filtered)} companies ===", flush=True)
    print("  Checks: market fit, brand/products, HQ/founded vs web evidence", flush=True)
    return await verify_companies(
        filtered,
        market_query=market_query,
        market_desc=market_desc,
        output_dir=market_dir,
        concurrent=concurrent,
        use_web_search=True,
        force=force,
    )


async def run_single(
    slug: str,
    *,
    suffix: str,
    export_only: bool,
    force: bool,
    concurrent: int,
    push_neon: bool,
    verify: bool,
    verify_force: bool,
) -> dict:
    enrich_stats = None
    if not export_only:
        enrich_stats = await enrich_market(
            slug,
            force=force,
            concurrent=concurrent,
            push_neon=push_neon,
        )
        if enrich_stats.get("error"):
            return {"slug": slug, "enrich": enrich_stats, "export": None}

    print(f"\n=== STEP 2/2: EXPORT Founded+HQ Excel ({slug}) ===", flush=True)
    export_stats = export_one(ROOT / "output" / slug, suffix=suffix)
    verify_stats = None
    if verify:
        verify_stats = await verify_market(slug, concurrent=min(3, concurrent), force=verify_force)
    return {"slug": slug, "enrich": enrich_stats, "export": export_stats, "verify": verify_stats}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Gap-fill enrich (crawl + LLM merge-only) then export Founded+HQ Excel v2",
    )
    parser.add_argument(
        "slug",
        nargs="?",
        default="cyber_physical_systems_security",
        help="Market folder under output/ (default: cyber_physical_systems_security)",
    )
    parser.add_argument("--all", action="store_true", help="Export every market (export-only unless --enrich)")
    parser.add_argument("--export-only", action="store_true", help="Skip gap-fill; only write filtered Excel")
    parser.add_argument(
        "--enrich",
        action="store_true",
        help="Run gap-fill before export (default for single market; off for --all)",
    )
    parser.add_argument("--force", action="store_true", help="Re-run gap-fill even if cache says done")
    parser.add_argument("--no-neon", action="store_true", help="Skip merge-push to Neon after gap-fill")
    parser.add_argument("--concurrent", type=int, default=4, help="Parallel companies during gap-fill (default: 4)")
    parser.add_argument("--suffix", default="founded_and_hq_v2", help="Output filename suffix")
    parser.add_argument(
        "--verify",
        action="store_true",
        help="After export: OpenAI verify market fit + field accuracy (uses OPENAI_API_KEY)",
    )
    parser.add_argument(
        "--verify-only",
        action="store_true",
        help="Skip gap-fill/export; run OpenAI verification on Founded+HQ companies only",
    )
    parser.add_argument("--verify-force", action="store_true", help="Re-run verification ignoring verify_cache")
    args = parser.parse_args()

    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.utils.output_filter import install_stderr_filter

    install_stderr_filter()
    apply_env_overrides()

    output_root = ROOT / "output"
    push_neon = not args.no_neon

    try:
        if args.all:
            do_enrich = args.enrich
            results = []
            for market_dir in discover_market_dirs(output_root):
                slug = market_dir.name
                if do_enrich:
                    results.append(
                        asyncio.run(
                            run_single(
                                slug,
                                suffix=args.suffix,
                                export_only=False,
                                force=args.force,
                                concurrent=args.concurrent,
                                push_neon=push_neon,
                                verify=False,
                                verify_force=False,
                            )
                        )
                    )
                else:
                    results.append({"slug": slug, "export": export_one(market_dir, suffix=args.suffix)})
            print(json.dumps(results, indent=2))
        else:
            if args.verify_only:
                result = {
                    "slug": args.slug,
                    "verify": asyncio.run(
                        verify_market(
                            args.slug,
                            concurrent=min(3, args.concurrent),
                            force=args.verify_force,
                        )
                    ),
                }
            else:
                do_enrich = not args.export_only
                if do_enrich:
                    result = asyncio.run(
                        run_single(
                            args.slug,
                            suffix=args.suffix,
                            export_only=False,
                            force=args.force,
                            concurrent=args.concurrent,
                            push_neon=push_neon,
                            verify=args.verify,
                            verify_force=args.verify_force,
                        )
                    )
                else:
                    result = {
                        "slug": args.slug,
                        "export": export_one(output_root / args.slug, suffix=args.suffix),
                    }
                    if args.verify:
                        result["verify"] = asyncio.run(
                            verify_market(
                                args.slug,
                                concurrent=min(3, args.concurrent),
                                force=args.verify_force,
                            )
                        )
            print(json.dumps(result, indent=2))
    finally:
        try:
            from vendor_intel.clients.ddg_worker_pool import shutdown_ddg_pool

            shutdown_ddg_pool(wait=True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
