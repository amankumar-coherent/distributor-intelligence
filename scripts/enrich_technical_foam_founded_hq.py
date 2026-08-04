#!/usr/bin/env python3
"""Fill + verify ONLY Technical Foam Founded+HQ companies (~98), then Excel.

Scope: companies with both Founded and Headquarters only (filtered Excel set).
Steps:
  1) Gap-fill — cities/regions, contacts, LinkedIn, locations, etc.
  2) LinkedIn MCP verify — Founded + HQ fill/verify
  3) Key Brands resolve — genuine brands (not company name)
  4) Export founded_and_hq Excel (Desktop + market folder)

Usage:
  .venv\\Scripts\\python.exe scripts/enrich_technical_foam_founded_hq.py
  .venv\\Scripts\\python.exe scripts/enrich_technical_foam_founded_hq.py --limit 5
  .venv\\Scripts\\python.exe scripts/enrich_technical_foam_founded_hq.py --export-only
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

# Load .env so AGENT_REACH_LINKEDIN_MCP / keys are visible before MCP checks
try:
    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
except Exception:
    pass
try:
    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()
except Exception:
    pass

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Faster local gap-fill for this one-off (no 5-min batch pause)
os.environ.setdefault("GAP_FILL_BATCH_PAUSE_SEC", "0")
os.environ.setdefault("GAP_FILL_CONCURRENT", "2")

MARKET_DIR = ROOT / "output" / "technical_foam"
JSON_NAME = "global_technical_foam_market_global.json"
DESKTOP = Path.home() / "Desktop"

CONTINENT_ONLY = re.compile(
    r"^(north america|south america|latin america|europe|asia|asia[- ]pacific|"
    r"africa|middle east|global|worldwide|emea|apac|americas)"
    r"(?:\s*[,;/&]\s*"
    r"(north america|south america|latin america|europe|asia|asia[- ]pacific|"
    r"africa|middle east|global|worldwide|emea|apac|americas))*$",
    re.I,
)


def _norm(v: object) -> str:
    return str(v or "").strip()


def _crawl(row: dict) -> dict:
    c = row.get("_crawl")
    if not isinstance(c, dict):
        c = {}
        row["_crawl"] = c
    return c


def prepare_row(row: dict) -> list[str]:
    """Clear bad cities / company-as-key-brand so gap-fill + brand resolve re-run."""
    from vendor_intel.export.fmcg_landscape import is_plausible_location

    cleared: list[str] = []
    crawl = _crawl(row)
    company = _norm(row.get("company") or row.get("brand"))

    for key in ("offices", "manufacturing_locations"):
        val = _norm(crawl.get(key))
        if not val:
            continue
        bad = (
            CONTINENT_ONLY.match(val) is not None
            or not is_plausible_location(val, allow_country_only=False)
        )
        if bad:
            crawl[key] = ""
            cleared.append(f"clear_{key}")

    kb = _norm(crawl.get("key_brands_represented"))
    if kb and company and kb.lower() == company.lower():
        crawl["key_brands_represented"] = ""
        cleared.append("clear_key_brands")

    return cleared


def _reload_founded_hq(data: dict, *, limit: int | None) -> list[dict]:
    from export_founded_hq_v2 import has_founded, has_headquarters

    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    founded_hq = [r for r in rows if has_founded(r) and has_headquarters(r)]
    if limit:
        founded_hq = founded_hq[:limit]
    return founded_hq


async def run(
    *,
    limit: int | None,
    export_only: bool,
    concurrent: int,
    skip_brands: bool,
    skip_verify: bool,
    from_verify: bool = False,
) -> None:
    from export_founded_hq_v2 import export_one, has_founded, has_headquarters
    from vendor_intel.enrichment.brand_domain_resolve import resolve_brand_domains_for_market
    from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_market

    json_path = MARKET_DIR / JSON_NAME
    if not json_path.is_file():
        raise SystemExit(f"Missing JSON: {json_path}")

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    founded_hq = [r for r in rows if has_founded(r) and has_headquarters(r)]
    if limit:
        founded_hq = founded_hq[:limit]

    print(
        f"Technical Foam: {len(rows)} relevant | "
        f"{len(founded_hq)} Founded+HQ ONLY (fill + verify scope)",
        flush=True,
    )

    if not export_only and not from_verify:
        prep_stats = {"cleared": 0}
        for row in founded_hq:
            got = prepare_row(row)
            if got:
                prep_stats["cleared"] += 1
        print(f"  Prepared {prep_stats['cleared']} rows (cleared bad cities / company key-brands)", flush=True)
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

        cache_path = MARKET_DIR / "gap_fill_cache.json"
        if cache_path.is_file():
            try:
                cache = json.loads(cache_path.read_text(encoding="utf-8"))
            except Exception:
                cache = {}
            cache.pop("market_completed", None)
            companies = dict(cache.get("companies") or {})
            for row in founded_hq:
                key = _norm(row.get("company") or row.get("brand")).lower()
                companies.pop(key, None)
            cache["companies"] = companies
            cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

        # Gap-fill ONLY the Founded+HQ rows (same object refs as main JSON)
        subset = {**data, "relevant_companies": founded_hq}
        subset_path = MARKET_DIR / "_founded_hq_enrich_subset.json"
        subset_path.write_text(json.dumps(subset, indent=2, ensure_ascii=False), encoding="utf-8")

        # Skip LinkedIn in gap-fill — STEP 2 does LinkedIn MCP verify (avoids hang on MCP).
        # Skip owler/crunchbase/backfill first pass — they hang when DDGS is rate-limited.
        print("\n=== STEP 1/4: FILL — Gap-fill (cities, contacts, locations; no LinkedIn yet) ===", flush=True)
        print(f"  Queue: {len(founded_hq)} companies (Founded+HQ only)", flush=True)
        print("  Tip: progress prints as [gap_fill] starting [n/N] Company", flush=True)
        gap_stats = await gap_fill_market(
            subset,
            json_path=subset_path,
            output_dir=MARKET_DIR,
            market_type="fmcg",
            fmcg_xlsx=False,
            also_csv=False,
            push_neon=False,
            concurrent=concurrent,
            sources="wikidata,contact_page,leadership,wikipedia,locations",
            force=False,
        )
        print(f"  Gap-fill done: {json.dumps(gap_stats, default=str)[:800]}", flush=True)
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        if cache_path.is_file():
            try:
                cache = json.loads(cache_path.read_text(encoding="utf-8"))
                cache["market_completed"] = False
                cache["subset_enrich_note"] = "founded_hq_only technical foam fill+verify"
                cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
            except Exception:
                pass
    elif from_verify and not export_only:
        print("\n=== STEP 1/4: FILL — SKIPPED (--from-verify; using gap-fill already done) ===", flush=True)

    if not export_only:
        # STEP 2: LinkedIn MCP verify ONLY on the same Founded+HQ set
        if not skip_verify:
            founded_hq = _reload_founded_hq(data, limit=limit)
            from vendor_intel.integrations.linkedin_mcp import (
                linkedin_mcp_available,
                linkedin_mcp_delay,
                linkedin_mcp_status,
            )

            print(f"  LinkedIn MCP status: {linkedin_mcp_status()}", flush=True)
            if linkedin_mcp_available():
                print(
                    f"\n=== STEP 2/4: VERIFY — LinkedIn MCP Founded+HQ "
                    f"(~{linkedin_mcp_delay():.1f}s/call, {len(founded_hq)} companies) ===",
                    flush=True,
                )
                from vendor_intel.enrichment.linkedin_firmographic_verify import (
                    verify_firmographics_for_market,
                )

                li_stats = await verify_firmographics_for_market(
                    founded_hq,
                    concurrent=1,
                    verify_all=True,
                )
                print(f"  LinkedIn verify: {json.dumps(li_stats, default=str)}", flush=True)
                json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
            else:
                print(
                    "\n=== STEP 2/4: VERIFY — SKIPPED (LinkedIn MCP not available on :8001) ===",
                    flush=True,
                )
                print(
                    "  Keep MCP running in another terminal, then re-run:\n"
                    "  .\\.venv\\Scripts\\python.exe scripts\\enrich_technical_foam_founded_hq.py --from-verify",
                    flush=True,
                )

        if not skip_brands:
            print("\n=== STEP 3/4: Key Brands (genuine brands, not company name) ===", flush=True)
            brand_cache = MARKET_DIR / "brand_resolve_cache.json"
            if brand_cache.is_file():
                try:
                    bc = json.loads(brand_cache.read_text(encoding="utf-8"))
                except Exception:
                    bc = {}
                founded_hq = _reload_founded_hq(data, limit=limit)
                for row in founded_hq:
                    key = _norm(row.get("company") or row.get("brand")).lower()
                    if key in bc and bc[key].get("key_brands_only"):
                        kb = _norm((row.get("_crawl") or {}).get("key_brands_represented"))
                        if not kb or kb.lower() == key:
                            bc.pop(key, None)
                brand_cache.write_text(json.dumps(bc, indent=2, ensure_ascii=False), encoding="utf-8")

            ctx = data.get("query_context") or {}
            market_query = str(ctx.get("industry") or data.get("query") or "Technical Foam")
            founded_hq = _reload_founded_hq(data, limit=limit)

            brand_stats = await resolve_brand_domains_for_market(
                founded_hq,
                market_query=market_query,
                market_dir=MARKET_DIR,
                concurrent=max(1, min(concurrent, 2)),
                founded_hq_only=False,
                key_brands_only=True,
                force=False,
            )
            print(f"  Brand resolve: {json.dumps(brand_stats, default=str)}", flush=True)
            json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    print("\n=== STEP 4/4: Export Founded+HQ Excel (filtered only) ===", flush=True)
    export_stats = export_one(MARKET_DIR, suffix="founded_and_hq")
    print(f"  Export: {json.dumps(export_stats, default=str)}", flush=True)

    out_file = Path(export_stats["file"])
    desktop_out = DESKTOP / out_file.name
    shutil.copy2(out_file, desktop_out)
    exports_dir = ROOT / "output" / "demo" / "exports"
    exports_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(out_file, exports_dir / out_file.name)

    print(f"\nDone. Excel written to:")
    print(f"  {out_file}")
    print(f"  {desktop_out}")
    print(f"  {exports_dir / out_file.name}")
    print(f"Completed at {datetime.now(timezone.utc).isoformat()}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fill + verify ONLY Technical Foam Founded+HQ companies (~98)",
    )
    parser.add_argument("--limit", type=int, default=None, help="Only first N Founded+HQ companies")
    parser.add_argument("--export-only", action="store_true", help="Skip enrich; only regenerate Excel")
    parser.add_argument(
        "--from-verify",
        action="store_true",
        help="Skip STEP 1 gap-fill; resume from LinkedIn verify + key brands + Excel",
    )
    parser.add_argument("--concurrent", type=int, default=2)
    parser.add_argument("--skip-brands", action="store_true", help="Skip key-brands LLM resolve")
    parser.add_argument("--skip-verify", action="store_true", help="Skip LinkedIn MCP verify step")
    args = parser.parse_args()
    asyncio.run(
        run(
            limit=args.limit,
            export_only=args.export_only,
            concurrent=args.concurrent,
            skip_brands=args.skip_brands,
            skip_verify=args.skip_verify,
            from_verify=args.from_verify,
        )
    )


if __name__ == "__main__":
    main()
