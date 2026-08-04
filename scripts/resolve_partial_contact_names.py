#!/usr/bin/env python3
"""Upgrade partial Apollo contact names via waterfall:

  LinkedIn MCP → leadership page → Google/web → RocketReach

Usage:
  .venv\\Scripts\\python.exe scripts/resolve_partial_contact_names.py technical_foam
  .venv\\Scripts\\python.exe scripts/resolve_partial_contact_names.py technical_foam --limit 10
  .venv\\Scripts\\python.exe scripts/resolve_partial_contact_names.py technical_foam --skip-linkedin-mcp
"""
from __future__ import annotations

import argparse
import asyncio
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))


async def main() -> None:
    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()

    parser = argparse.ArgumentParser(description="Resolve partial Apollo contact names")
    parser.add_argument("slug", nargs="?", default="technical_foam")
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument(
        "--concurrent",
        type=int,
        default=8,
        help="Parallel workers (default 8; forced to 1 when LinkedIn MCP is enabled)",
    )
    parser.add_argument("--skip-linkedin-mcp", action="store_true")
    parser.add_argument("--skip-leadership", action="store_true")
    parser.add_argument("--skip-google", action="store_true")
    parser.add_argument("--skip-rocketreach", action="store_true")
    parser.add_argument("--founded-hq-only", action="store_true", default=True)
    parser.add_argument("--all-companies", action="store_true")
    parser.add_argument(
        "--force",
        action="store_true",
        help="Clear name_resolve_attempted and retry masked/partial names",
    )
    parser.add_argument(
        "--market-dir",
        type=str,
        default="",
        help="Market export folder (overrides output/<slug>)",
    )
    args = parser.parse_args()

    from export_founded_hq_v2 import export_one, find_main_json, has_founded, has_headquarters
    from vendor_intel.enrichment.contact_name_resolve import (
        is_partial_name,
        needs_name_resolve,
        resolve_contact_names_for_market,
    )
    from vendor_intel.enrichment.gap_fill.gaps import norm

    market_dir = Path(args.market_dir) if args.market_dir else (ROOT / "output" / args.slug)
    json_path = find_main_json(market_dir)
    if not json_path:
        raise SystemExit(f"No export JSON in {market_dir}")

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    if not args.all_companies and args.founded_hq_only:
        rows = [r for r in rows if has_founded(r) and has_headquarters(r)]

    partials = [
        r
        for r in rows
        if needs_name_resolve(r)
        or is_partial_name(norm((r.get("_crawl") or {}).get("contact_person")))
    ]
    if args.force:
        for r in partials:
            crawl = dict(r.get("_crawl") or {})
            crawl.pop("name_resolve_attempted", None)
            r["_crawl"] = crawl
    if args.limit:
        partials = partials[: max(0, args.limit)]

    workers = max(1, int(args.concurrent or 8))
    print(f"=== PARTIAL NAME WATERFALL: {args.slug} ===")
    print(f"  Founded+HQ rows: {len(rows)}")
    print(f"  Need resolve: {len(partials)}")
    print(f"  Workers: {workers} (skip_linkedin_mcp={args.skip_linkedin_mcp})")
    print("  Order: LinkedIn MCP → leadership → Google → RocketReach")
    for r in partials[:12]:
        c = r.get("_crawl") or {}
        print(f"    - {r.get('company')}: {c.get('contact_person')!r}")
    if len(partials) > 12:
        print(f"    … +{len(partials) - 12} more")

    if not partials:
        print("  Nothing to do.")
        return

    def _persist() -> None:
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    stats = await resolve_contact_names_for_market(
        partials,
        concurrent=workers,
        skip_leadership=args.skip_leadership,
        skip_linkedin_mcp=args.skip_linkedin_mcp,
        skip_google=args.skip_google,
        skip_rocketreach=args.skip_rocketreach,
        persist_fn=_persist,
    )
    print(f"  Stats: {json.dumps(stats)}")
    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")

    export_stats = export_one(market_dir, suffix="founded_and_hq")
    print(f"  Excel: {export_stats.get('file')}")


if __name__ == "__main__":
    asyncio.run(main())
