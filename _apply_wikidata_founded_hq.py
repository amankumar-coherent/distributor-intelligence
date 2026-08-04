#!/usr/bin/env python3
"""Apply the Wikidata Founded+HQ fallback to existing market folders, then re-export.

Runs the same fallback that is now built into the post-market pipeline
(vendor_intel.enrichment.wikidata_firmographic_fill) on one or more market
folders under output/demo/, re-exports the Founded+HQ Excel, and (optionally)
pushes the refreshed rows to Neon.

Usage:
  python _apply_wikidata_founded_hq.py global_firewall_market_global global_green_chemical_market_global
  python _apply_wikidata_founded_hq.py --all
  python _apply_wikidata_founded_hq.py <slug> --no-neon
Safe to Ctrl+C — JSON is saved after every company.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402
from vendor_intel.enrichment.gap_fill.gaps import has_founded, has_headquarters  # noqa: E402
from vendor_intel.enrichment.wikidata_firmographic_fill import (  # noqa: E402
    fill_firmographics_via_wikidata_for_market,
)
from vendor_intel.enrichment.wikipedia_firmographic_fill import (  # noqa: E402
    fill_firmographics_via_wikipedia_for_market,
)
from export_founded_hq_v2 import export_one, find_main_json  # noqa: E402

DEMO = ROOT / "output" / "demo"


def _count_both(rows: list) -> int:
    return sum(
        1 for r in rows
        if r.get("is_relevant", True) and has_founded(r) and has_headquarters(r)
    )


async def apply_one(market_dir: Path, *, push_neon: bool) -> dict:
    jp = find_main_json(market_dir)
    if not jp:
        return {"market_dir": str(market_dir), "error": "no main json"}
    data = json.loads(jp.read_text(encoding="utf-8"))
    rows = data.get("relevant_companies") or []

    def _persist() -> None:
        jp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    before = _count_both(rows)
    print(f"\n=== {market_dir.name} — Founded+HQ before: {before} ===", flush=True)

    wd = await fill_firmographics_via_wikidata_for_market(rows, persist_fn=_persist)
    _persist()
    print("  wikidata stats:", json.dumps(wd), flush=True)

    wp = await fill_firmographics_via_wikipedia_for_market(rows, persist_fn=_persist)
    _persist()
    print("  wikipedia stats:", json.dumps(wp), flush=True)

    after = _count_both(rows)
    exp = export_one(market_dir, suffix="founded_and_hq_v2")
    print(f"  Founded+HQ after: {after} (+{after - before}) | export filtered: {exp.get('filtered')}", flush=True)

    result = {"slug": market_dir.name, "before": before, "after": after, "export": exp, "wikidata": wd, "wikipedia": wp}

    if push_neon:
        try:
            from vendor_intel.neon_push import push_json_to_neon, should_push_neon

            neon_path = market_dir / "_neon_filtered_export.json"
            if should_push_neon() and neon_path.is_file():
                ns = push_json_to_neon(neon_path, merge_companies=False)
                result["neon_push"] = ns
                print(f"  [neon] Pushed {ns.get('vendor_intel_companies', 0)} companies", flush=True)
            elif not should_push_neon():
                print("  [neon] Skipped — DATABASE_URL not set", flush=True)
        except Exception as exc:
            result["neon_push_error"] = str(exc)
            print(f"  [neon] Push failed: {exc}", flush=True)

    return result


async def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("slugs", nargs="*", help="market folder names under output/demo/")
    ap.add_argument("--all", action="store_true", help="process every folder under output/demo/")
    ap.add_argument("--no-neon", action="store_true", help="skip Neon push")
    args = ap.parse_args()

    apply_env_overrides()

    if args.all:
        dirs = [p for p in sorted(DEMO.iterdir()) if p.is_dir() and find_main_json(p)]
    else:
        dirs = [DEMO / s for s in args.slugs]

    if not dirs:
        print("No market folders given. Pass slugs or --all.")
        return

    results = []
    for d in dirs:
        if not d.is_dir():
            print(f"SKIP (missing): {d}")
            continue
        results.append(await apply_one(d, push_neon=not args.no_neon))

    print("\n=== SUMMARY ===")
    for r in results:
        if r.get("error"):
            print(f"  {r.get('slug', r.get('market_dir'))}: ERROR {r['error']}")
        else:
            print(f"  {r['slug']}: {r['before']} -> {r['after']} Founded+HQ")


if __name__ == "__main__":
    asyncio.run(main())
