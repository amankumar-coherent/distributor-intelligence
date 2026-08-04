#!/usr/bin/env python3
"""Owler firmographic enrichment for pipeline JSON exports."""
from __future__ import annotations

import argparse
import asyncio
import glob
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from vendor_intel.enrichment.owler_enrich import enrich_companies_owler
from vendor_intel.enrichment.post_export_enrich import re_export_market
from vendor_intel.placeholders.load_keys import apply_env_overrides


async def enrich_market_file(json_path: Path, *, market_type: str = "general") -> None:
    print(f"\n{'=' * 56}")
    print(f"ENRICHING MARKET FROM OWLER: {json_path.parent.name}")
    print(f"File: {json_path.name}")
    print(f"{'=' * 56}")

    try:
        result = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"Failed to load JSON {json_path}: {e}")
        return

    companies = [c for c in (result.get("relevant_companies") or []) if c.get("is_relevant", True)]
    total = len(companies)
    if total == 0:
        print("No relevant companies to enrich. Skipping.")
        return

    print(f"Found {total} relevant companies to check on Owler.")
    updated, _ = await enrich_companies_owler(companies)
    result["owler_enriched"] = updated
    result["owler_enrich_total"] = total

    re_export_market(result, json_path, market_type=market_type, fmcg_xlsx=True)
    print(f"Saved exports under {json_path.parent.resolve()}")


async def main_async(json_paths: list[Path] | None = None) -> None:
    start_time = time.perf_counter()
    apply_env_overrides()

    from backend.config import get_settings

    cfg = get_settings()
    if not cfg.openai_api_key:
        print("ERROR: OPENAI_API_KEY is not set in environment or .env file.")
        sys.exit(1)

    if json_paths is None:
        json_paths = [Path(p) for p in glob.glob("output/*/*.json")]
        json_paths = [p for p in json_paths if not p.name.startswith("_")]

    if not json_paths:
        print("No JSON files found under output/*/*.json")
        return

    print(f"Found {len(json_paths)} market file(s) to enrich.")

    for json_path in json_paths:
        await enrich_market_file(json_path)

    elapsed = time.perf_counter() - start_time
    print(f"\n{'=' * 56}")
    print(f"ALL OWLER ENRICHMENTS COMPLETED in {elapsed / 60.0:.2f} minutes.")
    print(f"{'=' * 56}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Enrich pipeline JSON from Owler")
    parser.add_argument(
        "json_paths",
        nargs="*",
        help="JSON file(s). If omitted, scans output/*/*.json",
    )
    args = parser.parse_args()

    paths = [Path(p) for p in args.json_paths] if args.json_paths else None

    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main_async(paths))


if __name__ == "__main__":
    main()
