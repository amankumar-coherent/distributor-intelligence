#!/usr/bin/env python3
"""Verify market fit + field accuracy using OpenAI and web evidence.

Usage:
  .venv\\Scripts\\python.exe verify_company_data.py cyber_physical_systems_security
  .venv\\Scripts\\python.exe verify_company_data.py cyber_physical_systems_security --founded-hq-only
  .venv\\Scripts\\python.exe verify_company_data.py cyber_physical_systems_security --no-web-search
"""
from __future__ import annotations

import argparse
import asyncio
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

MARKET_DESCRIPTIONS = {
    "cyber_physical_systems_security": (
        "cybersecurity for cyber-physical systems (CPS), OT/ICS security, SCADA protection, "
        "industrial control system security, critical infrastructure protection"
    ),
    "avocado_oil": "avocado oil production, processing, refining, bottling, or distribution",
    "technical_foam": "technical or industrial foam manufacturing for industrial/commercial use",
}


def _filled(v: object) -> bool:
    s = str(v or "").strip()
    return bool(s) and s.lower() not in ("null", "none", "n/a", "not publicly disclosed", "unknown", "-")


def _has_founded(row: dict) -> bool:
    c = row.get("_crawl") or {}
    if _filled(c.get("founded_year")) or _filled(row.get("founded_year")):
        return True
    return bool(re.search(r"\bfounded(?:\s+in)?\s+\d{4}\b", str(c.get("crawl_text") or ""), re.I))


def _has_hq(row: dict) -> bool:
    c = row.get("_crawl") or {}
    return _filled(c.get("headquarters")) or _filled(c.get("contact_address"))


def main() -> None:
    parser = argparse.ArgumentParser(description="OpenAI verify market fit + company fields")
    parser.add_argument("slug", nargs="?", default="cyber_physical_systems_security")
    parser.add_argument(
        "--founded-hq-only",
        action="store_true",
        help="Verify only companies with both Founded and HQ filled",
    )
    parser.add_argument("--force", action="store_true", help="Re-verify even if verify_cache exists")
    parser.add_argument("--concurrent", type=int, default=3)
    parser.add_argument("--limit", type=int, default=None, help="Verify only first N companies (smoke test)")
    parser.add_argument("--no-web-search", action="store_true", help="Use stored crawl text only")
    args = parser.parse_args()

    from vendor_intel.enrichment.company_data_verify import verify_companies
    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()

    out_dir = ROOT / "output" / args.slug
    json_files = [p for p in sorted(out_dir.glob("*.json")) if "market_research" not in p.name and "verify" not in p.name]
    json_path = next((p for p in json_files if "global" in p.name), json_files[0] if json_files else None)
    if not json_path:
        raise SystemExit(f"No export JSON in {out_dir}")

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    if args.founded_hq_only:
        rows = [r for r in rows if _has_founded(r) and _has_hq(r)]
    if args.limit is not None:
        rows = rows[: max(0, args.limit)]

    ctx = data.get("query_context") or {}
    market_query = str(ctx.get("industry") or data.get("query") or args.slug.replace("_", " "))
    market_desc = MARKET_DESCRIPTIONS.get(args.slug, market_query)

    print(f"\n=== VERIFY: {market_query} ===")
    print(f"  Companies: {len(rows)}")
    print(f"  OpenAI: required | Web search: {not args.no_web_search}")

    stats = asyncio.run(
        verify_companies(
            rows,
            market_query=market_query,
            market_desc=market_desc,
            output_dir=out_dir,
            concurrent=args.concurrent,
            use_web_search=not args.no_web_search,
            force=args.force,
        )
    )
    print(json.dumps(stats, indent=2))


if __name__ == "__main__":
    main()
