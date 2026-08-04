#!/usr/bin/env python3
"""
contact_fill_all.py — fill email, phone, and CEO/contact person for ALL market exports.

Visits each company's Contact Us and About/Leadership pages, then merge-pushes to Neon.

Usage:
  .venv\\Scripts\\python contact_fill_all.py --yes
  .venv\\Scripts\\python contact_fill_all.py --yes --force
  .venv\\Scripts\\python contact_fill_all.py --yes --limit 2
  .venv\\Scripts\\python contact_fill_all.py --yes --max-concurrent 2 --gap-concurrent 3
"""
from __future__ import annotations

import argparse
import asyncio
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

SKIP_NAMES = {
    "audit_cache.json",
    "gap_fill_cache.json",
    "contact_fill_cache.json",
    "discovery_new_companies.json",
    "session_log.json",
    "discovery_log.json",
    "removal_log.json",
    "flagged_companies.json",
    "final_audit_v2_summary.json",
}
SKIP_PREFIXES = ("phase1_plan_", "phase2_companies_", "market_research_", "gap_fill_summary_")


def _is_export(path: Path) -> bool:
    if path.name in SKIP_NAMES:
        return False
    if path.name.startswith(SKIP_PREFIXES):
        return False
    if path.parent.name in ("phase1", "phase2", "neon_export"):
        return False
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return False
    if not isinstance(data, dict):
        return False
    rows = data.get("relevant_companies") or []
    return isinstance(rows, list) and len(rows) > 0


def discover_market_slugs(output_root: Path) -> list[str]:
    slugs: list[str] = []
    seen: set[str] = set()
    for path in sorted(output_root.rglob("*.json")):
        if not _is_export(path):
            continue
        slug = path.parent.name
        if slug not in seen:
            seen.add(slug)
            slugs.append(slug)
    return slugs


async def contact_fill_all_async(
    *,
    max_concurrent: int = 2,
    gap_concurrent: int = 3,
    force: bool = False,
    fmcg_xlsx: bool = False,
    also_csv: bool = False,
    limit: int | None = None,
    push_neon: bool = True,
) -> dict:
    from vendor_intel.config import _project_root
    from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_market_from_slug, load_cache

    output_root = _project_root() / "output"
    slugs = discover_market_slugs(output_root)
    if not slugs:
        print("  [contact_fill] No market exports found under output/", flush=True)
        return {"markets": 0}

    todo: list[str] = []
    for slug in slugs:
        cache = load_cache(output_root / slug, ("contact_page", "leadership"))
        if cache.get("market_completed") and not force:
            continue
        todo.append(slug)

    if limit is not None:
        todo = todo[: max(0, limit)]

    print(f"\n  [contact_fill] {len(todo)} market(s) to process (of {len(slugs)} exports)", flush=True)
    if not todo:
        print("  [contact_fill] Nothing pending — use --force to re-run.", flush=True)
        return {"markets": len(slugs), "processed": 0}

    sem = asyncio.Semaphore(max(1, max_concurrent))
    results: list[dict] = []
    t0 = time.perf_counter()

    async def _one(slug: str) -> None:
        async with sem:
            print(f"\n  [contact_fill] === {slug} ===", flush=True)
            try:
                stats = await gap_fill_market_from_slug(
                    slug,
                    fmcg_xlsx=fmcg_xlsx,
                    concurrent=gap_concurrent,
                    sources="contact_page,leadership",
                    force=force,
                    also_csv=also_csv,
                    push_neon=push_neon,
                )
                results.append({"slug": slug, **stats})
            except Exception as exc:
                print(f"  [contact_fill] ERROR {slug}: {exc}", flush=True)
                results.append({"slug": slug, "error": str(exc)})

    await asyncio.gather(*[_one(s) for s in todo])

    summary = {
        "total_exports": len(slugs),
        "processed": len(todo),
        "results": results,
        "elapsed_minutes": round((time.perf_counter() - t0) / 60.0, 2),
    }
    out_path = output_root / "contact_fill_summary_all.json"
    out_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\n  [contact_fill] Done — summary: {out_path}", flush=True)
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Contact + CEO enrichment for all market exports")
    parser.add_argument("--max-concurrent", "-c", type=int, default=2, help="Markets in parallel (default: 2)")
    parser.add_argument("--gap-concurrent", type=int, default=3, help="Companies per market in parallel (default: 3)")
    parser.add_argument("--force", action="store_true", help="Re-run even if contact_fill_cache says done")
    parser.add_argument("--limit", type=int, default=None, help="Process only N markets (smoke test)")
    parser.add_argument("--no-push-neon", action="store_true", help="Skip Neon merge-push")
    parser.add_argument("--yes", "-y", action="store_true", help="Skip confirmation")
    args = parser.parse_args()

    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.utils.output_filter import install_stderr_filter

    install_stderr_filter()
    apply_env_overrides()

    from vendor_intel.config import _project_root

    slugs = discover_market_slugs(_project_root() / "output")
    if not args.yes:
        try:
            confirm = input(f"  Run contact fill on up to {len(slugs)} market(s)? [Enter=go]: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n  Cancelled.")
            return
        if confirm:
            print("  Cancelled.")
            return

    try:
        asyncio.run(
            contact_fill_all_async(
                max_concurrent=args.max_concurrent,
                gap_concurrent=args.gap_concurrent,
                force=args.force,
                limit=args.limit,
                push_neon=not args.no_push_neon,
            )
        )
    finally:
        try:
            from vendor_intel.clients.ddg_worker_pool import shutdown_ddg_pool

            shutdown_ddg_pool(wait=True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
