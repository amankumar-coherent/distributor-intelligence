#!/usr/bin/env python3
"""Push every vendor-intel pipeline JSON export under output/ into Neon.

Skips caches, phase plans, and files without relevant_companies.

Usage:
  python scripts/push_all_exports_to_neon.py
  python scripts/push_all_exports_to_neon.py --dry-run
  python scripts/push_all_exports_to_neon.py output/medical_alert_systems_market_global
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv

load_dotenv(ROOT / ".env")

SKIP_NAMES = {
    "audit_cache.json",
    "gap_fill_cache.json",
    "discovery_new_companies.json",
    "session_log.json",
    "discovery_log.json",
    "removal_log.json",
    "flagged_companies.json",
    "final_audit_v2_summary.json",
}
SKIP_PREFIXES = ("phase1_plan_", "phase2_companies_", "market_research_")


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


def _discover_exports(root: Path) -> list[Path]:
    found: list[Path] = []
    for path in sorted(root.rglob("*.json")):
        if _is_export(path):
            found.append(path)
    return found


def main() -> None:
    parser = argparse.ArgumentParser(description="Push all vendor-intel exports to Neon")
    parser.add_argument(
        "paths",
        nargs="*",
        help="Optional folders or JSON files (default: output/)",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    if args.paths:
        roots = [Path(p).resolve() for p in args.paths]
    else:
        roots = [ROOT / "output"]

    exports: list[Path] = []
    for root in roots:
        if root.is_file() and root.suffix == ".json":
            if _is_export(root):
                exports.append(root)
        elif root.is_dir():
            exports.extend(_discover_exports(root))

    exports = sorted(set(exports))
    if not exports:
        print("No pipeline exports found.")
        return

    import importlib.util

    mod_path = Path(__file__).resolve().parent / "push_export_to_neon.py"
    spec = importlib.util.spec_from_file_location("push_export_to_neon", mod_path)
    if spec is None or spec.loader is None:
        raise SystemExit(f"Cannot load {mod_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    push_export = mod.push_export

    print(f"Found {len(exports)} export(s)\n")
    totals = {
        "files": 0,
        "companies_upserted": 0,
        "markets_upserted": 0,
        "links_upserted": 0,
        "metadata_upserted": 0,
        "skipped_rows": 0,
        "errors": 0,
    }

    for path in exports:
        rel = path.relative_to(ROOT) if path.is_relative_to(ROOT) else path
        print(f"==> {rel}")
        try:
            stats = push_export(path, dry_run=args.dry_run)
            totals["files"] += 1
            for k in ("companies_upserted", "markets_upserted", "links_upserted", "metadata_upserted", "skipped"):
                if k == "skipped":
                    totals["skipped_rows"] += stats.get(k, 0)
                else:
                    totals[k] += stats.get(k, 0)
            print(
                f"    companies={stats.get('companies_upserted')} "
                f"market={stats.get('markets_upserted')} "
                f"metadata={stats.get('metadata_upserted')}\n"
            )
        except Exception as exc:
            totals["errors"] += 1
            print(f"    ERROR: {exc}\n")

    print("---- TOTAL ----")
    for k, v in totals.items():
        print(f"  {k}: {v}")
    print("Done.")


if __name__ == "__main__":
    main()
