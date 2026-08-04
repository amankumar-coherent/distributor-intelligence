#!/usr/bin/env python3
"""
Remove audit-flagged companies (action=remove) from all market CSV files.
Keep and Review companies are preserved exactly as-is.
"""
import json
import csv
import io
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent

MARKETS = [
    "avocado_oil",
    "technical_foam",
    "spectral_sensor",
    "gel_ice_pack",
    "cyber_physical_systems_security",
    "high_energy_laser",
]


def load_remove_set(cache_path: Path) -> dict[str, set[str]]:
    with open(cache_path, encoding="utf-8") as f:
        cache = json.load(f)
    remove_sets = {}
    for market, results in cache.items():
        to_remove = {r["company"].strip().lower() for r in results if r.get("action") == "remove"}
        remove_sets[market] = to_remove
    return remove_sets


def remove_from_csv(market: str, remove_names: set[str]) -> dict:
    csv_path = ROOT / f"output/{market}/global_{market}_market_global.csv"
    if not csv_path.exists():
        return {"market": market, "error": f"File not found: {csv_path}"}

    # Backup
    backup_path = csv_path.with_suffix(".csv.bak")
    if not backup_path.exists():
        shutil.copy2(csv_path, backup_path)
        print(f"  Backed up to: {backup_path.name}")

    with open(csv_path, encoding="utf-8", errors="replace") as f:
        raw_lines = f.readlines()

    # Split header line, section headers (=== ... ===) and data lines
    csv_header_line = raw_lines[0]
    rest_lines = raw_lines[1:]

    # Parse with DictReader (skip === section lines)
    data_lines_only = [l for l in rest_lines if not l.strip().startswith("===")]
    section_headers = {i: l for i, l in enumerate(rest_lines) if l.strip().startswith("===")}

    reader = csv.DictReader(io.StringIO(csv_header_line + "".join(data_lines_only)))
    fieldnames = reader.fieldnames
    all_rows = list(reader)

    before = len(all_rows)
    kept_rows = [r for r in all_rows if r.get("Company", "").strip().lower() not in remove_names]
    removed_names = [r.get("Company", "").strip() for r in all_rows if r.get("Company", "").strip().lower() in remove_names]
    after = len(kept_rows)

    # Write back — rebuild with original section headers preserved in right spots
    # Rebuild: we'll write without section headers first, then re-insert them
    # Strategy: write clean CSV (section headers stripped — they were cosmetic)
    with open(csv_path, "w", encoding="utf-8", newline="", errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(kept_rows)

    print(f"  Saved: {csv_path.name}")
    return {
        "market": market,
        "before": before,
        "after": after,
        "removed_count": len(removed_names),
        "removed": removed_names,
    }


def main():
    print("=" * 65)
    print("REMOVING AUDIT-FLAGGED COMPANIES FROM CSV FILES")
    print("(Keep + Review companies are preserved)")
    print("=" * 65)

    cache_path = ROOT / "audit_results_cache.json"
    remove_sets = load_remove_set(cache_path)

    total_removed = 0
    results = []

    for market in MARKETS:
        print(f"\n[{market.upper()}]")
        remove_names = remove_sets.get(market, set())
        if not remove_names:
            print("  Nothing to remove.")
            results.append({"market": market, "removed_count": 0})
            continue

        result = remove_from_csv(market, remove_names)
        results.append(result)

        if "error" in result:
            print(f"  ERROR: {result['error']}")
        else:
            total_removed += result["removed_count"]
            print(f"  Removed {result['removed_count']} rows | Before: {result['before']} | After: {result['after']}")
            if result["removed"]:
                print(f"  Deleted: {', '.join(result['removed'])}")

    print("\n" + "=" * 65)
    print(f"DONE! Total companies removed from CSVs: {total_removed}")
    print("=" * 65)

    print(f"\n{'Market':<40} {'Before':>7} {'Removed':>8} {'After':>7}")
    print("-" * 65)
    for r in results:
        if "error" not in r:
            print(f"  {r['market']:<38} {r.get('before','?'):>7} {r.get('removed_count',0):>8} {r.get('after','?'):>7}")


if __name__ == "__main__":
    main()
