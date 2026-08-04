#!/usr/bin/env python3
"""
Remove audit-flagged companies (action=remove) from all market Excel files.
Keep and Review companies are preserved exactly as-is.
"""
import json
import openpyxl
import shutil
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    """Return market -> set of company names to remove (action == 'remove')."""
    with open(cache_path, encoding="utf-8") as f:
        cache = json.load(f)
    remove_sets = {}
    for market, results in cache.items():
        to_remove = set()
        for r in results:
            if r.get("action") == "remove":
                # normalize: strip, lower for matching
                to_remove.add(r["company"].strip().lower())
        remove_sets[market] = to_remove
        print(f"  {market}: {len(to_remove)} companies to remove")
    return remove_sets


def find_header_row(ws) -> int:
    """Return the 1-indexed row number where 'Company' header appears."""
    for i, row in enumerate(ws.rows, 1):
        if row[0].value == "Company":
            return i
    raise ValueError("Could not find 'Company' header row in worksheet")


def delete_company_rows(ws, header_row: int, remove_names: set[str]) -> list[str]:
    """
    Delete rows where col-A company name (lowercased) is in remove_names.
    Returns list of deleted company names.
    """
    deleted = []
    # Collect rows to delete (iterate bottom-up to preserve indices)
    rows_to_delete = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        cell_val = row[0].value
        if cell_val is None:
            continue
        if str(cell_val).strip().lower() in remove_names:
            rows_to_delete.append(row[0].row)
            deleted.append(str(cell_val).strip())

    # Delete from bottom to top so row numbers stay valid
    for row_num in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_num)

    return deleted


def re_number_rows(ws, header_row: int):
    """Re-number the # column (first col may have row numbers, skip if not numeric)."""
    # Check if there's a number column — typically col B is Company and col A is #
    # In our Excel col A = Company, so skip renumbering
    pass


def remove_from_excel(market: str, remove_names: set[str]) -> dict:
    xlsx_path = ROOT / f"output/{market}/global_{market}_market_global.xlsx"
    if not xlsx_path.exists():
        return {"market": market, "error": f"File not found: {xlsx_path}"}

    # Backup original
    backup_path = xlsx_path.with_suffix(".xlsx.bak")
    if not backup_path.exists():
        shutil.copy2(xlsx_path, backup_path)
        print(f"  Backed up to: {backup_path.name}")

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    before_count = ws.max_row
    header_row = find_header_row(ws)
    data_rows_before = ws.max_row - header_row

    deleted = delete_company_rows(ws, header_row, remove_names)
    data_rows_after = ws.max_row - header_row

    wb.save(xlsx_path)
    print(f"  Saved: {xlsx_path.name}")

    return {
        "market": market,
        "companies_before": data_rows_before,
        "companies_after": data_rows_after,
        "deleted_count": len(deleted),
        "deleted": deleted,
    }


def main():
    print("=" * 65)
    print("REMOVING AUDIT-FLAGGED COMPANIES FROM EXCEL FILES")
    print("(Keep + Review companies are preserved)")
    print("=" * 65)

    cache_path = ROOT / "audit_results_cache.json"
    if not cache_path.exists():
        print("ERROR: audit_results_cache.json not found. Run final_audit.py first.")
        return

    print("\nLoading companies to remove...")
    remove_sets = load_remove_set(cache_path)

    print("\nProcessing Excel files...")
    results = []
    total_removed = 0

    for market in MARKETS:
        print(f"\n[{market.upper()}]")
        remove_names = remove_sets.get(market, set())
        if not remove_names:
            print("  Nothing to remove.")
            results.append({"market": market, "deleted_count": 0, "deleted": []})
            continue

        print(f"  Removing: {', '.join(remove_names)}")
        result = remove_from_excel(market, remove_names)
        results.append(result)

        if "error" in result:
            print(f"  ERROR: {result['error']}")
        else:
            removed = result["deleted_count"]
            total_removed += removed
            print(f"  Removed {removed} rows | Before: {result['companies_before']} | After: {result['companies_after']}")
            if result["deleted"]:
                print(f"  Deleted: {', '.join(result['deleted'])}")

    print("\n" + "=" * 65)
    print(f"DONE! Total companies removed: {total_removed}")
    print("=" * 65)

    # Save removal log
    log_path = ROOT / "output" / "removal_log.json"
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nRemoval log saved to: {log_path}")

    # Print summary table
    print("\nSummary:")
    print(f"{'Market':<40} {'Before':>7} {'Removed':>8} {'After':>7}")
    print("-" * 65)
    for r in results:
        if "error" not in r:
            before = r.get("companies_before", "?")
            after = r.get("companies_after", "?")
            removed = r.get("deleted_count", 0)
            print(f"  {r['market']:<38} {before:>7} {removed:>8} {after:>7}")


if __name__ == "__main__":
    main()
