#!/usr/bin/env python3
"""Export the final verified-distributor Excel files to Desktop\\DI.

Reads each market's *_verified_distributors.xlsx (KEEP-only rows, all
original columns, produced by verify_distributors_ai_mode.py) and writes a
polished copy to the destination folder: readable column widths and row
heights sized so every cell's content is visible without manually
expanding, instead of the narrow default columns Excel starts with.
"""
from __future__ import annotations

import shutil
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

DEST_DIR = Path(r"C:\Users\aman.kumar\Desktop\DI")

SOURCES = [
    ("Semiconductor_Market_verified_distributors.xlsx", "Semiconductor_Market_Distributors.xlsx"),
    (
        "output/web_expand/wearable_tech_distribution/wearable_tech_distributors_pasted_verified_distributors.xlsx",
        "Wearable_Tech_Distributors.xlsx",
    ),
    (
        "output/web_expand/glp1_receptor_agonist_distribution/glp1_receptor_agonist_distributors_verified_distributors.xlsx",
        "GLP1_Receptor_Agonist_Distributors.xlsx",
    ),
    (
        "output/market_research/global-data-center-system-market-distributors-only_verified_distributors.xlsx",
        "Global_Data_Center_System_Market_Distributors.xlsx",
    ),
    (
        "output/market_research/global-semiconductor-market-distributors-only_verified_distributors.xlsx",
        "Global_Semiconductor_Market_Distributors.xlsx",
    ),
    (
        "output/market_research/global-liquefied-natural-gas-market-distributors-only_verified_distributors.xlsx",
        "Global_LNG_Market_Distributors.xlsx",
    ),
]

# Per-header width hints (characters). Columns not listed fall back to
# DEFAULT_WIDTH. Wide text columns get enough room to show a full line or
# two without truncating; short categorical columns stay compact.
WIDTH_HINTS = {
    "#": 5,
    "Company": 28,
    "Website": 30,
    "Founded": 9,
    "Headquarters": 30,
    "Continent / Geography": 55,
    "Operational Presence": 45,
    "Ownership": 16,
    "Employees": 12,
    "Core Categories": 30,
    "Specialty Focus": 32,
    "Online / Offline / Both": 14,
    "Contact Person": 20,
    "Role": 24,
    "Email": 30,
    "LinkedIn": 40,
    "Office No.": 18,
    "Office Website": 30,
    "Office Email": 30,
    "Country Code / Region Code": 45,
}
DEFAULT_WIDTH = 32
HEADER_ROW_HEIGHT = 34
DATA_ROW_MIN_HEIGHT = 15


def polish_workbook(src_path: Path, dest_path: Path) -> tuple[int, int]:
    wb = openpyxl.load_workbook(src_path)
    ws = wb.active

    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    n_rows = ws.max_row - 1
    n_cols = ws.max_column

    # Column widths sized per header; anything containing long free text
    # (Core Categories, Specialty Focus, products/brands columns, etc.)
    # gets extra room since those are the ones that get clipped by
    # Excel's default width.
    for idx, header in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        width = WIDTH_HINTS.get(header)
        if width is None:
            h = (header or "")
            if any(
                kw in h
                for kw in (
                    "Key ", "Products", "Technologies", "Brands", "Focus",
                    "Categories", "Geography", "Presence", "Region Code",
                )
            ):
                width = 48
            else:
                width = DEFAULT_WIDTH
        ws.column_dimensions[letter].width = width

    # Header row: bold-ish visual weight already set by the source script;
    # just make sure it's tall enough to show wrapped header text.
    ws.row_dimensions[1].height = HEADER_ROW_HEIGHT

    # Every data cell: wrap text + top-align so long values show in full
    # (Excel auto-grows row height to fit wrapped text when a fixed height
    # isn't pinned), and set a sane minimum row height so short rows don't
    # look cramped.
    for r in range(2, ws.max_row + 1):
        for c in range(1, n_cols + 1):
            ws.cell(r, c).alignment = Alignment(vertical="top", horizontal="left", wrap_text=True)
        ws.row_dimensions[r].height = None  # let Excel auto-fit to wrapped content

    ws.freeze_panes = "A2"

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest_path)
    return n_rows, n_cols


def main() -> None:
    DEST_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Exporting verified distributor files to {DEST_DIR}\n")

    total_rows = 0
    for src_rel, dest_name in SOURCES:
        src_path = Path(src_rel)
        if not src_path.exists():
            print(f"  SKIP — not found: {src_path}")
            continue
        dest_path = DEST_DIR / dest_name
        n_rows, n_cols = polish_workbook(src_path, dest_path)
        total_rows += n_rows
        print(f"  {dest_name}: {n_rows} companies, {n_cols} columns")

    print(f"\nDone. {total_rows} verified distributor companies exported across {len(SOURCES)} files.")


if __name__ == "__main__":
    main()
