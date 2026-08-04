"""Format avocado oil FINAL so full cell text is visible (wrap + row heights)."""
from __future__ import annotations

import math
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
)

# Preferred display widths (characters). Long-text cols get wrap + taller rows.
COL_WIDTHS: dict[str, float] = {
    "Company": 28,
    "Founded": 10,
    "Headquarters": 28,
    "Cities / Regions": 36,
    "Ownership": 22,
    "Business Type": 28,
    "Employees": 12,
    "Turnover / Scale": 16,
    "Contact Person": 22,
    "Role": 24,
    "Email": 28,
    "Phone / WhatsApp": 18,
    "LinkedIn": 36,
    "Website": 30,
    "Core Categories": 40,
    "Specialty Focus": 42,
    "Price Segment": 22,
    "Key Brands Represented": 40,
    "Exclusivity": 18,
    "Partnership Duration": 16,
    "Retail Chains": 14,
    "Specialty Grocers": 14,
    "Foodservice / HoReCa": 16,
    "E-commerce": 14,
    "Channel Strength": 16,
    "Distribution Type": 22,
    "Regions Served": 18,
    "Regional Extensions": 18,
    "Warehouse / Logistics": 18,
    "Delivery / Storage": 18,
    "Competitive Benchmark": 18,
    "Data Confidence": 12,
    "Quality Score": 12,
    "Data Sources": 48,
    "Summary": 55,
}

WRAP_COLS = {
    "Headquarters",
    "Cities / Regions",
    "Ownership",
    "Business Type",
    "Role",
    "LinkedIn",
    "Website",
    "Core Categories",
    "Specialty Focus",
    "Price Segment",
    "Key Brands Represented",
    "Exclusivity",
    "Distribution Type",
    "Regions Served",
    "Regional Extensions",
    "Warehouse / Logistics",
    "Delivery / Storage",
    "Competitive Benchmark",
    "Data Sources",
    "Summary",
    "Contact Person",
    "Email",
}


def _lines_needed(text: str, width: float) -> int:
    if not text:
        return 1
    w = max(8, int(width))
    lines = 0
    for para in str(text).splitlines() or [""]:
        if not para:
            lines += 1
            continue
        lines += max(1, math.ceil(len(para) / w))
    return max(1, lines)


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_format_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    body_font = Font(name="Calibri", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill = PatternFill("solid", fgColor="F2F2F2")

    # Column widths
    for h, idx in cols.items():
        letter = get_column_letter(idx)
        width = COL_WIDTHS.get(str(h), 18)
        ws.column_dimensions[letter].width = width

    # Header
    ws.row_dimensions[1].height = 32
    for idx, h in enumerate(hdr, start=1):
        cell = ws.cell(1, idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = border

    # Body: wrap + estimate row height from longest wrapped column
    for r in range(2, ws.max_row + 1):
        max_lines = 1
        for h, idx in cols.items():
            cell = ws.cell(r, idx)
            cell.font = body_font
            cell.border = border
            if r % 2 == 0:
                if cell.fill.fgColor is None or cell.fill.fgColor.rgb in (None, "00000000"):
                    cell.fill = alt_fill
            wrap = str(h) in WRAP_COLS
            cell.alignment = Alignment(
                wrap_text=wrap or True,
                vertical="top",
                horizontal="left",
            )
            width = COL_WIDTHS.get(str(h), 18)
            val = cell.value
            if val is not None:
                max_lines = max(max_lines, _lines_needed(str(val), width))

        # Cap row height so extreme Summary/Data Sources don't blow the sheet
        height = min(220, max(18, 14 * max_lines + 4))
        ws.row_dimensions[r].height = height

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    # Print / view niceties
    ws.sheet_view.showGridLines = True
    ws.print_title_rows = "1:1"

    wb.save(PATH)
    print("backup", bak.name)
    print("formatted", PATH.name, "rows", ws.max_row - 1, "cols", len(hdr))
    print("wrap_enabled=True freeze=B2 autofilter=on row_heights_set=True")


if __name__ == "__main__":
    main()
