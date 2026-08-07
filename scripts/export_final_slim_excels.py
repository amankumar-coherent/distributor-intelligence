#!/usr/bin/env python3
"""Create SLIM copies of market FINAL Excels with selected columns removed.

Does NOT modify/delete the original FINAL files.
Output: <same_folder>/<stem>_SLIM.xlsx
"""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(r"D:\vendor-intel\output")

# Prefer the most complete FINAL for each market
SOURCES: list[Path] = [
    ROOT / "run4/global_smartwatch_market_global/global_smartwatch_market_FINAL_320plus.xlsx",
    ROOT / "run4/global_green_chemical_market_global/global_green_chemical_market_FINAL_320plus.xlsx",
    ROOT / "run4/global_firewall_market_global/global_firewall_market_FINAL_320plus.xlsx",
    ROOT / "web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL.xlsx",
    ROOT / "web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL.xlsx",
    ROOT / "web_expand/global_antibody_drug_conjugates_market_global/global_antibody_drug_conjugates_market_FINAL.xlsx",
]

# Fallback if 320plus missing
FALLBACKS = {
    "global_smartwatch_market_FINAL_320plus.xlsx": "global_smartwatch_market_FINAL.xlsx",
    "global_green_chemical_market_FINAL_320plus.xlsx": "global_green_chemical_market_FINAL.xlsx",
    "global_firewall_market_FINAL_320plus.xlsx": "global_firewall_market_FINAL.xlsx",
}

# Columns to remove (normalized match)
DROP_NORMS = {
    "channel strength",
    "warehouse / logistics",
    "warehouse/logistics",
    "warehouse logistics",
    "delivery / storage",
    "delivery/storage",
    "delivery storage",
    "competitive benchmark",
    "cities / regions",
    "cities/regions",
    "cities / region",
    "cities/region",
    "business type",
    "price segment",
    "contact person",
    "role",
    "linkedin",
    "email",
    "phone / whatsapp",
    "phone/whatsapp",
    "contact no",
    "contact number",
    "phone",
    "whatsapp",
    "turnover",
    "turnover / scale",
    "turnover/scale",
    "annual turnover",
    "revenue / turnover",
    "revenue/turnover",
    "region segment",
    "data sources",
    "data source",
    "data confidence",
    "quality score",
}

# Drop for most markets; keep for avocado oil (food channel is relevant)
FOOD_CHANNEL_DROP = {
    "specialty grocers",
    "specialty grocer",
    "foodservice / horeca",
    "foodservice/horeca",
    "foodservice",
    "horeca",
}


def norm_header(h: object) -> str:
    s = str(h or "").strip().lower()
    s = s.replace("–", "/").replace("—", "/")
    s = re.sub(r"\s+", " ", s)
    return s


def should_drop(header: object, *, keep_food_channels: bool = False) -> bool:
    n = norm_header(header)
    if n in FOOD_CHANNEL_DROP:
        return not keep_food_channels
    if n in DROP_NORMS:
        return True
    # soft aliases
    if n.startswith("phone") and "whatsapp" in n:
        return True
    if n.startswith("turnover"):
        return True
    if n in {"contact", "contacts"}:
        return False
    return False


def is_channel_present(val: object) -> bool:
    """True when Retail Chains / E-commerce cell indicates presence."""
    s = str(val or "").strip().lower()
    if not s:
        return False
    s = re.sub(r"\s+", " ", s)
    # explicit negatives / N/A
    if s in {"no", "n", "none", "n/a", "na", "nil", "-", "—", "false", "0"}:
        return False
    if s.startswith("n/a") or s.startswith("not applicable") or s.startswith("not applicable"):
        return False
    if s.startswith("no ") or s.startswith("no—") or s.startswith("no -") or s.startswith("no–"):
        return False
    if "not applicable" in s or s.startswith("not a "):
        return False
    # explicit yes / partial presence
    if s in {"yes", "y", "true", "1"}:
        return True
    if s.startswith("yes") or s.startswith("partial") or s.startswith("moderate") or s.startswith("core"):
        return True
    # any other descriptive text = present
    return True


def channel_combo(retail_val: object, ecom_val: object) -> str:
    retail = is_channel_present(retail_val)
    ecom = is_channel_present(ecom_val)
    if retail and ecom:
        return "Both"
    if retail:
        return "Retail"
    if ecom:
        return "Ecommerce"
    return "No"


def is_retail_header(h: object) -> bool:
    n = norm_header(h)
    return n in {"retail chains", "retail chain", "retail"}


def is_ecommerce_header(h: object) -> bool:
    n = norm_header(h)
    return n in {"e-commerce", "ecommerce", "e commerce"}


MERGED_CHANNEL_HEADER = "Retail / E-commerce / Both"

APPEND_OTHERS_HEADERS = {
    "core categories",
    "specialty focus",
    "key brands represented",
    "key brands",
}


def append_others(val: object) -> object:
    """Append ', others' at end of text if not already present."""
    if val in (None, ""):
        return val
    s = str(val).rstrip()
    if not s:
        return val
    low = s.lower()
    if low.endswith(", others") or low.endswith(",others") or low.endswith(" others"):
        return s
    return f"{s}, others"


def should_append_others(header: str) -> bool:
    return norm_header(header) in APPEND_OTHERS_HEADERS


def detect_header_row(ws) -> int:
    """Return 1-based header row. Supports title / blank / header layouts."""
    for r in range(1, min(6, ws.max_row + 1)):
        vals = [norm_header(c.value) for c in ws[r] if c.value not in (None, "")]
        if "company" in vals or "brand" in vals:
            return r
    # fallback: densest early row
    best_r, best_n = 1, 0
    for r in range(1, min(6, ws.max_row + 1)):
        n = sum(1 for c in ws[r] if c.value not in (None, ""))
        if n > best_n:
            best_n, best_r = n, r
    return best_r


def pick_data_sheet(wb):
    """Prefer the main companies sheet over Region Summary / Audit sheets."""
    preferred = []
    for name in wb.sheetnames:
        low = name.lower()
        if any(x in low for x in ("companies", "distributors", "partners")):
            preferred.append(name)
        elif low in {"sheet1", "data"}:
            preferred.append(name)
    if preferred:
        return wb[preferred[0]]
    # else first sheet that has a Company header in first 5 rows
    for name in wb.sheetnames:
        ws = wb[name]
        for r in range(1, min(6, ws.max_row + 1)):
            vals = [norm_header(c.value) for c in ws[r] if c.value not in (None, "")]
            if "company" in vals:
                return ws
    return wb.active


def slim_one(src: Path) -> tuple:
    if not src.exists():
        fb = FALLBACKS.get(src.name)
        if fb:
            alt = src.with_name(fb)
            if alt.exists():
                src = alt
        if not src.exists():
            raise FileNotFoundError(src)

    wb = load_workbook(src, data_only=True)
    ws = pick_data_sheet(wb)
    header_row = detect_header_row(ws)
    headers = [c.value for c in ws[header_row]]
    keep_food_channels = "avocado" in src.name.lower()

    retail_idx: int | None = None
    ecom_idx: int | None = None
    keep_idx: list[int] = []
    keep_headers: list[str] = []
    dropped: list[str] = []
    merged_inserted = False

    for i, h in enumerate(headers):
        if h in (None, ""):
            continue
        if is_retail_header(h):
            retail_idx = i
            dropped.append(str(h))
            if not merged_inserted:
                keep_headers.append(MERGED_CHANNEL_HEADER)
                keep_idx.append(-1)  # placeholder for merged column
                merged_inserted = True
            continue
        if is_ecommerce_header(h):
            ecom_idx = i
            dropped.append(str(h))
            if not merged_inserted:
                keep_headers.append(MERGED_CHANNEL_HEADER)
                keep_idx.append(-1)
                merged_inserted = True
            continue
        if should_drop(h, keep_food_channels=keep_food_channels):
            dropped.append(str(h))
            continue
        keep_idx.append(i)
        keep_headers.append(str(h))

    if (retail_idx is not None or ecom_idx is not None) and not merged_inserted:
        keep_headers.append(MERGED_CHANNEL_HEADER)
        keep_idx.append(-1)

    # Put Website immediately after Company
    company_pos = next((i for i, h in enumerate(keep_headers) if norm_header(h) == "company"), None)
    website_pos = next((i for i, h in enumerate(keep_headers) if norm_header(h) == "website"), None)
    if company_pos is not None and website_pos is not None and website_pos != company_pos + 1:
        wh = keep_headers.pop(website_pos)
        wi = keep_idx.pop(website_pos)
        insert_at = company_pos + 1 if website_pos > company_pos else company_pos
        keep_headers.insert(insert_at, wh)
        keep_idx.insert(insert_at, wi)

    out = Workbook()
    ows = out.active
    # Preserve a simple title if source had one above the header
    title = None
    if header_row > 1 and ws["A1"].value:
        title = str(ws["A1"].value)
    start_row = 1
    if title:
        ows.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(keep_headers)))
        cell = ows.cell(1, 1, title if "SLIM" in title.upper() else f"{title} — SLIM")
        cell.font = Font(bold=True, size=13, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        start_row = 2

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2E75B6")
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for col, h in enumerate(keep_headers, 1):
        cell = ows.cell(start_row, col, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin
        ows.column_dimensions[get_column_letter(col)].width = min(36, max(14, len(h) + 2))

    data_start = header_row + 1
    out_r = start_row + 1
    n_rows = 0
    for r in range(data_start, ws.max_row + 1):
        vals: list[object] = []
        for i in keep_idx:
            if i == -1:
                retail_v = ws.cell(r, retail_idx + 1).value if retail_idx is not None else None
                ecom_v = ws.cell(r, ecom_idx + 1).value if ecom_idx is not None else None
                vals.append(channel_combo(retail_v, ecom_v))
            else:
                vals.append(ws.cell(r, i + 1).value)
        # skip fully empty rows
        if all(v in (None, "") for v in vals):
            continue
        # skip if first kept col empty and looks like company column
        if keep_headers and norm_header(keep_headers[0]) in {"company", "brand"} and vals and vals[0] in (None, ""):
            continue
        for c, v in enumerate(vals, 1):
            if should_append_others(keep_headers[c - 1]):
                v = append_others(v)
            cell = ows.cell(out_r, c, v if v is not None else "")
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if out_r % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")
        out_r += 1
        n_rows += 1

    ows.auto_filter.ref = ows.dimensions
    ows.freeze_panes = ows.cell(start_row + 1, 1).coordinate

    # Output name: replace FINAL / FINAL_320plus with FINAL_SLIM
    stem = src.stem
    if stem.endswith("_FINAL_320plus"):
        out_stem = stem.replace("_FINAL_320plus", "_FINAL_SLIM")
    elif stem.endswith("_FINAL"):
        out_stem = stem.replace("_FINAL", "_FINAL_SLIM")
    elif "_FINAL" in stem:
        out_stem = stem + "_SLIM"
    else:
        out_stem = stem + "_SLIM"
    dest = src.with_name(out_stem + ".xlsx")
    out.save(dest)

    return dest, n_rows, keep_headers, dropped, src


def main() -> None:
    print("Creating SLIM Excel copies (originals untouched)...\n")
    for src in SOURCES:
        try:
            dest, n_rows, keep, dropped, used = slim_one(src)
        except FileNotFoundError as e:
            print(f"MISSING: {e}")
            continue
        print(f"OK  {used.name}")
        print(f" -> {dest.name}")
        print(f"    rows={n_rows}  kept_cols={len(keep)}  dropped_cols={len(dropped)}")
        print(f"    dropped: {', '.join(dropped)}")
        if MERGED_CHANNEL_HEADER in keep:
            print(f"    merged: Retail Chains + E-commerce -> {MERGED_CHANNEL_HEADER}")
        print()


if __name__ == "__main__":
    main()
