"""Split Global Multi-Region avocado companies into concrete regional sheets."""
from __future__ import annotations

import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
)

REGIONS = [
    "North America",
    "Europe",
    "Middle East and Africa",
    "Asia-Pacific",
    "Latin America",
]

COL_WIDTHS = {
    "Region Segment": 36,
    "Company": 28,
    "Founded": 10,
    "Headquarters": 28,
    "Cities / Regions": 32,
    "Ownership": 20,
    "Business Type": 26,
    "Employees": 12,
    "Turnover / Scale": 14,
    "Contact Person": 20,
    "Role": 22,
    "Email": 26,
    "Phone / WhatsApp": 16,
    "LinkedIn": 32,
    "Website": 28,
    "Core Categories": 36,
    "Specialty Focus": 38,
    "Price Segment": 20,
    "Key Brands Represented": 36,
    "Retail Chains": 36,
    "Specialty Grocers": 36,
    "Foodservice / HoReCa": 36,
    "E-commerce": 28,
    "Channel Strength": 28,
    "Distribution Type": 28,
    "Regions Served": 28,
    "Regional Extensions": 28,
    "Warehouse / Logistics": 32,
    "Delivery / Storage": 28,
    "Competitive Benchmark": 32,
    "Data Confidence": 12,
    "Quality Score": 12,
    "Data Sources": 40,
    "Summary": 48,
}

# Company-specific multi-region footprints (verified / public)
COMPANY_REGIONS: dict[str, list[str]] = {
    "sysco": ["North America", "Europe", "Latin America"],
    "sysco specialty": ["North America", "Europe", "Latin America"],
    "sysco canada": ["North America"],
    "sysco europe": ["Europe"],
    "brenntag": ["North America", "Europe", "Middle East and Africa", "Asia-Pacific", "Latin America"],
    "brenntag north america": ["North America"],
    "azelis": ["North America", "Europe", "Asia-Pacific", "Latin America", "Middle East and Africa"],
    "azelis americas": ["North America", "Latin America"],
    "barentz": ["North America", "Europe", "Asia-Pacific"],
    "barentz north america": ["North America"],
    "imcd": ["North America", "Europe", "Asia-Pacific", "Latin America", "Middle East and Africa"],
    "imcd us": ["North America"],
    "metro ag": ["Europe", "Middle East and Africa", "Asia-Pacific", "Latin America"],
    "fyffes": ["Europe", "Latin America", "North America"],
    "tradin organic": ["Europe", "North America", "Asia-Pacific", "Latin America", "Middle East and Africa"],
    "univar solutions": ["North America", "Europe", "Asia-Pacific", "Latin America"],
    "olam food ingredients distribution": ["Asia-Pacific", "Europe", "North America", "Middle East and Africa", "Latin America"],
    "itochu food distribution": ["Asia-Pacific", "North America", "Europe"],
    "mitsubishi corporation food distribution": ["Asia-Pacific", "North America", "Europe", "Latin America"],
    "sojitz": ["Asia-Pacific", "North America", "Europe"],
    "fonterra foodservice distribution": ["Asia-Pacific", "Middle East and Africa", "Latin America", "North America"],
    "diverse worldwide trading": ["Europe", "North America", "Asia-Pacific"],
}


def nk(s: str) -> str:
    s = (s or "").lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def infer_regions(company: str, regions_served: str, hq: str, cities: str) -> list[str]:
    key = nk(company)
    if key in COMPANY_REGIONS:
        return list(COMPANY_REGIONS[key])
    for k, regs in COMPANY_REGIONS.items():
        if key.startswith(k) or k in key:
            return list(regs)

    blob = f"{regions_served} {hq} {cities} {company}".lower()
    found: list[str] = []

    if any(x in blob for x in ("united states", "usa", "u.s.", "canada", "north america", "ohio", "illinois", "texas", "houston")):
        found.append("North America")
    if any(
        x in blob
        for x in (
            "europe",
            "uk",
            "united kingdom",
            "germany",
            "netherlands",
            "france",
            "spain",
            "belgium",
            "switzerland",
            "bulgaria",
            "rotterdam",
            "amsterdam",
            "antwerp",
            "essen",
            "düsseldorf",
            "dusseldorf",
            "bologna",
            "plovdiv",
            "geneva",
            "london",
        )
    ):
        found.append("Europe")
    if any(x in blob for x in ("uae", "gcc", "middle east", "africa", "saudi", "dubai", "mena")):
        found.append("Middle East and Africa")
    if any(
        x in blob
        for x in (
            "asia",
            "japan",
            "china",
            "singapore",
            "australia",
            "new zealand",
            "auckland",
            "tokyo",
            "india",
            "malaysia",
        )
    ):
        found.append("Asia-Pacific")
    if any(x in blob for x in ("latin", "latam", "mexico", "brazil", "chile", "colombia", "argentina", "peru", "costa rica", "caribbean", "americas")):
        # "Americas" for Fyffes-style → NA + LatAm if not already
        if "costa rica" in blob or "latin" in blob or "latam" in blob or any(
            x in blob for x in ("brazil", "chile", "colombia", "argentina", "peru")
        ):
            found.append("Latin America")
        elif "americas" in blob:
            if "North America" not in found:
                found.append("North America")
            found.append("Latin America")
        elif "mexico" in blob:
            found.append("Latin America")

    # Dedupe preserve order
    out: list[str] = []
    for r in found:
        if r not in out and r in REGIONS:
            out.append(r)

    if not out:
        # HQ-based single region fallback
        if any(x in blob for x in ("united states", "usa", "canada")):
            out = ["North America"]
        elif any(x in blob for x in ("japan", "singapore", "new zealand", "australia", "china")):
            out = ["Asia-Pacific"]
        elif "costa rica" in blob:
            out = ["Latin America"]
        elif any(x in blob for x in ("europe", "bulgaria", "germany", "netherlands", "switzerland", "belgium")):
            out = ["Europe"]
        else:
            out = ["Europe"]  # last resort for leftover trading houses in this set
    return out


def primary_region(regions: list[str], hq: str) -> str:
    """Pick one primary region for the master row (HQ-biased)."""
    h = (hq or "").lower()
    if any(x in h for x in ("united states", "usa", "canada", "ohio", "illinois", "texas")):
        if "North America" in regions:
            return "North America"
    if any(x in h for x in ("japan", "singapore", "auckland", "new zealand", "australia", "china")):
        if "Asia-Pacific" in regions:
            return "Asia-Pacific"
    if any(x in h for x in ("costa rica", "mexico", "brazil", "chile")):
        if "Latin America" in regions:
            return "Latin America"
    if any(x in h for x in ("dubai", "uae", "saudi", "africa")):
        if "Middle East and Africa" in regions:
            return "Middle East and Africa"
    if any(x in h for x in ("germany", "netherlands", "belgium", "switzerland", "uk", "united kingdom", "france", "spain", "europe", "bulgaria", "italy")):
        if "Europe" in regions:
            return "Europe"
    return regions[0]


def style_header(ws, ncols: int) -> None:
    fill = PatternFill("solid", fgColor="1F4E79")
    font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = border
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "B2"


def style_body(ws, nrows: int, ncols: int, headers: list[str]) -> None:
    font = Font(name="Calibri", size=10)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt = PatternFill("solid", fgColor="F2F2F2")
    for r in range(2, nrows + 1):
        ws.row_dimensions[r].height = 48
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            cell.font = font
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if r % 2 == 0:
                cell.fill = alt
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = COL_WIDTHS.get(str(h), 18)
    ws.auto_filter.ref = ws.dimensions


def rewrite_sheet(wb, title: str, headers: list[str], rows: list[list], index: int | None = None) -> None:
    name = title[:31]
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, index=index)
    for c, h in enumerate(headers, start=1):
        ws.cell(1, c).value = h
    for r, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            ws.cell(r, c).value = val
    style_header(ws, len(headers))
    if rows:
        style_body(ws, len(rows) + 1, len(headers), headers)


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_split_global_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    src = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in src[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    data_rows: list[dict] = []
    for r in range(2, src.max_row + 1):
        row = {h: src.cell(r, cols[h]).value for h in cols}
        company = str(row.get("Company") or "").strip()
        if not company:
            continue
        data_rows.append(row)

    # Split former globals + ensure every row has concrete region list
    by_region: dict[str, list[dict]] = defaultdict(list)
    master_rows: list[dict] = []
    split_log: list[tuple[str, list[str]]] = []

    for row in data_rows:
        company = str(row.get("Company") or "").strip()
        current = str(row.get("Region Segment") or "").strip()
        regs_served = str(row.get("Regions Served") or "")
        hq = str(row.get("Headquarters") or "")
        cities = str(row.get("Cities / Regions") or "")

        if current == "Global Multi-Region" or "Global" in current:
            regions = infer_regions(company, regs_served, hq, cities)
            split_log.append((company, regions))
        elif current in REGIONS:
            regions = [current]
        else:
            # Multi-listed already?
            parts = [p.strip() for p in re.split(r"[;|/]", current) if p.strip()]
            regions = [p for p in parts if p in REGIONS] or infer_regions(company, regs_served, hq, cities)

        primary = primary_region(regions, hq)
        # Master row: show all operating regions
        master = dict(row)
        master["Region Segment"] = "; ".join(regions)
        master["Primary Region"] = primary
        master_rows.append(master)

        for reg in regions:
            copy = dict(row)
            copy["Region Segment"] = reg
            copy["Primary Region"] = primary
            by_region[reg].append(copy)

    # Headers with Primary Region after Region Segment
    base_hdr = [h for h in hdr if h not in ("Region Segment", "Primary Region")]
    new_headers = ["Region Segment", "Primary Region"] + base_hdr

    def to_list(row: dict) -> list:
        return [row.get(h) for h in new_headers]

    # Summary: unique companies per region (row counts on regional sheets)
    unique_total = len(master_rows)
    region_company_counts = {r: len(by_region.get(r, [])) for r in REGIONS}

    # Rebuild sheets
    for name in list(wb.sheetnames):
        if name.startswith("Region -") or name in {"By Region Summary", "All Distributors", "Global Split Note"}:
            del wb[name]

    # Summary
    summary_headers = ["Region Segment", "Companies (rows on sheet)", "% of regional placements", "Notes"]
    summary_rows = []
    placement_total = sum(region_company_counts.values())
    for reg in REGIONS:
        n = region_company_counts[reg]
        summary_rows.append(
            [
                reg,
                n,
                round(100.0 * n / placement_total, 1) if placement_total else 0,
                "Includes global players that operate in this region",
            ]
        )
    summary_rows.append(["UNIQUE COMPANIES (master)", unique_total, 100.0, "Each company listed once in All Distributors"])
    summary_rows.append(
        [
            "FORMER GLOBAL SPLIT",
            len(split_log),
            "",
            "Global Multi-Region companies assigned to concrete regions",
        ]
    )
    rewrite_sheet(wb, "By Region Summary", summary_headers, summary_rows, index=0)

    # Note sheet for split globals
    note_headers = ["Company", "Assigned Regions", "Count"]
    note_rows = [[c, "; ".join(regs), len(regs)] for c, regs in sorted(split_log)]
    rewrite_sheet(wb, "Global Split Note", note_headers, note_rows, index=1)

    # All distributors (unique)
    rewrite_sheet(wb, "All Distributors", new_headers, [to_list(r) for r in master_rows], index=2)

    # Region sheets
    for i, reg in enumerate(REGIONS):
        rows = [to_list(r) for r in by_region.get(reg, [])]
        rewrite_sheet(wb, f"Region - {reg}", new_headers, rows, index=3 + i)

    # Update main Avocado Oil Distributors = unique master
    src = wb["Avocado Oil Distributors"]
    if src.max_column:
        src.delete_cols(1, src.max_column)
    if src.max_row > 1:
        src.delete_rows(2, src.max_row - 1)
    for c, h in enumerate(new_headers, start=1):
        src.cell(1, c).value = h
    for r, row in enumerate(master_rows, start=2):
        for c, val in enumerate(to_list(row), start=1):
            src.cell(r, c).value = val
    style_header(src, len(new_headers))
    style_body(src, len(master_rows) + 1, len(new_headers), new_headers)

    # Remove empty Global Multi-Region sheet if any leftover name
    if "Region - Global Multi-Region" in wb.sheetnames:
        del wb["Region - Global Multi-Region"]

    wb.save(PATH)
    print("backup", bak.name)
    print("unique_companies", unique_total)
    print("split_globals", len(split_log))
    for reg in REGIONS:
        print(f"  {reg}: {region_company_counts[reg]}")
    print("--- split detail ---")
    for c, regs in sorted(split_log):
        print(f"  {c} -> {', '.join(regs)}")
    print("sheets", wb.sheetnames)


if __name__ == "__main__":
    main()
