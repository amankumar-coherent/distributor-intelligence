"""Segment green chemical FINAL Excel by macro region (HQ-based)."""
from __future__ import annotations

import re
import shutil
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PATHS = [
    Path(
        r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market_FINAL.xlsx"
    ),
    Path(
        r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market_FINAL_320plus.xlsx"
    ),
]

# Country / city hints → macro region
REGION_RULES: list[tuple[str, list[str]]] = [
    (
        "North America",
        [
            "united states",
            "usa",
            "u.s.",
            "canada",
            "mexico",
            "california",
            "texas",
            "new york",
            "illinois",
            "florida",
            "harrisburg",
            "irvine",
            "san jose",
            "rosemont",
            "houston",
            "chicago",
            "toronto",
            "ontario",
            "vancouver",
            "pennsylvania",
            "virginia",
            "new jersey",
            "north america",
        ],
    ),
    (
        "Europe",
        [
            "united kingdom",
            "uk",
            "germany",
            "france",
            "italy",
            "spain",
            "netherlands",
            "poland",
            "sweden",
            "norway",
            "denmark",
            "finland",
            "belgium",
            "switzerland",
            "austria",
            "ireland",
            "portugal",
            "greece",
            "czech",
            "slovakia",
            "hungary",
            "bulgaria",
            "romania",
            "croatia",
            "slovenia",
            "latvia",
            "lithuania",
            "estonia",
            "riga",
            "europe",
            "eu ",
            "london",
            "berlin",
            "paris",
            "amsterdam",
            "warsaw",
            "milan",
            "madrid",
            "hoofddorp",
            "antwerp",
            "offenbach",
            "essen",
            "hamburg",
            "athens",
            "budapest",
            "prague",
            "sofia",
        ],
    ),
    (
        "Asia Pacific",
        [
            "india",
            "china",
            "japan",
            "singapore",
            "australia",
            "new zealand",
            "south korea",
            "korea",
            "thailand",
            "vietnam",
            "indonesia",
            "malaysia",
            "philippines",
            "taiwan",
            "hong kong",
            "asia",
            "apac",
            "asia pacific",
            "mumbai",
            "delhi",
            "noida",
            "bangalore",
            "bengaluru",
            "gurugram",
            "gurgaon",
            "thane",
            "maharashtra",
            "pune",
            "hyderabad",
            "chennai",
            "shanghai",
            "beijing",
            "tokyo",
            "sydney",
            "melbourne",
            "auckland",
            "bangkok",
            "jakarta",
            "manila",
            "ho chi minh",
            "hanoi",
        ],
    ),
    (
        "Middle East",
        [
            "united arab emirates",
            "uae",
            "dubai",
            "abu dhabi",
            "saudi",
            "arabia",
            "qatar",
            "kuwait",
            "bahrain",
            "oman",
            "middle east",
            "gcc",
            "israel",
            "turkey",
            "türkiye",
        ],
    ),
    (
        "Latin America",
        [
            "brazil",
            "chile",
            "colombia",
            "argentina",
            "peru",
            "uruguay",
            "montevideo",
            "latin america",
            "latam",
            "sao paulo",
            "são paulo",
            "mexico city",  # also NA — Mexico already in NA; keep LatAm for south cone
        ],
    ),
    (
        "Africa",
        [
            "south africa",
            "nigeria",
            "egypt",
            "kenya",
            "morocco",
            "africa",
            "johannesburg",
            "cape town",
            "lagos",
            "cairo",
        ],
    ),
    (
        "Eurasia / CIS",
        [
            "russia",
            "moscow",
            "saint petersburg",
            "st. petersburg",
            "vladivostok",
            "krasnogorsk",
            "cis",
            "kazakhstan",
            "belarus",
            "ukraine",
            "uzbekistan",
            "eurasia",
        ],
    ),
]

REGION_ORDER = [
    "Global",
    "North America",
    "Europe",
    "Asia Pacific",
    "Middle East",
    "Latin America",
    "Africa",
    "Eurasia / CIS",
    "Other - Unclear",
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
CELL_FONT = Font(name="Calibri", size=9)
THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)
WRAP = Alignment(wrap_text=True, vertical="top")


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def infer_region(hq: str, cities: str, regions_served: str, extensions: str) -> str:
    served = _norm(regions_served)
    if served in {"global", "worldwide", "international", "multi-region", "multi region"}:
        # still prefer HQ for primary segment if HQ is clear
        pass

    blob = " | ".join([hq, cities, extensions, regions_served])
    low = _norm(blob)

    # Explicit Global primary if Regions Served is Global AND HQ empty/unclear
    scores: dict[str, int] = defaultdict(int)
    for region, hints in REGION_RULES:
        for h in hints:
            if h in low:
                # HQ/cities weigh more than served
                w = 3 if h in _norm(hq) or h in _norm(cities) else 1
                scores[region] += w

    if scores:
        best = max(scores.items(), key=lambda x: x[1])
        return best[0]

    # Do not mark as Global solely because Regions Served says Global —
    # require HQ also empty/global; otherwise leave unclear for manual fix.
    hq_low = _norm(hq)
    if served in {"global", "worldwide", "international"} and (
        not hq_low or hq_low in {"global", "worldwide", "international"}
    ):
        return "Global"
    return "Other - Unclear"


def _style_header(ws, headers: list[str]) -> None:
    for i, h in enumerate(headers, 1):
        cell = ws.cell(1, i, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = 18 if h != "Summary" else 36
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"


def _write_rows(ws, headers: list[str], rows: list[dict[str, str]]) -> None:
    _style_header(ws, headers)
    for r in rows:
        vals = [r.get(h, "") for h in headers]
        ws.append(vals)
        rr = ws.max_row
        for i in range(1, len(headers) + 1):
            cell = ws.cell(rr, i)
            cell.font = CELL_FONT
            cell.border = THIN
            cell.alignment = WRAP
        ws.row_dimensions[rr].height = 32
    last = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last}{ws.max_row}"


def segment_workbook(path: Path) -> dict:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.with_name(f"{path.stem}_pre_region_seg_{stamp}{path.suffix}")
    shutil.copy2(path, bak)

    wb = openpyxl.load_workbook(path)
    ws = wb["Green Chemical Companies"]

    # Header is row 3 (title rows above)
    hdr = [str(c.value or "") for c in ws[3]]
    if "Company" not in hdr:
        raise RuntimeError("Company header not found on row 3")

    # Load all company rows
    companies: list[dict[str, str]] = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row or not row[0]:
            continue
        d = {hdr[i]: str(row[i] if i < len(row) and row[i] is not None else "") for i in range(len(hdr))}
        companies.append(d)

    # Infer region
    for d in companies:
        d["Region Segment"] = infer_region(
            d.get("Headquarters", ""),
            d.get("Cities / Regions", ""),
            d.get("Regions Served", ""),
            d.get("Regional Extensions", ""),
        )

    # Insert Region Segment column after Company (col 2) on main sheet
    if "Region Segment" in hdr:
        # already present — refresh values
        seg_col = hdr.index("Region Segment") + 1
        co_col = hdr.index("Company") + 1
        for r_i, d in enumerate(companies, start=4):
            ws.cell(r_i, seg_col).value = d["Region Segment"]
        new_headers = hdr
    else:
        insert_at = 2  # after Company
        ws.insert_cols(insert_at)
        ws.cell(3, insert_at).value = "Region Segment"
        cell = ws.cell(3, insert_at)
        cell.font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
        for r_i, d in enumerate(companies, start=4):
            ws.cell(r_i, insert_at).value = d["Region Segment"]
        new_headers = hdr[:1] + ["Region Segment"] + hdr[1:]

    # Rebuild region sheets
    for name in list(wb.sheetnames):
        if name.startswith("Region - ") or name in {"Region Summary", "By Region"}:
            del wb[name]

    # Summary sheet
    counts = Counter(d["Region Segment"] for d in companies)
    sum_ws = wb.create_sheet("Region Summary", 1)
    sum_ws["A1"] = "Global Green Chemical Market — Region-wise segmentation"
    sum_ws["A1"].font = Font(bold=True, size=14, color="1F4E79")
    sum_ws.merge_cells("A1:D1")
    sum_ws["A3"] = "Region Segment"
    sum_ws["B3"] = "Company Count"
    sum_ws["C3"] = "% of Total"
    sum_ws["D3"] = "Notes"
    for col in range(1, 5):
        c = sum_ws.cell(3, col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = THIN
    total = len(companies)
    r = 4
    for region in REGION_ORDER:
        n = counts.get(region, 0)
        if n == 0 and region == "Other - Unclear":
            continue
        if n == 0:
            continue
        sum_ws.cell(r, 1, region).border = THIN
        sum_ws.cell(r, 2, n).border = THIN
        sum_ws.cell(r, 3, f"{100.0 * n / total:.1f}%").border = THIN
        note = {
            "Global": "Regions Served = Global / multi-region (HQ may still be local)",
            "North America": "USA, Canada, Mexico",
            "Europe": "UK + EU / wider Europe (incl. Baltics, CEE, Greece)",
            "Asia Pacific": "India, China, Japan, SEA, ANZ, Korea",
            "Middle East": "GCC, UAE, Saudi, Turkey, Israel",
            "Latin America": "Brazil, Chile, Colombia, Argentina, Peru, Uruguay",
            "Africa": "South Africa, Nigeria, Egypt, Kenya, etc.",
            "Eurasia / CIS": "Russia / CIS channel partners",
            "Other - Unclear": "Insufficient HQ/region text to classify",
        }.get(region, "")
        sum_ws.cell(r, 4, note).border = THIN
        r += 1
    sum_ws.cell(r, 1, "TOTAL").font = Font(bold=True)
    sum_ws.cell(r, 2, total).font = Font(bold=True)
    sum_ws.column_dimensions["A"].width = 22
    sum_ws.column_dimensions["B"].width = 16
    sum_ws.column_dimensions["C"].width = 12
    sum_ws.column_dimensions["D"].width = 55

    # Country rollup from HQ
    sum_ws.cell(r + 2, 1, "Top Headquarters / Countries").font = Font(bold=True, color="1F4E79")
    sum_ws.cell(r + 3, 1, "HQ / Country text")
    sum_ws.cell(r + 3, 2, "Count")
    for col in range(1, 3):
        c = sum_ws.cell(r + 3, col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    hq_counts = Counter()
    for d in companies:
        hq = (d.get("Headquarters") or "").strip() or (d.get("Regions Served") or "").strip() or "Unknown"
        # simplify to country-ish first token group
        hq_counts[hq[:80]] += 1
    rr = r + 4
    for hq, n in hq_counts.most_common(25):
        sum_ws.cell(rr, 1, hq)
        sum_ws.cell(rr, 2, n)
        rr += 1

    # Per-region sheets
    by_region: dict[str, list[dict[str, str]]] = defaultdict(list)
    for d in companies:
        by_region[d["Region Segment"]].append(d)

    for region in REGION_ORDER:
        rows = by_region.get(region) or []
        if not rows:
            continue
        rows = sorted(rows, key=lambda x: _norm(x.get("Company", "")))
        safe = region.replace("/", "-")
        title = f"Region - {safe}"[:31]
        rws = wb.create_sheet(title)
        _write_rows(rws, new_headers, rows)

    wb.save(path)
    return {
        "file": str(path),
        "backup": str(bak),
        "total": total,
        "counts": dict(counts),
    }


def main() -> None:
    for p in PATHS:
        rep = segment_workbook(p)
        print("FILE", Path(rep["file"]).name)
        print(" backup", Path(rep["backup"]).name)
        print(" total", rep["total"])
        for k in REGION_ORDER:
            if k in rep["counts"]:
                print(f"  {k}: {rep['counts'][k]}")
        print()


if __name__ == "__main__":
    main()
