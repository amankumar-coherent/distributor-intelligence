"""Resolve smartwatch 'Other - Unclear' rows: verify region + fill firmographics."""
from __future__ import annotations

import shutil
from collections import Counter, defaultdict
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PATHS = [
    Path(r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.xlsx"),
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL_320plus.xlsx"
    ),
]

# Verified fills for previously unclassified rows (agent web research)
FIXES: dict[str, dict[str, str]] = {
    "cover company": {
        "Region Segment": "Latin America",
        "Headquarters": "Montevideo, Uruguay",
        "Cities / Regions": "Montevideo; Maldonado; Ciudad de la Costa, Uruguay",
        "Regions Served": "Uruguay",
        "Regional Extensions": "Montevideo, Uruguay",
        "Founded": "2014",
        "Employees": "11-50",
        "Business Type": "Retailer / Authorized Distributor",
        "Ownership": "Privately held",
        "Phone / WhatsApp": "+598 2901 6984",
        "Website": "https://covercompany.com.uy",
        "LinkedIn": "https://www.linkedin.com/company/cover-company-uruguay",
        "Core Categories": "Smartwatches; smartphones; consumer electronics",
        "Specialty Focus": "Uruguay electronics retail & distribution (Amazfit and wearables)",
        "Key Brands Represented": "Amazfit; Xiaomi; Samsung; Apple; DJI; Insta360; Shokz",
        "E-commerce": "Yes",
        "Retail Chains": "Yes",
        "Distribution Type": "Retail + wholesale channel",
        "Summary": "Cover Company — Uruguay electronics retailer/distributor (est. 2014, Montevideo); sells smartwatches & wearables.",
    },
    "wildberries": {
        "Region Segment": "Eurasia / CIS",
        "Headquarters": "Moscow, Russia",
        "Cities / Regions": "Moscow, Russia",
        "Regions Served": "Russia; CIS",
        "Regional Extensions": "Russia; CIS",
        "Business Type": "E-commerce marketplace / retailer",
        "Website": "https://www.wildberries.ru",
        "Summary": "Wildberries — major Russia/CIS e-commerce marketplace carrying smartwatches & wearables.",
    },
    "marvel distribution": {
        "Region Segment": "Eurasia / CIS",
        "Headquarters": "Moscow / Saint Petersburg, Russia",
        "Cities / Regions": "Moscow; Saint Petersburg, Russia",
        "Regions Served": "Russia; CIS",
        "Regional Extensions": "Russia; CIS",
        "Founded": "1991",
        "Business Type": "IT Distributor",
        "Website": "https://www.marvel.ru",
        "Specialty Focus": "Broadline IT distribution (Russia/CIS) including consumer devices",
        "Summary": "Marvel Distribution — large Russia/CIS IT distributor (founded 1991); offices in Moscow & St. Petersburg.",
    },
    "merlion": {
        "Region Segment": "Eurasia / CIS",
        "Headquarters": "Krasnogorsk / Moscow region, Russia",
        "Cities / Regions": "Moscow region, Russia",
        "Regions Served": "Russia; CIS",
        "Regional Extensions": "Russia; CIS",
        "Business Type": "IT Distributor",
        "Website": "https://merlion.ru",
        "Specialty Focus": "Broadline IT distribution — Russia/CIS channel",
        "Summary": "MERLION — leading Russia IT distributor / group; channel for consumer electronics & wearables.",
    },
    "ocs distribution": {
        "Region Segment": "Eurasia / CIS",
        "Headquarters": "Saint Petersburg, Russia",
        "Cities / Regions": "Saint Petersburg; regional offices across Russia",
        "Regions Served": "Russia; CIS",
        "Regional Extensions": "Russia; CIS",
        "Business Type": "IT Distributor",
        "Website": "https://ocs.ru",
        "Specialty Focus": "Project & retail IT distribution in Russia",
        "Summary": "OCS Distribution — major Russian IT distributor (HQ St. Petersburg).",
    },
    "citilink": {
        "Region Segment": "Eurasia / CIS",
        "Headquarters": "Moscow, Russia",
        "Cities / Regions": "Russia",
        "Regions Served": "Russia",
        "Regional Extensions": "Russia",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.citilink.ru",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "Citilink — Russian consumer electronics retail chain (smartwatches/wearables).",
    },
    "dns retail": {
        "Region Segment": "Eurasia / CIS",
        "Headquarters": "Vladivostok / Russia",
        "Cities / Regions": "Russia; CIS",
        "Regions Served": "Russia; CIS",
        "Regional Extensions": "Russia; CIS",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.dns-shop.ru",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "DNS Retail — large Russia/CIS electronics retail chain.",
    },
    "m.video-eldorado": {
        "Region Segment": "Eurasia / CIS",
        "Headquarters": "Moscow, Russia",
        "Cities / Regions": "Russia",
        "Regions Served": "Russia; CIS",
        "Regional Extensions": "Russia; CIS",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.mvideo.ru",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "M.Video-Eldorado — leading Russian consumer electronics retail group.",
    },
    "datart": {
        "Region Segment": "Europe",
        "Headquarters": "Czech Republic",
        "Cities / Regions": "Czech Republic; Slovakia",
        "Regions Served": "Czech Republic; Slovakia",
        "Regional Extensions": "Czech Republic; Slovakia",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.datart.cz",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "Datart — Czech/Slovak consumer electronics retailer (wearables channel).",
    },
    "kotsovolos": {
        "Region Segment": "Europe",
        "Headquarters": "Greece",
        "Cities / Regions": "Greece",
        "Regions Served": "Greece",
        "Regional Extensions": "Greece",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.kotsovolos.gr",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "Kotsovolos — major Greek consumer electronics retailer.",
    },
    "plaisio": {
        "Region Segment": "Europe",
        "Headquarters": "Greece",
        "Cities / Regions": "Greece",
        "Regions Served": "Greece",
        "Regional Extensions": "Greece",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.plaisio.gr",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "Plaisio — Greek electronics & IT retailer (smartwatches).",
    },
    "public": {
        "Region Segment": "Europe",
        "Headquarters": "Greece",
        "Cities / Regions": "Greece; Cyprus",
        "Regions Served": "Greece; Cyprus",
        "Regional Extensions": "Greece; Cyprus",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.public.gr",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "Public — Greek/Cypriot electronics retail chain.",
    },
    "mediamarkt hungary": {
        "Region Segment": "Europe",
        "Headquarters": "Hungary",
        "Cities / Regions": "Hungary",
        "Regions Served": "Hungary",
        "Regional Extensions": "Hungary",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.mediamarkt.hu",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "MediaMarkt Hungary — consumer electronics retail (MediaMarkt group).",
    },
    "technomarket": {
        "Region Segment": "Europe",
        "Headquarters": "Bulgaria",
        "Cities / Regions": "Bulgaria",
        "Regions Served": "Bulgaria",
        "Regional Extensions": "Bulgaria",
        "Business Type": "Electronics Retailer",
        "Website": "https://www.technomarket.bg",
        "Retail Chains": "Yes",
        "E-commerce": "Yes",
        "Summary": "Technomarket — Bulgarian consumer electronics retailer.",
    },
    "elko group": {
        "Region Segment": "Europe",
        "Headquarters": "Riga, Latvia",
        "Cities / Regions": "Riga, Latvia; Baltics; CEE",
        "Regions Served": "Latvia; Estonia; Lithuania; Central & Eastern Europe; Central Asia",
        "Regional Extensions": "Baltics; CEE; Kazakhstan",
        "Founded": "1993",
        "Employees": "1001-5000",
        "Business Type": "IT Distributor",
        "Ownership": "Public (Nasdaq Baltic)",
        "Website": "https://www.elkogroup.com",
        "LinkedIn": "https://www.linkedin.com/company/elko-group",
        "Core Categories": "IT; consumer electronics; smartphones; wearables",
        "Specialty Focus": "Broadline IT & CE distribution (Baltics/CEE)",
        "Key Brands Represented": "Apple; Samsung; Xiaomi; Lenovo; ASUS; Dell",
        "Turnover": "$1–1.5B",
        "Summary": "ELKO Group (ELKO Grupa AS) — Latvian IT/CE distributor founded 1993 in Riga; ~1,100+ employees; operates across Baltics, CEE & Central Asia.",
    },
}

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
    return " ".join((s or "").lower().split())


def _match(name: str) -> dict[str, str] | None:
    key = _norm(name)
    return FIXES.get(key)


def _style_header(ws, headers: list[str]) -> None:
    for i, h in enumerate(headers, 1):
        cell = ws.cell(1, i, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = THIN
        ws.column_dimensions[get_column_letter(i)].width = 18 if h != "Summary" else 36
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28


def _write_company_sheet(ws, headers: list[str], rows: list[dict[str, str]]) -> None:
    # clear
    ws.delete_rows(1, ws.max_row)
    _style_header(ws, headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
        rr = ws.max_row
        for i in range(1, len(headers) + 1):
            cell = ws.cell(rr, i)
            cell.font = CELL_FONT
            cell.border = THIN
            cell.alignment = WRAP
        ws.row_dimensions[rr].height = 32
    last = get_column_letter(len(headers))
    if rows:
        ws.auto_filter.ref = f"A1:{last}{ws.max_row}"


def fix_file(path: Path) -> dict:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.with_name(f"{path.stem}_pre_unclear_fix_{stamp}{path.suffix}")
    shutil.copy2(path, bak)

    wb = openpyxl.load_workbook(path)
    ws = wb["Smartwatch Companies"]
    hdr = [str(c.value or "") for c in ws[3]]
    assert "Company" in hdr and "Region Segment" in hdr

    updated = []
    companies: list[dict[str, str]] = []
    for r_i, row in enumerate(ws.iter_rows(min_row=4, values_only=False), start=4):
        vals = [c.value for c in row[: len(hdr)]]
        if not vals or not vals[0]:
            continue
        d = {hdr[i]: str(vals[i] if vals[i] is not None else "") for i in range(len(hdr))}
        patch = _match(d["Company"])
        if patch:
            for k, v in patch.items():
                if k not in hdr:
                    continue
                # always overwrite region; for other fields overwrite if empty/NPD/weak
                cur = d.get(k, "")
                weak = cur.strip().lower() in {
                    "",
                    "n/a",
                    "not publicly disclosed",
                    "unknown",
                    "other - unclear",
                    "russia",
                    "moscow",
                    "greece",
                    "hungary",
                    "bulgaria",
                    "czech republic",
                    "uruguay",
                }
                if k == "Region Segment" or weak or k in {
                    "Headquarters",
                    "Cities / Regions",
                    "Regions Served",
                    "Regional Extensions",
                    "Summary",
                    "Business Type",
                    "Website",
                }:
                    col = hdr.index(k) + 1
                    ws.cell(r_i, col).value = v
                    d[k] = v
            updated.append(d["Company"])
        companies.append(d)

    # Rebuild Region Summary + region sheets
    for name in list(wb.sheetnames):
        if name.startswith("Region - ") or name == "Region Summary":
            del wb[name]

    counts = Counter(d.get("Region Segment", "Other - Unclear") for d in companies)
    sum_ws = wb.create_sheet("Region Summary", 1)
    sum_ws["A1"] = "Global Smartwatch Market — Region-wise segmentation"
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
    notes = {
        "Global": "Multi-region / global coverage",
        "North America": "USA, Canada, Mexico",
        "Europe": "UK, EU, Balkans, Central Europe",
        "Asia Pacific": "India, China, Japan, SEA, ANZ, Korea",
        "Middle East": "GCC, UAE, Saudi, Turkey, Israel",
        "Latin America": "Brazil, Chile, Colombia, Uruguay, etc.",
        "Africa": "South Africa, Nigeria, Egypt, Kenya, etc.",
        "Eurasia / CIS": "Russia & CIS distributors/retailers",
        "Other - Unclear": "Still insufficient location signal",
    }
    for region in REGION_ORDER:
        n = counts.get(region, 0)
        if not n:
            continue
        sum_ws.cell(r, 1, region).border = THIN
        sum_ws.cell(r, 2, n).border = THIN
        sum_ws.cell(r, 3, f"{100 * n / total:.1f}%").border = THIN
        sum_ws.cell(r, 4, notes.get(region, "")).border = THIN
        r += 1
    sum_ws.cell(r, 1, "TOTAL").font = Font(bold=True)
    sum_ws.cell(r, 2, total).font = Font(bold=True)
    sum_ws.column_dimensions["A"].width = 22
    sum_ws.column_dimensions["B"].width = 16
    sum_ws.column_dimensions["C"].width = 12
    sum_ws.column_dimensions["D"].width = 48

    by_reg: dict[str, list[dict[str, str]]] = defaultdict(list)
    for d in companies:
        by_reg[d.get("Region Segment", "Other - Unclear")].append(d)

    for region in REGION_ORDER:
        rows = sorted(by_reg.get(region) or [], key=lambda x: _norm(x.get("Company", "")))
        if not rows:
            continue
        safe = region.replace("/", "-")
        title = f"Region - {safe}"[:31]
        rws = wb.create_sheet(title)
        _write_company_sheet(rws, hdr, rows)

    wb.save(path)
    unclear_left = [d["Company"] for d in companies if d.get("Region Segment") == "Other - Unclear"]
    return {
        "file": path.name,
        "updated": updated,
        "unclear_left": unclear_left,
        "counts": dict(counts),
        "total": total,
    }


def main() -> None:
    for p in PATHS:
        rep = fix_file(p)
        print("FILE", rep["file"])
        print(" updated", len(rep["updated"]), ":", ", ".join(rep["updated"]))
        print(" unclear left", rep["unclear_left"] or "none")
        for k in REGION_ORDER:
            if k in rep["counts"]:
                print(f"  {k}: {rep['counts'][k]}")
        print()


if __name__ == "__main__":
    main()
