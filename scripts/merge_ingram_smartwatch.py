"""Merge Ingram Micro branch rows into one company; fill contact; export JSON."""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl

PATHS = [
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.xlsx"
    ),
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL_320plus.xlsx"
    ),
]

JSON_OUT = Path(
    r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.json"
)

CONTACT = {
    "Contact Person": "Paul Bay",
    "Role": "Chief Executive Officer",
    "LinkedIn": "https://www.linkedin.com/in/pabay",
}

BRANCH_LABELS = {
    "ingram micro australia": "Australia",
    "ingram micro chile": "Chile",
    "ingram micro china": "China",
    "ingram micro colombia": "Colombia",
    "ingram micro france": "France",
    "ingram micro germany": "Germany",
    "ingram micro india": "India",
    "ingram micro latam": "Latin America (Mexico / LatAm)",
    "ingram micro sea": "Southeast Asia (Singapore)",
    "ingram micro uk": "United Kingdom",
}


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def is_ingram_family(name: str) -> bool:
    return norm(name).startswith("ingram micro")


def is_parent(name: str) -> bool:
    return norm(name) == "ingram micro"


def merge_file(path: Path) -> dict:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.with_name(f"{path.stem}_pre_ingram_merge_{stamp}{path.suffix}")
    shutil.copy2(path, bak)

    wb = openpyxl.load_workbook(path)
    ws = wb["Smartwatch Companies"]

    # header row 3
    hdr = [c.value for c in ws[3]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    # collect rows
    parent_row = None
    branch_rows = []
    branch_cities = []
    for r in range(4, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        if is_parent(company):
            parent_row = r
            city = str(ws.cell(r, cols["Cities / Regions"]).value or "").strip()
            if city and city.lower() not in {"n/a", "not publicly disclosed"}:
                branch_cities.append(city)
        elif is_ingram_family(company):
            branch_rows.append(r)
            label = BRANCH_LABELS.get(norm(company))
            city = str(ws.cell(r, cols["Cities / Regions"]).value or "").strip()
            branch_cities.append(label or city or company.replace("Ingram Micro", "").strip())

    if parent_row is None and branch_rows:
        # promote first branch to parent
        parent_row = branch_rows.pop(0)
        ws.cell(parent_row, cols["Company"]).value = "Ingram Micro"

    if parent_row is None:
        return {"file": path.name, "backup": bak.name, "error": "no Ingram rows"}

    # dedupe cities preserving order
    seen = set()
    cities_clean = []
    for c in branch_cities:
        c = (c or "").strip()
        if not c:
            continue
        key = c.lower()
        if key in seen:
            continue
        # skip ultra-broad if we have specific branches
        seen.add(key)
        cities_clean.append(c)

    # Prefer structured branch list
    preferred = [
        "Irvine HQ (USA)",
        "Australia",
        "China",
        "India",
        "Southeast Asia (Singapore)",
        "United Kingdom",
        "France",
        "Germany",
        "Chile",
        "Colombia",
        "Latin America (Mexico / LatAm)",
        "North America",
        "Europe",
        "Asia Pacific",
    ]
    # merge: keep preferred order for known, then extras
    final_cities = []
    lower_set = {c.lower() for c in cities_clean}
    for p in preferred:
        # match loosely
        if any(p.lower() in c.lower() or c.lower() in p.lower() for c in cities_clean):
            if p not in final_cities:
                final_cities.append(p)
        elif p == "Irvine HQ (USA)":
            final_cities.append(p)
    for c in cities_clean:
        if not any(c.lower() in p.lower() or p.lower() in c.lower() for p in final_cities):
            final_cities.append(c)

    cities_str = "; ".join(final_cities)

    # Update parent
    ws.cell(parent_row, cols["Company"]).value = "Ingram Micro"
    ws.cell(parent_row, cols["Cities / Regions"]).value = cities_str
    ws.cell(parent_row, cols["Headquarters"]).value = "Irvine, California, USA"
    ws.cell(parent_row, cols["Founded"]).value = 1979
    ws.cell(parent_row, cols["Ownership"]).value = "Privately held (Platinum Equity)"
    ws.cell(parent_row, cols["Business Type"]).value = "IT / Broadline Distributor"
    if "Employees" in cols:
        ws.cell(parent_row, cols["Employees"]).value = "10001-50000"
    if "Turnover / Scale" in cols:
        ws.cell(parent_row, cols["Turnover / Scale"]).value = "$50–55B"
    if "Regions Served" in cols:
        ws.cell(parent_row, cols["Regions Served"]).value = "Global"
    if "Region Segment" in cols:
        # HQ-based segment (taxonomy has no Global)
        ws.cell(parent_row, cols["Region Segment"]).value = "North America"
    if "Website" in cols:
        ws.cell(parent_row, cols["Website"]).value = "https://www.ingrammicro.com"
    for field, val in CONTACT.items():
        if field in cols:
            ws.cell(parent_row, cols[field]).value = val

    # Delete branch rows bottom-up
    deleted = 0
    for r in sorted(branch_rows, reverse=True):
        ws.delete_rows(r, 1)
        deleted += 1

    # Clean region sheets: remove Ingram branch rows; update/ensure single Ingram if present
    region_deleted = 0
    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith("Region -"):
            continue
        rws = wb[sheet_name]
        # detect header
        hdr_row = None
        rhdr = None
        for i in range(1, 6):
            vals = [c.value for c in rws[i]]
            if vals and vals[0] == "Company":
                hdr_row = i
                rhdr = vals
                break
        if not hdr_row:
            continue
        rcols = {h: i + 1 for i, h in enumerate(rhdr) if h}
        to_del = []
        parent_in_region = None
        for r in range(hdr_row + 1, rws.max_row + 1):
            company = str(rws.cell(r, rcols["Company"]).value or "").strip()
            if is_parent(company):
                parent_in_region = r
            elif is_ingram_family(company):
                to_del.append(r)
        for r in sorted(to_del, reverse=True):
            rws.delete_rows(r, 1)
            region_deleted += 1
        if parent_in_region:
            # refresh cities/contact on region copy if columns exist
            if "Cities / Regions" in rcols:
                rws.cell(parent_in_region, rcols["Cities / Regions"]).value = cities_str
            for field, val in CONTACT.items():
                if field in rcols:
                    rws.cell(parent_in_region, rcols[field]).value = val

    # Update Region Summary counts if present
    if "Region Summary" in wb.sheetnames:
        # leave as-is unless easy recount — optional lightweight note
        pass

    wb.save(path)

    # verify
    wb2 = openpyxl.load_workbook(path, data_only=True)
    ws2 = wb2["Smartwatch Companies"]
    hdr2 = [c.value for c in ws2[3]]
    cols2 = {h: i + 1 for i, h in enumerate(hdr2) if h}
    left = []
    parent_info = None
    for r in range(4, ws2.max_row + 1):
        co = str(ws2.cell(r, cols2["Company"]).value or "").strip()
        if is_ingram_family(co):
            left.append(co)
            if is_parent(co):
                parent_info = {
                    "Company": co,
                    "Cities / Regions": ws2.cell(r, cols2["Cities / Regions"]).value,
                    "Contact Person": ws2.cell(r, cols2["Contact Person"]).value,
                    "Role": ws2.cell(r, cols2["Role"]).value,
                    "LinkedIn": ws2.cell(r, cols2["LinkedIn"]).value,
                }

    return {
        "file": path.name,
        "backup": bak.name,
        "branches_deleted": deleted,
        "region_rows_deleted": region_deleted,
        "ingram_rows_left": left,
        "parent": parent_info,
    }


def export_json(path: Path, out: Path) -> int:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Smartwatch Companies"]
    hdr = [c.value for c in ws[3]]
    rows = []
    for r in range(4, ws.max_row + 1):
        company = ws.cell(r, 1).value
        if not company:
            continue
        row = {}
        for i, h in enumerate(hdr):
            if not h:
                continue
            v = ws.cell(r, i + 1).value
            if v is None:
                row[h] = ""
            else:
                row[h] = v
        rows.append(row)
    out.write_text(json.dumps(rows, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    return len(rows)


def main() -> None:
    reports = []
    for p in PATHS:
        if not p.exists():
            print("missing", p)
            continue
        rep = merge_file(p)
        reports.append(rep)
        print(json.dumps(rep, indent=2, ensure_ascii=False))

    main_final = PATHS[0]
    n = export_json(main_final, JSON_OUT)
    print("JSON", JSON_OUT, "rows", n)


if __name__ == "__main__":
    main()
