# -*- coding: utf-8 -*-
"""Fix avocado oil Region Segment: remapping Global (and other non-six) to the 6 regions."""
from __future__ import annotations

import re
import shutil
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

SIX = {
    "Asia Pacific",
    "Europe",
    "North America",
    "Middle East",
    "Latin America",
    "Africa",
}

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
)

# Explicit overrides for known parents / ambiguous HQ text
EXPLICIT: dict[str, str] = {
    "azelis": "Europe",  # Antwerp HQ
    "barentz": "Europe",  # Netherlands HQ
    "bidfood": "Europe",  # UK HQ
    "brenntag": "Europe",  # Essen HQ
    "costco": "North America",  # Mexico listed under NA in 6-region model
    "diverse worldwide trading": "Europe",
    "fonterra foodservice distribution": "Asia Pacific",  # New Zealand
    "fresh produce consortium members distribution": "Europe",  # Bologna
    "fyffes": "Europe",  # Geneva
    "imcd": "Europe",  # Rotterdam
    "imcd us": "North America",
    "itochu food distribution": "Asia Pacific",  # Tokyo
    "makro": "Europe",  # Amsterdam HQ
    "metro ag": "Europe",  # Düsseldorf
    "mitsubishi corporation food distribution": "Asia Pacific",
    "olam food ingredients distribution": "Asia Pacific",  # Singapore / London → Singapore primary
    "sojitz": "Asia Pacific",
    "sysco": "North America",  # Houston
    "sysco specialty": "North America",
    "tradin organic": "Europe",  # Amsterdam
    "univar solutions": "North America",  # Illinois
}

CHECKS = [
    (
        "Asia Pacific",
        [
            "australia",
            "new zealand",
            "auckland",
            "india",
            "china",
            "japan",
            "tokyo",
            "singapore",
            "malaysia",
            "korea",
            "taiwan",
            "hong kong",
            "indonesia",
            "thailand",
            "vietnam",
            "philippines",
        ],
    ),
    (
        "Europe",
        [
            "united kingdom",
            "uk",
            "germany",
            "essen",
            "düsseldorf",
            "dusseldorf",
            "france",
            "italy",
            "bologna",
            "spain",
            "netherlands",
            "amsterdam",
            "rotterdam",
            "hoofddorp",
            "belgium",
            "antwerp",
            "switzerland",
            "geneva",
            "poland",
            "ireland",
            "sweden",
            "europe",
        ],
    ),
    (
        "North America",
        [
            "united states",
            "usa",
            "u.s.",
            "houston",
            "illinois",
            "downers grove",
            "canada",
            "mexico",
            "texas",
        ],
    ),
    (
        "Middle East",
        ["uae", "dubai", "saudi", "qatar", "kuwait", "bahrain", "oman", "israel", "turkey"],
    ),
    (
        "Latin America",
        ["brazil", "chile", "colombia", "argentina", "peru", "uruguay", "ecuador"],
    ),
    ("Africa", ["south africa", "nigeria", "egypt", "kenya", "morocco", "ghana"]),
]


def nk(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def resolve(company: str, hq: str, cities: str, served: str) -> str:
    key = nk(company)
    if key in EXPLICIT:
        return EXPLICIT[key]
    blob = f"{hq} {cities} {served}".lower()
    for region, keys in CHECKS:
        if any(k in blob for k in keys):
            return region
    return "Europe"


def find_header(ws):
    for r in range(1, 6):
        vals = [c.value for c in ws[r]]
        labels = [str(v or "").strip() for v in vals]
        if "Company" in labels and "Region Segment" in labels:
            return r, {h: i + 1 for i, h in enumerate(labels) if h}
    raise RuntimeError("header not found")


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_fix_global_region_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr, cols = find_header(ws)

    fixes = []
    for r in range(hdr + 1, ws.max_row + 1):
        co = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not co:
            continue
        old = str(ws.cell(r, cols["Region Segment"]).value or "").strip()
        if old in SIX:
            continue
        hq = str(ws.cell(r, cols.get("Headquarters", 1)).value or "")
        cities = str(ws.cell(r, cols.get("Cities / Regions", 1)).value or "")
        served = (
            str(ws.cell(r, cols["Regions Served"]).value or "")
            if "Regions Served" in cols
            else ""
        )
        new = resolve(co, hq, cities, served)
        ws.cell(r, cols["Region Segment"]).value = new
        fixes.append((r, co, old, new, hq[:50]))

    # counts
    left = Counter()
    for r in range(hdr + 1, ws.max_row + 1):
        co = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not co:
            continue
        left[str(ws.cell(r, cols["Region Segment"]).value or "").strip()] += 1

    an = "Region Six Fix Audit"
    if an in wb.sheetnames:
        del wb[an]
    aws = wb.create_sheet(an)
    aws.append(["Row", "Company", "Old Region", "New Region", "HQ"])
    for row in fixes:
        aws.append(list(row))

    wb.save(PATH)
    print(f"backup: {bak.name}")
    print(f"fixed {len(fixes)} rows:")
    for r, co, old, new, hq in fixes:
        print(f"  {r}: {co}: {old} -> {new}  (HQ={hq})")
    print("final counts:")
    for k in [
        "Asia Pacific",
        "Europe",
        "North America",
        "Middle East",
        "Latin America",
        "Africa",
    ]:
        print(f"  {k}: {left.get(k, 0)}")
    bad = [k for k in left if k not in SIX]
    print(f"non-six left: {bad or 'none'}")


if __name__ == "__main__":
    main()
