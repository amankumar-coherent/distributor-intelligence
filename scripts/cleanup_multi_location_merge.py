# -*- coding: utf-8 -*-
"""Cleanup pass: merge leftover Bidfood/Nordic rows; fix Cities / Regions."""
from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path

import openpyxl

PREFERRED_CITIES = {
    "Alliance Healthcare": "Spain; France; Greece; Ireland; Italy; Netherlands; Norway; Poland; Portugal; Romania; UK; Germany (GEHE Pharma Handel)",
    "Nordic BioSite": "Sweden; Norway; Finland; Denmark",
    "2BScientific": "United Kingdom; Germany",
    "All Eights": "Malaysia; Singapore",
    "Biologix Pharma": "MENA",
    "Farmaenlace": "Colombia; Ecuador",
    "Genomax Technologies": "Malaysia; Singapore",
    "Zuellig Pharma": "Asia Pacific; Philippines",
    "Bidfood": "Asia; Australia; Belgium; Chile; Czech Republic; Netherlands; New Zealand; Portugal; Singapore; South Africa; Spain; Turkey; UK; Baltics; Poland (Farutex)",
    "Makro": "Brazil; Colombia; Netherlands",
    "Sysco": "USA; Canada; Europe",
    "Transgourmet": "France; Germany; Switzerland",
    "Azelis": "Belgium (HQ); Americas",
    "Barentz": "Netherlands (HQ); North America",
    "Brenntag": "Germany (HQ); North America",
    "Gordon Food Service": "USA; Canada",
    "Exertis": "United Kingdom; Ireland",
    "iStore": "Australia; South Africa",
}

CLEAN_PARENTS = {
    Path(
        r"D:\vendor-intel\output\web_expand\global_antibody_drug_conjugates_market_global\global_antibody_drug_conjugates_market_FINAL.xlsx"
    ): [
        "Alliance Healthcare",
        "Nordic BioSite",
        "2BScientific",
        "All Eights",
        "Biologix Pharma",
        "Farmaenlace",
        "Genomax Technologies",
        "Zuellig Pharma",
    ],
    Path(
        r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
    ): [
        "Bidfood",
        "Makro",
        "Sysco",
        "Transgourmet",
        "Azelis",
        "Barentz",
        "Brenntag",
        "Gordon Food Service",
    ],
    Path(
        r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market_FINAL.xlsx"
    ): ["Exertis"],
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.xlsx"
    ): ["iStore"],
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL_320plus.xlsx"
    ): ["iStore"],
    Path(
        r"D:\vendor-intel\output\web_expand\global_smartphone_market_global\global_smartphone_market_global_FINAL.xlsx"
    ): ["Exertis"],
}


def find_header(ws):
    for r in range(1, 6):
        vals = [c.value for c in ws[r]]
        labels = [str(v or "").strip() for v in vals]
        if "Company" in labels:
            return r, {h: i + 1 for i, h in enumerate(labels) if h}
    raise RuntimeError("no hdr")


def open_company_sheet(path: Path):
    wb = openpyxl.load_workbook(path)
    for name in wb.sheetnames:
        try:
            hdr, cols = find_header(wb[name])
            return wb, wb[name], hdr, cols
        except RuntimeError:
            continue
    raise RuntimeError(f"no company sheet in {path}")


def merge_by_prefix(path: Path, canonical: str, prefix: str, fields: dict) -> dict:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = path.with_name(f"{path.stem}_pre_extra_merge_{stamp}{path.suffix}")
    shutil.copy2(path, bak)
    wb, ws, hdr, cols = open_company_sheet(path)
    members = []
    for r in range(hdr + 1, ws.max_row + 1):
        co = str(ws.cell(r, cols["Company"]).value or "").strip()
        if co and co.lower().startswith(prefix.lower()):
            members.append((r, co))
    print(f"{path.name}: {prefix}* -> {[(r, c) for r, c in members]}")
    if len(members) <= 1:
        wb.save(path)
        return {"merged": 0, "members": [c for _, c in members], "backup": bak.name}

    parent = next(((r, c) for r, c in members if c == canonical), members[0])
    parent_r = parent[0]
    dels = [r for r, _ in members if r != parent_r]
    ws.cell(parent_r, cols["Company"]).value = canonical
    for f, v in fields.items():
        if f in cols:
            ws.cell(parent_r, cols[f]).value = v
    for r in sorted(dels, reverse=True):
        ws.delete_rows(r, 1)
    wb.save(path)
    return {
        "merged": len(members),
        "deleted": len(dels),
        "members": [c for _, c in members],
        "backup": bak.name,
    }


def clean_cities(path: Path, names: list[str]) -> int:
    wb, ws, hdr, cols = open_company_sheet(path)
    if "Cities / Regions" not in cols:
        wb.close()
        return 0
    fixed = 0
    for r in range(hdr + 1, ws.max_row + 1):
        co = str(ws.cell(r, cols["Company"]).value or "").strip()
        if co in names and co in PREFERRED_CITIES:
            ws.cell(r, cols["Cities / Regions"]).value = PREFERRED_CITIES[co]
            fixed += 1
    wb.save(path)
    return fixed


def main():
    print(
        merge_by_prefix(
            Path(
                r"D:\vendor-intel\output\web_expand\global_antibody_drug_conjugates_market_global\global_antibody_drug_conjugates_market_FINAL.xlsx"
            ),
            "Nordic BioSite",
            "Nordic BioSite",
            {
                "Headquarters": "Taby, Sweden",
                "Website": "https://www.nordicbiosite.com",
                "Cities / Regions": PREFERRED_CITIES["Nordic BioSite"],
            },
        )
    )
    print(
        merge_by_prefix(
            Path(
                r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
            ),
            "Bidfood",
            "Bidfood",
            {
                "Headquarters": "United Kingdom / international",
                "Website": "https://www.bidfood.co.uk",
                "Region Segment": "Global",
                "Cities / Regions": PREFERRED_CITIES["Bidfood"],
            },
        )
    )
    for path, names in CLEAN_PARENTS.items():
        if path.exists():
            n = clean_cities(path, names)
            print(f"cleaned cities {path.name}: {n}")


if __name__ == "__main__":
    main()
