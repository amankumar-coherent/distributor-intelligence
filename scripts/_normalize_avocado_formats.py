#!/usr/bin/env python3
"""Normalize continent labels + Founded years to numbers; trim trailing empty cols."""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

X = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_OFFICE_SLIM.xlsx"
)

CONT_MAP = {
    "APAC(Malaysia)": "Asia-Pacific",
    "APAC": "Asia-Pacific",
    "Asia Pacific": "Asia-Pacific",
    "LATAM(Mexico)": "Latin America",
    "LATAM(Chile)": "Latin America",
    "LATAM(Colombia)": "Latin America",
    "LATAM(Brazil)": "Latin America",
    "LATAM(Argentina)": "Latin America",
    "LATAM": "Latin America",
    "LatAm": "Latin America",
    "ME": "Middle East",
    "NA": "North America",
    "EU": "Europe",
    "AF": "Africa",
}

wb = openpyxl.load_workbook(X)
ws = wb.active
hdr = [c.value for c in ws[2]]
ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}

# drop trailing fully-empty columns beyond last named header
last_named = max(i for i, h in enumerate(hdr, 1) if h)
while ws.max_column > last_named:
    ws.delete_cols(ws.max_column)

cont_fixed = founded_fixed = 0
for r in range(3, ws.max_row + 1):
    co = str(ws.cell(r, ic["Company"]).value or "").strip()
    if not co:
        continue

    cont = str(ws.cell(r, ic["Continent / Geography"]).value or "").strip()
    if cont in CONT_MAP:
        ws.cell(r, ic["Continent / Geography"]).value = CONT_MAP[cont]
        cont_fixed += 1
    elif cont.startswith("LATAM(") or cont.startswith("APAC(") or cont.startswith("NA(") or cont.startswith("EU(") or cont.startswith("ME("):
        # generic PREFIX(Country) -> mapped region
        prefix = cont.split("(", 1)[0]
        mapped = CONT_MAP.get(prefix) or CONT_MAP.get(prefix.upper())
        if mapped:
            ws.cell(r, ic["Continent / Geography"]).value = mapped
            cont_fixed += 1

    fv = ws.cell(r, ic["Founded"]).value
    if isinstance(fv, str):
        m = re.search(r"(19|20)\d{2}", fv)
        if m:
            ws.cell(r, ic["Founded"]).value = int(m.group(0))
            founded_fixed += 1
    elif isinstance(fv, float) and fv == int(fv):
        ws.cell(r, ic["Founded"]).value = int(fv)
        founded_fixed += 1

wb.save(X)
wb.close()

# verify
wb = openpyxl.load_workbook(X)
ws = wb.active
hdr = [c.value for c in ws[2]]
ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}
from collections import Counter

cont = Counter()
ft = Counter()
for r in range(3, ws.max_row + 1):
    if not str(ws.cell(r, ic["Company"]).value or "").strip():
        continue
    cont[str(ws.cell(r, ic["Continent / Geography"]).value or "").strip()] += 1
    ft[type(ws.cell(r, ic["Founded"]).value).__name__] += 1
print("cont_fixed", cont_fixed, "founded_fixed", founded_fixed)
print("continents", dict(cont))
print("founded_types", dict(ft))
print("max_col", ws.max_column, "headers", [h for h in hdr if h])
wb.close()
