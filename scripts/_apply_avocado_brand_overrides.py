#!/usr/bin/env python3
"""One-shot apply of web-verified Key Brands overrides."""
from pathlib import Path

import openpyxl

X = Path(r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_OFFICE_SLIM.xlsx")

VERIFIED = {
    "Associated Food Stores": "Food Club, Full Circle, TopCare, Buckley Farms, Culinary Tours, BirchBerry, Crav'n Flavor, others",
    "Affiliated Foods": "IGA, others",
    "Bozzuto's": "IGA, Hy-Top, Seven Farms, others",
    "C&S Wholesale Grocers": "Best Yet, Our Family, Fresh & Finest by Our Family, IGA, others",
    "Alex Lee": "Brown Bag, IGA, others",
    "Chefs' Warehouse": "BelAria Essentials, Grand Reserve, Allen Brothers, Proper Stock, others",
}

wb = openpyxl.load_workbook(X)
ws = wb.active
hdr = [c.value for c in ws[2]]
ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}
n = 0
for r in range(3, ws.max_row + 1):
    co = str(ws.cell(r, ic["Company"]).value or "").strip()
    if co in VERIFIED:
        ws.cell(r, ic["Key Brands Represented"]).value = VERIFIED[co]
        n += 1
        print("set", co)
wb.save(X)
wb.close()
print("applied", n)
