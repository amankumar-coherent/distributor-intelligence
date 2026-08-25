#!/usr/bin/env python3
"""Fill verified public Office No. / Email for remaining NPD contact rows."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

X = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_OFFICE_SLIM.xlsx"
)
AUDIT = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\_avocado_phones_fill_audit.json"
)
NPD = "Not publicly disclosed"

PHONES = {
    "Core-Mark International, Inc. (Performance Food Group)": {
        "Office No.": "+1 800-622-1713",
        "Headquarters": "1500 Solana Blvd. Ste 3200, Westlake, Texas, United States",
    },
    "Aramark Middle East Distribution": {
        "Office No.": "+1 215-238-3000",
        "Headquarters": "2400 Market Street, Philadelphia, Pennsylvania, United States",
    },
    "Mitsubishi Shokuhin Co., Ltd.": {
        "Office No.": "+81 3-4553-5111",
        "Headquarters": "1-1 Koishikawa 1-chome, Bunkyo-ku, Tokyo, Japan",
        "Employees": "5001-10000",
    },
    "ofi (Olam Food Ingredients)": {
        "Office No.": "+65 6329 0800",
        "Headquarters": "7 Straits View, Marina One East Tower, Singapore",
    },
    "Patties Food Group": {
        "Office No.": "1800 650 069",
        "Office Email": "info@patties.com.au",
    },
    "Qualifirst Foods": {
        "Office No.": "+1 800-206-1177",
        "Office Email": "customerservice@qualifirst.com",
        "Headquarters": "89 Carlingview Drive, Etobicoke, Ontario, Canada",
    },
    "Prinova": {
        "Office No.": "+1 630-868-0300",
        "Office Email": "info@prinovausa.com",
        "Headquarters": "300 Park Blvd, Suite 500, Itasca, Illinois, United States",
    },
    "Supervalu (UNFI)": {"Office No.": "+1 401-528-8634"},
    "Supervalu Wholesale (UNFI)": {"Office No.": "+1 401-528-8634"},
    "Sid Wainer & Son": {
        "Office No.": "+1 800-423-8333",
    },
    "Vistar (Performance Food Group)": {
        "Office No.": "+1 800-880-9900",
    },
}

wb = openpyxl.load_workbook(X)
ws = wb.active
hdr = [c.value for c in ws[2]]
ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}
n = 0
for r in range(3, ws.max_row + 1):
    co = str(ws.cell(r, ic["Company"]).value or "").strip()
    if co not in PHONES:
        continue
    for k, v in PHONES[co].items():
        if k in ic and v != NPD:
            ws.cell(r, ic[k]).value = v
    n += 1
wb.save(X)
wb.close()

wb = openpyxl.load_workbook(X, data_only=True)
ws = wb.active
hdr = [c.value for c in ws[2]]
ic = {str(h): i for i, h in enumerate(hdr) if h}
real = npd = phone_npd = 0
still_brands = []
for row in ws.iter_rows(min_row=3, values_only=True):
    c = str(row[ic["Company"]] or "").strip()
    if not c:
        continue
    b = str(row[ic["Key Brands Represented"]] or "").strip()
    if b in ("", NPD):
        npd += 1
        still_brands.append(c)
    else:
        real += 1
    if str(row[ic["Office No."]] or "").strip() in ("", NPD):
        phone_npd += 1
wb.close()

audit = {
    "ts": datetime.now(timezone.utc).isoformat(),
    "phone_rows_updated": n,
    "real_brands": real,
    "npd_brands": npd,
    "still_npd_brands": still_brands,
    "phone_npd_remaining": phone_npd,
}
AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
print(json.dumps(audit, indent=2))
