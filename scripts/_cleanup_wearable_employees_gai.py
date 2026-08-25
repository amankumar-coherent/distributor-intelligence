# -*- coding: utf-8 -*-
"""Post-audit cleanup for Google AI Employees fills on wearable FINAL."""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import openpyxl

OUT = Path(r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global")
XLSX = OUT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = OUT / "wearable_employees_google_ai_audit.json"
NPD = "Not publicly disclosed"

# Corrected after reading Google AI evidence + agent web search
FIX: dict[str, str] = {
    "Al Noor Medical": "11-50",  # LinkedIn ~12; Overview used Hospitals Group ~4,500
    "Edan Instruments Asia Pacific": NPD,  # parent ~2,475; no APAC breakout
    "AED USA": "11-50",  # Overview/LinkedIn 11-50
    "EMS Supply Co.": NPD,  # ambiguous multi-entity namesakes
    "HSI Emergency Care": "11-50",  # local 11-50; parent HSI 1,000+
    "VitalConnect Asia Pacific": NPD,  # avoid parent VitalConnect leak
    "Farmacias Siman": "201-500",  # Clodura 201-500; 260+ stores
    "Claflin Company": "51-200",  # LeadIQ ~92
    "Buffalo Hospital Supply": "1-10",  # Overview ~3-10
    "MediMark Europe GmbH": "1-10",  # French registry ~1-2
}


def main() -> None:
    audit = json.loads(AUDIT.read_text(encoding="utf-8"))
    auto_npd: list[str] = []
    markers = (
        "parent",
        "hospitals group",
        "company-wide",
        "broader parent",
        "can refer to several",
        "not broken out",
        "asia pacific",
        "related entity",
        "several independent",
    )
    large = {
        "1,001-5,000",
        "501-1,000",
        "5,001-10,000",
        "10,001-50,000",
        "50,001+",
    }
    for a in audit:
        if a.get("source") != "google_ai_regex":
            continue
        co = a["company"]
        if co in FIX:
            continue
        prev = (a.get("evidence_preview") or "").lower()
        emp = a.get("employees")
        if emp in large and any(m in prev for m in markers):
            auto_npd.append(co)

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}

    changed: list[tuple[str, str, str, str]] = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, ic["Company"]).value or "").strip()
        old = str(ws.cell(r, ic["Employees"]).value or "").strip()
        if name in FIX:
            new = FIX[name]
            if old != new:
                ws.cell(r, ic["Employees"]).value = new
                changed.append((name, old, new, "fix"))
        elif name in auto_npd and old != NPD:
            ws.cell(r, ic["Employees"]).value = NPD
            changed.append((name, old, NPD, "auto_npd"))

    wb.save(XLSX)
    wb.close()

    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ei = headers.index("Employees")
    ctr: Counter[str] = Counter()
    npd = ranged = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        v = str(row[ei] or "").strip() or NPD
        ctr[v] += 1
        if v == NPD:
            npd += 1
        else:
            ranged += 1
    wb.close()

    print("changes:")
    for row in changed:
        print(" ", row)
    print("auto_npd:", auto_npd)
    print("ranged", ranged, "npd", npd)
    print(dict(ctr.most_common()))


if __name__ == "__main__":
    main()
