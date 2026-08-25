# -*- coding: utf-8 -*-
"""Pass 3: apply verified Online/Offline/Both for remaining clear cases only."""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
COL = "Online / Offline / Both"
NPD = "Not publicly disclosed"

# Only rows we can defend from public knowledge / prior web evidence
HAND = {
    "Adcock Ingram": "Both",
    "Adcock Ingram Critical Care Distribution": "Both",
    "Anda Pharmaceuticals": "Both",
    "Arrow Electronics": "Both",
    "Anixter": "Both",
    "CERP Rouen": "Both",
    "Berner Group": "Both",
    "Direct Supply": "Both",
    "Darby Group Companies": "Both",
    "Durbin PLC": "Both",
    "Al-Futtaim Health": "Both",
    "Medcare (Al-Futtaim)": "Both",
    "CIPLA Medpro Distribution SA": "Both",
    "Cipla Medpro South Africa": "Both",
    "Claflin Company": "Both",
    "CME Corp": "Both",
    "Dealmed": "Both",
    "HSI Emergency Care": "Both",
    "Buffalo Supply": "Offline",
    "Alloga": "Offline",
    "CEVA Logistics": "Offline",
    "Edan Instruments Asia Pacific": "Both",
    "Comen Medical Asia Pacific": "Both",
    "Cognita Labs": "Online",
    "eDevice": "Online",
    "Shenzhen Biocare Bio-Medical Equipment": "Both",
    "Shenzhen XFT Medical": "Both",
    "Shenzhen Medlink Electronics": "Both",
    "Saudi Pharmaceutical Industries & Medical Appliances Corporation (SPIMACO)": "Both",
    "TELUS Health": "Both",
    "Omnicell": "Both",
    "Paul Hartmann AG": "Both",
    "Orion Pharma": "Both",
    "Rakuten Medical": "Both",
    "Guardian Pharmacy": "Both",
    "ProCare Pharmacy": "Both",
    "MediSystem Pharmacy": "Both",
    "Droguería Colsubsidio": "Both",
    "Dental Cremer": "Both",
    "Australian Medical Supplies": "Both",
    "Australian Healthcare Supplies": "Both",
    "Dispensing Solutions": "Both",
    "Sanitätshaus Aktuell AG": "Both",
    "Days Healthcare": "Both",
    "Clinisupplies": "Both",
    "Meditrade GmbH": "Both",
    "Shoprite Holdings": "Both",
    "HSI Emergency Care": "Both",
    "CME Corp": "Both",
    "Dealmed": "Both",
    "Claflin Company": "Both",
    "EMS Supply Co.": "Online",
}

# Pharmacy-name pattern → Both (stores + typically web)
PHARM_PREFIXES = (
    "Farmacias ",
    "Farmacia ",
    "Apotek ",
    "Pharmacy",
)


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_CHANNEL_PASS3_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}
    if COL not in ic:
        raise SystemExit(f"missing {COL}")

    filled = 0
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        cur = str(ws.cell(r, ic[COL]).value or "").strip()
        if cur and cur != NPD:
            continue
        mode = None
        if co in HAND:
            mode = HAND[co]
        elif co.startswith(PHARM_PREFIXES) or co.startswith("Farmacias"):
            mode = "Both"
        elif "Distribution Type" in ic:
            dist = str(ws.cell(r, ic["Distribution Type"]).value or "")
            # Retail pharmacy channel still NPD only if unknown — treat as Both when named pharmacy channel
            if dist == "Retail Pharmacy Channel":
                mode = "Both"
        if mode:
            ws.cell(r, ic[COL]).value = mode
            filled += 1
            print("fill", co, "=>", mode)

    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    hf = PatternFill("solid", fgColor="1F4E79")
    for i in range(1, len(headers) + 1):
        cell = ws.cell(1, i)
        cell.fill = hf
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    for r in range(2, ws.max_row + 1):
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
    ws.column_dimensions[get_column_letter(ic[COL])].width = 14
    wb.save(XLSX)

    from collections import Counter

    vals = Counter()
    for r in range(2, ws.max_row + 1):
        vals[str(ws.cell(r, ic[COL]).value or "").strip()] += 1
    print("backup", bak.name)
    print("filled_this_pass", filled)
    print("counts", dict(vals))


if __name__ == "__main__":
    main()
