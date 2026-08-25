# -*- coding: utf-8 -*-
"""Pass 3: verified brand overrides + honest soft fills (no invented brand names)."""
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
NPD = "Not publicly disclosed"

VERIFIED = {
    "Adcock Ingram": "Adcock Ingram, Adco-Hygiene, others",
    "Adcock Ingram Critical Care Distribution": "Adcock Ingram Critical Care, Adco-Hygiene, others",
    "Arrow Electronics": (
        "TE Connectivity, STMicroelectronics, Silicon Labs, onsemi, Molex, "
        "Infineon, Analog Devices, Amphenol, others"
    ),
    "Medi-Dyne Healthcare Products": "2Toms, Tuli's, Cho-Pat, ProStretch, others",
    "Ecolab": "Ecolab, Belimed, Censis, Nalco Water, others",
    "Shenzhen Biocare Bio-Medical Equipment": "Biocare, others",
    "Saudi Pharmaceutical Industries & Medical Appliances Corporation (SPIMACO)": "SPIMACO, others",
    "Medipal Holdings Corporation": "Multi-vendor pharmaceutical principals (MEDICEO/PALTAC group), others",
    "Alfresa Holdings Corporation": "Multi-vendor pharmaceutical principals, others",
    "Cipla Medpro South Africa": "Cipla, others",
    "CIPLA Medpro Distribution SA": "Cipla, others",
    "Boryung": "Boryung, others",
    "Al-Dawaa Medical Services": "Multi-brand pharmacy assortment, others",
    "Dis-Chem Pharmacies": "Multi-brand pharmacy assortment, others",
    "Apotek Hjärtat": "Multi-brand pharmacy assortment, others",
    "Jean Coutu": "Multi-brand pharmacy assortment, others",
    "Rite Aid": "Multi-brand pharmacy assortment, others",
    "London Drugs": "Multi-brand pharmacy and retail assortment, others",
    "Giant Eagle": "Multi-brand pharmacy and grocery retail assortment, others",
}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_BRANDS_PASS3_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    filled_v = soft_n = 0
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        cur = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
        if co in VERIFIED:
            ws.cell(r, ic["Key Brands Represented"]).value = VERIFIED[co]
            filled_v += 1
            continue
        if cur and cur != NPD:
            continue
        dist = (
            str(ws.cell(r, ic["Distribution Type"]).value or "").lower()
            if "Distribution Type" in ic
            else ""
        )
        core = str(ws.cell(r, ic["Core Categories"]).value or "").lower()
        if "retail pharmacy" in dist or "drugstore" in core or (
            "retail pharmacy" in core and "otc" in core
        ):
            ws.cell(r, ic["Key Brands Represented"]).value = (
                "Multi-brand pharmacy assortment, others"
            )
            soft_n += 1
        elif "broadline distributor" in dist and (
            "pharmaceutical" in core or "otc healthcare" in core or core.startswith("pharmaceuticals")
        ):
            ws.cell(r, ic["Key Brands Represented"]).value = (
                "Multi-vendor pharmaceutical principals, others"
            )
            soft_n += 1
        # specialty distributors/resellers stay NPD unless verified — do not invent

    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    hf = PatternFill("solid", fgColor="1F4E79")
    widths = {
        "Company": 34,
        "Website": 36,
        "Key Brands Represented": 42,
        "Specialty Focus": 55,
        "Core Categories": 40,
        "Continent / Geography": 52,
        "Operational Presence": 48,
        "Ownership": 42,
        "Summary": 48,
    }
    for i, hv in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(str(hv or ""), 16)
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

    wb.save(XLSX)

    filled = npd = named = soft = 0
    for r in range(2, ws.max_row + 1):
        v = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
        if v == NPD:
            npd += 1
        else:
            filled += 1
            if v.lower().startswith("multi-"):
                soft += 1
            else:
                named += 1
    print("backup", bak.name)
    print("verified_applied", filled_v, "soft_applied", soft_n)
    print(
        "TOTAL filled",
        filled,
        "named",
        named,
        "soft",
        soft,
        "npd",
        npd,
        "rows",
        ws.max_row - 1,
    )


if __name__ == "__main__":
    main()
