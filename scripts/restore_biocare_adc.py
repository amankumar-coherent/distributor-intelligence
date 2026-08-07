#!/usr/bin/env python3
"""Restore BioCareSD false-negative and finalize ADC audit summary."""
from __future__ import annotations

import json
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = Path("output/web_expand/global_antibody_drug_conjugates_market_global")
XLSX = OUT / "global_antibody_drug_conjugates_market_FINAL.xlsx"

HEADERS = [
    "Company", "Founded", "Headquarters", "Cities / Regions", "Ownership",
    "Business Type", "Employees", "Turnover / Scale", "Contact Person", "Role",
    "Email", "Phone / WhatsApp", "LinkedIn", "Website", "Core Categories",
    "Specialty Focus", "Price Segment", "Key Brands Represented", "Exclusivity",
    "Partnership Duration", "Retail Chains", "Specialty Grocers", "Foodservice / HoReCa",
    "E-commerce", "Channel Strength", "Distribution Type", "Regions Served",
    "Regional Extensions", "Warehouse / Logistics", "Delivery / Storage",
    "Competitive Benchmark", "Data Confidence", "Quality Score", "Data Sources", "Summary",
]
WIDTHS = {
    "Company": 28, "Headquarters": 22, "Key Brands Represented": 32,
    "Specialty Focus": 28, "Core Categories": 28, "Summary": 36,
    "Website": 24, "Regions Served": 22,
}


def main() -> None:
    wb = load_workbook(XLSX, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows = []
    for r in range(3, ws.max_row + 1):
        d = {h: str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)

    if not any("biocare" in r["Company"].lower() for r in rows):
        rows.append({
            "Company": "BioCare Specialty Distribution (BioCareSD)",
            "Founded": "1982",
            "Headquarters": "Tempe, Arizona, United States",
            "Cities / Regions": "Tempe, Arizona, United States",
            "Ownership": "Private",
            "Business Type": "Specialty Distributor",
            "Employees": "Not publicly disclosed",
            "Turnover / Scale": "Not publicly disclosed",
            "Contact Person": "Not publicly disclosed",
            "Role": "Not publicly disclosed",
            "Email": "Not publicly disclosed",
            "Phone / WhatsApp": "Not publicly disclosed",
            "LinkedIn": "Not publicly disclosed",
            "Website": "https://biocare-us.com",
            "Core Categories": "Antibody-drug conjugates; Oncology biologics; Specialty pharmaceutical distribution",
            "Specialty Focus": "National specialty distributor for oncology, rare/orphan, and complex biologics (BioCareSD)",
            "Price Segment": "Specialty / Hospital",
            "Key Brands Represented": "Authorized specialty / oncology manufacturer partners (portfolio varies)",
            "Exclusivity": "Not publicly disclosed",
            "Partnership Duration": "Not publicly disclosed",
            "Retail Chains": "Not applicable",
            "Specialty Grocers": "Not applicable",
            "Foodservice / HoReCa": "Not applicable",
            "E-commerce": "B2B / partner portal",
            "Channel Strength": "8",
            "Distribution Type": "Specialty Distributor",
            "Regions Served": "United States",
            "Regional Extensions": "United States",
            "Warehouse / Logistics": "Cold-chain specialty DCs (Tempe AZ; Olive Branch MS) + forward stocking",
            "Delivery / Storage": "Temperature-controlled specialty fulfillment; STAT/emergency delivery",
            "Competitive Benchmark": "8",
            "Data Confidence": "90",
            "Quality Score": "90",
            "Data Sources": "web_verify biocare-us.com; allygpo pr; adc_restore_false_negative",
            "Summary": (
                "BioCareSD is a US specialty pharmaceutical distributor (since 1982) serving "
                "oncology and rare-disease channels — distinct from BioCare Medical IHC reagent OEM."
            ),
        })
        print("Restored BioCareSD")
    else:
        print("BioCare already present")

    for r in rows:
        if r["Company"] == "AAH Pharmaceuticals":
            r["Founded"] = "1912"
            r["Headquarters"] = "Warwick, United Kingdom"
            r["Cities / Regions"] = "Warwick, United Kingdom"
            if "web_verify_aah" not in r.get("Data Sources", ""):
                r["Data Sources"] = (
                    r.get("Data Sources", "") + "; web_verify_aah wikipedia companies_house"
                ).strip("; ")

    rows.sort(key=lambda x: x["Company"].lower())

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "ADC Distributors"
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    t = ws2.cell(
        1, 1,
        "Global Antibody Drug Conjugates Market — Distributors / Channel Partners (audited)",
    )
    t.font = Font(bold=True, size=13, color="FFFFFF")
    t.fill = PatternFill("solid", fgColor="1F4E79")
    hf = PatternFill("solid", fgColor="2E75B6")
    hfont = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="BDD7EE"),
        right=Side(style="thin", color="BDD7EE"),
        top=Side(style="thin", color="BDD7EE"),
        bottom=Side(style="thin", color="BDD7EE"),
    )
    for i, h in enumerate(HEADERS, 1):
        c = ws2.cell(2, i, h)
        c.fill = hf
        c.font = hfont
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = thin
        ws2.column_dimensions[get_column_letter(i)].width = WIDTHS.get(h, 16)
    for ri, row in enumerate(rows, 3):
        for ci, h in enumerate(HEADERS, 1):
            cell = ws2.cell(ri, ci, row.get(h, ""))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")
    cws = wb2.create_sheet("Audit Checklist")
    cws["A1"] = "Audited with web search + DeepSeek (deepseek-chat / v4-flash via OpenAI SDK)"
    cws["A2"] = "Kept distributors / specialty pharmacies / channel partners for ADC/oncology/specialty biologics"
    cws["A3"] = "Removed manufacturers, CDMOs, insurers, hospitals, pure 3PL logistics, unrelated entities"
    cws["A4"] = f"Rows={len(rows)} | {datetime.now(timezone.utc).isoformat()}"
    wb2.save(XLSX)

    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "final_count": len(rows),
        "original_count": 372,
        "removed_net": 372 - len(rows),
        "business_types": dict(Counter(r.get("Business Type") for r in rows)),
        "model": "deepseek-chat (deepseek-v4-flash)",
        "output": str(XLSX),
    }
    (OUT / "adc_audit_final_summary.json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8"
    )
    print("FINAL", len(rows), summary["business_types"])


if __name__ == "__main__":
    main()
