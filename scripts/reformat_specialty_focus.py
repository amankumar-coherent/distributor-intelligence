#!/usr/bin/env python3
"""Reformat Specialty Focus into a compact structured phrase:

    [Product Category] + [Specific Application / Customer Segment] + [Value-added expertise (optional)]

Pure local regex/keyword composition over data already in the sheet
(Core Categories + current Specialty Focus + Summary) -- no new web calls,
no LLM. Rows already flagged "Needs manual verification" are left untouched.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

import openpyxl

XLSX = ROOT / "output/web_expand/global_wearable_medical_devices_market_global/global_wearable_medical_devices_market_global_FINAL.xlsx"
SHEET = "Wearable Medical Companies"
FLAG_PREFIX = "Needs manual verification"

# (keyword pattern, segment phrase) — first match wins, most specific first.
_SEGMENT_RULES: list[tuple[str, str]] = [
    (r"neonatal|maternal|paediatric|pediatric", "neonatal, maternal, and pediatric care providers"),
    (r"veterinar", "veterinary clinics"),
    (r"\bdental\b", "dental practices"),
    (r"home\s?care|home\s?health|homecare", "home healthcare patients and caregivers"),
    (r"retail pharmac|pharmacy chain|drugstore|drug store", "retail pharmacy consumers"),
    (r"community pharmac", "community pharmacies"),
    (r"laborator(?:y|ies)|diagnostic (?:center|centre)|clinical lab", "clinical and diagnostic laboratories"),
    (r"government|public sector|ministry of health|\bngo\b", "government health institutions and NGOs"),
    (r"direct.to.consumer|\bd2c\b|consumers directly|online pharmacy|telehealth", "consumers directly"),
    (r"hospitals? and clinics|hospitals?,? clinics?|healthcare providers", "hospitals and healthcare providers"),
    (r"animal (?:health|production)|livestock", "veterinary and animal-health providers"),
]
_DEFAULT_SEGMENT = "hospitals and healthcare providers"

# (keyword pattern, value-add phrase) — optional, only included if matched.
_VALUE_ADD_RULES: list[tuple[str, str]] = [
    (r"cold.chain", "cold-chain logistics"),
    (r"installation,?\s*(?:repair,?\s*)?(?:and\s*)?maintenance(?:,?\s*and\s*training)?|maintenance and training", "installation, maintenance, and training support"),
    (r"regulatory (?:compliance|approval|registration)", "regulatory compliance support"),
    (r"after.sales", "after-sales support"),
    (r"turnkey", "turnkey solutions"),
    (r"nationwide (?:depot|distribution|logistics)|depot fulfil?ment", "a nationwide logistics network"),
    (r"warehousing|cold storage", "warehousing and storage solutions"),
    (r"technical support|biomedical (?:engineering|service)", "technical and biomedical engineering support"),
    (r"staffing|nursing staff", "healthcare staffing services"),
    (r"training and (?:technical )?support|training services", "training and technical support"),
]


def compose(core_categories: str, source_text: str) -> str:
    low = source_text.lower()
    segment = _DEFAULT_SEGMENT
    for pattern, phrase in _SEGMENT_RULES:
        if re.search(pattern, low):
            segment = phrase
            break

    value_add = ""
    for pattern, phrase in _VALUE_ADD_RULES:
        if re.search(pattern, low):
            value_add = phrase
            break

    out = f"{core_categories} for {segment}"
    if value_add:
        out += f", with {value_add}"
    return out + "."


def main() -> None:
    wb = openpyxl.load_workbook(XLSX)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(header)}

    updated = 0
    skipped_flagged = 0
    for r in range(2, ws.max_row + 1):
        company = ws.cell(row=r, column=idx["Company"]).value
        if not company:
            continue
        core = str(ws.cell(row=r, column=idx["Core Categories"]).value or "").strip()
        focus = str(ws.cell(row=r, column=idx["Specialty Focus"]).value or "").strip()
        summary = str(ws.cell(row=r, column=idx["Summary"]).value or "").strip()

        if core.startswith(FLAG_PREFIX) or focus.startswith(FLAG_PREFIX):
            skipped_flagged += 1
            continue
        if not core:
            continue

        source_text = f"{focus} {summary}"
        new_focus = compose(core, source_text)
        ws.cell(row=r, column=idx["Specialty Focus"], value=new_focus)
        updated += 1

    wb.save(XLSX)
    print(f"Updated {updated} rows, skipped {skipped_flagged} flagged rows.")


if __name__ == "__main__":
    main()
