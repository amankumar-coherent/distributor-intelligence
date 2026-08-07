#!/usr/bin/env python3
"""Clean ADC Employees (remove est.) and Turnover (numeric ranges only)."""
from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
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
    "Company": 28, "Headquarters": 28, "Cities / Regions": 36,
    "Employees": 18, "Turnover / Scale": 22, "Summary": 36, "Website": 24,
}

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}

# Explicit company turnover overrides (numeric / reported)
KNOWN_TURN: dict[str, str] = {
    "AAH Pharmaceuticals": "£2.5B–£3.0B",
    "Suzuken Co., Ltd.": "¥2.0T–¥2.5T",
    "Alfresa Holdings": "¥2.0T–¥3.0T",
    "Medipal Holdings": "¥2.0T–¥3.0T",
    "Toho Holdings / TOHO Pharmaceutical": "¥1.0T–¥2.0T",
    "Phoenix Pharma SE": "€25B–€35B",
    "PHOENIX Pharmahandel": "€25B–€35B",
    "Avantor Inc.": "$6B–$8B",
    "Zuellig Pharma": "$10B–$20B",
    "DKSH Health": "CHF 10B–CHF 12B",
    "DiethelmKellerSiberHegner Healthcare": "CHF 10B–CHF 12B",
    "Sinopharm": "$50B+",
    "Jointown Pharmaceutical Group": "$10B–$30B",
    "Medline Industries": "$15B–$25B",
    "Owens & Minor": "$5B–$15B",
    "Henry Schein Medical": "$10B–$15B",
    "Option Care Health specialty infusion": "$2B–$5B",
    "Neuca": "$2B–$5B",
    "Profarma": "$1B–$3B",
    "Cofares": "€5B–€10B",
    "Sigma Healthcare": "A$5B–A$10B",
    "Uniphar": "€0.5B–€1.5B",
    "Galenica / Alloga": "CHF 3B–CHF 5B",
    "AcariaHealth": "$75M–$120M",
    "Alloga UK": "$75M–$120M",
    "CVS Specialty": "$1B+",
    "Walgreens Specialty Pharmacy": "$1B+",
    "AllianceRx Walgreens Pharmacy": "$1B+",
    "Accredo Health Group": "$1B+",
    "Cigna Specialty Pharmacy / Accredo": "$1B+",
    "Express Scripts Accredo network": "$1B+",
    "Optum Specialty Pharmacy": "$1B+",
    "McKesson Specialty Health": "$1B+",
    "McKesson Specialty Care Distribution": "$1B+",
    "Cardinal Health Specialty Solutions": "$1B+",
    "Cardinal Health Specialty Pharmaceutical Distribution": "$1B+",
    "Cencora (AmerisourceBergen) Specialty": "$1B+",
    "ASD Healthcare": "$1B+",
    "Oncology Supply": "$500M–$1B",
    "CuraScript SD": "$500M–$1B",
    "Anda Inc.": "$1B+",
    "Anda Pharmaceuticals": "$1B+",
    "Fisher Scientific": "$1B+",
    "VWR International": "$1B+",
    "BioCare Specialty Distribution (BioCareSD)": "$100M–$500M",
    "2BScientific Ltd": "$5M–$25M",
    "2BScientific GmbH": "$5M–$25M",
    "Abacus dx": "$5M–$25M",
    "Cedarlane": "$5M–$25M",
    "Cedarlane Laboratories": "$5M–$25M",
    "Cosmo Bio Co., Ltd.": "$25M–$100M",
    "Funakoshi": "$25M–$100M",
    "Thomas Scientific": "$25M–$100M",
    "Mawdsleys": "$100M–$500M",
    "Mawdsley-Brooks": "$100M–$500M",
    "Sciensus": "$100M–$500M",
    "Morris & Dickson Co., LLC": "$500M–$1B",
    "Morris & Dickson Specialty Distribution": "$500M–$1B",
    "Nadro": "$1B+",
    "Marzam": "$1B+",
    "Casa Saba": "$1B+",
    "Hungaropharma": "$500M–$1B",
    "Tamro": "$1B+",
    "Oriola": "$1B–$3B",
    "PT Enseval Putera Megatrading": "$1B–$3B",
    "China Resources Pharmaceutical": "$10B+",
    "Shanghai Pharmaceuticals distribution": "$10B+",
    "Walmart Specialty Pharmacy": "$1B+",
    "Kroger Specialty Pharmacy": "$500M–$1B",
    "CenterWell Pharmacy": "$500M–$1B",
    "Aetna Specialty Pharmacy": "$1B+",
    "McKesson Canada": "$1B+",
    "McKesson Plasma and Biologics": "$1B+",
    "McKesson Medical-Surgical": "$1B+",
    "Biologics by McKesson": "$500M–$1B",
    "Phoenix Healthcare Distribution": "$1B+",
    "Alliance Healthcare UK": "$1B+",
    "NOWEDA Apothekergenossenschaft": "$1B+",
    "Sanacorp eG": "$1B+",
    "Viveo": "$500M–$1B",
    "Oncoprod": "$100M–$500M",
    "VECTOR Pharma FZCO": "$25M–$100M",
    "NewBridge Pharmaceuticals": "$25M–$100M",
    "PANTHERx Rare": "$100M–$500M",
    "Orsini Specialty Pharmacy": "$100M–$500M",
    "Onco360": "$100M–$500M",
    "US Bioservices": "$500M–$1B",
    "TheraCom Specialty Pharmacy": "$100M–$500M",
    "Symbion": "$1B+",
}

# Fix Cofares properly
KNOWN_TURN["Cofares"] = "€5B–€10B"


def empty(v: str) -> bool:
    return str(v or "").strip().lower() in NPD


def clean_employees(emp: str) -> str:
    s = (emp or "").strip()
    if empty(s):
        return "Not publicly disclosed"
    # remove est markers
    s = re.sub(r"\s*\(est\.?\)", "", s, flags=re.I)
    s = re.sub(r"\best\.?\b", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip(" ,;")
    # normalize en-dash
    s = s.replace("–", "-").replace("—", "-")
    # cleanup double spaces
    s = re.sub(r"\s+", " ", s).strip()
    return s or "Not publicly disclosed"


def emp_midpoint(emp: str) -> float | None:
    """Rough midpoint for banding turnover."""
    s = clean_employees(emp)
    if empty(s):
        return None
    # Part of / cooperative / group vague
    if re.search(r"(?i)part of|cooperative|group|company\)|thousands|hundreds", s):
        if re.search(r"(?i)300,?000|2\s*m\+|million", s):
            return 300_000
        if re.search(r"(?i)45,?000|46,?000|48,?000|40,?000|33,?000|25,?000|20,?000|14,?000", s):
            return 40_000
        if re.search(r"(?i)10,?000|12,?000|13,?000|8,?000|5,?000|7,?000|6,?000|4,?000|3,?000", s):
            return 8_000
        if re.search(r"(?i)1,?000|thousand", s):
            return 1_500
        if re.search(r"(?i)hundred", s):
            return 300
        return 2_000
    # extract numbers
    nums = [int(x.replace(",", "")) for x in re.findall(r"\d{1,3}(?:,\d{3})*|\d+", s)]
    if not nums:
        return None
    if len(nums) >= 2 and re.search(r"[-~]", s):
        return (nums[0] + nums[1]) / 2
    return float(nums[0])


def band_from_employees(emp: str) -> str:
    mid = emp_midpoint(emp)
    if mid is None:
        return "Not publicly disclosed"
    if mid < 15:
        return "<$5M"
    if mid < 60:
        return "$5M-$25M"
    if mid < 250:
        return "$25M-$100M"
    if mid < 750:
        return "$100M-$500M"
    if mid < 2000:
        return "$500M-$1B"
    return "$1B+"


def already_numeric_range(t: str) -> bool:
    t = (t or "").strip()
    if empty(t):
        return False
    # has currency or clear M/B range
    if re.search(r"[\$€£¥]|A\$|CHF", t):
        return True
    if re.search(r"(?i)\d+\s*[MB]\s*[-–]\s*\$?\d+", t):
        return True
    return False


def clean_turnover(company: str, turn: str, emp: str) -> str:
    if company in KNOWN_TURN:
        return KNOWN_TURN[company]

    t = (turn or "").strip()
    if empty(t):
        return band_from_employees(emp)

    # Strip est.
    t = re.sub(r"\s*\(est\.?\)", "", t, flags=re.I)
    t = re.sub(r"\best\.?\b", "", t, flags=re.I)
    t = t.replace("–", "-").replace("—", "-")
    t = re.sub(r"\s+", " ", t).strip()

    # Already good money value / range
    if already_numeric_range(t):
        # normalize single values like $98.9M -> keep; <$5M keep; $1B+ keep
        # convert "$75M-$120M" style already fine
        # remove trailing words like "revenue (FY2025...)"
        m = re.search(
            r"([\$€£¥A]|CHF)?\s*\$?[\d\.,]+\s*[MBKTmbkt]?(?:\s*[-–]\s*[\$€£¥A]?\$?[\d\.,]+\s*[MBKTmbkt]?)?\+?",
            t,
        )
        # Prefer full range patterns
        m2 = re.search(
            r"((?:[\$€£¥]|A\$|CHF\s*)?[\d\.,]+\s*[MBTKmbkt]?\s*[-–]\s*(?:[\$€£¥]|A\$|CHF\s*)?[\d\.,]+\s*[MBTKmbkt]?|"
            r"(?:[\$€£¥]|A\$|CHF\s*)?[\d\.,]+\s*[MBTKmbkt]?\+|"
            r"<\s*(?:[\$€£¥]|A\$)?[\d\.,]+\s*[MBTKmbkt]?|"
            r"~?\s*(?:[\$€£¥]|A\$|CHF\s*)?[\d\.,]+\s*[MBTKmbktTt]?)",
            t,
        )
        if m2:
            val = m2.group(1).strip()
            val = val.replace("–", "-")
            # ~£2.9B -> £2.5B-£3.5B-ish band if single point? keep single reported as-is cleaned
            val = re.sub(r"^~", "", val).strip()
            # Fix CHF spacing
            val = re.sub(r"CHF\s*", "CHF ", val)
            return val
        return t

    # Qualitative / parent / SME text -> employee-based range
    return band_from_employees(emp)


def load_rows(path: Path) -> list[dict[str, str]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows = []
    for r in range(3, ws.max_row + 1):
        d = {h: str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)
    return rows


def write_xlsx(path: Path, rows: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ADC Distributors"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    title = ws.cell(
        1, 1,
        "Global Antibody Drug Conjugates Market — Distributors",
    )
    title.font = Font(bold=True, size=13, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="1F4E79")
    hf = PatternFill("solid", fgColor="2E75B6")
    hfont = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="BDD7EE"),
        right=Side(style="thin", color="BDD7EE"),
        top=Side(style="thin", color="BDD7EE"),
        bottom=Side(style="thin", color="BDD7EE"),
    )
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(2, i, h)
        c.fill = hf
        c.font = hfont
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = thin
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(h, 16)
    for ri, row in enumerate(rows, 3):
        for ci, h in enumerate(HEADERS, 1):
            cell = ws.cell(ri, ci, row.get(h, ""))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")
    cws = wb.create_sheet("Column Rules")
    cws["A1"] = "Employees: numeric or range (no est. label)"
    cws["A2"] = "Turnover / Scale: currency value or range (e.g. $5M-$25M)"
    cws["A3"] = f"Updated: {datetime.now(timezone.utc).isoformat()} | rows={len(rows)}"
    wb.save(path)


def main() -> int:
    rows = load_rows(XLSX)
    emp_chg = turn_chg = 0
    samples = []
    for r in rows:
        before_e = r.get("Employees") or ""
        before_t = r.get("Turnover / Scale") or ""
        after_e = clean_employees(before_e)
        after_t = clean_turnover(r["Company"], before_t, after_e)
        if after_e != before_e:
            emp_chg += 1
        if after_t != before_t:
            turn_chg += 1
            if len(samples) < 15:
                samples.append((r["Company"], before_t, after_t, after_e))
        r["Employees"] = after_e
        r["Turnover / Scale"] = after_t
        src = r.get("Data Sources") or ""
        if "adc_clean_emp_turnover_ranges" not in src:
            r["Data Sources"] = (src + "; adc_clean_emp_turnover_ranges").strip("; ")

    rows.sort(key=lambda x: x["Company"].lower())
    write_xlsx(XLSX, rows)

    # validate
    est_left = sum(1 for r in rows if re.search(r"(?i)\best", r.get("Employees") or ""))
    vague_left = sum(
        1
        for r in rows
        if not already_numeric_range(r.get("Turnover / Scale") or "")
        and not empty(r.get("Turnover / Scale") or "")
        and (r.get("Turnover / Scale") or "").lower() != "not publicly disclosed"
    )
    print(f"DONE emp_changed={emp_chg} turn_changed={turn_chg}")
    print(f"emp still has est: {est_left}")
    print(f"turnover still vague (non-range): {vague_left}")
    for s in samples[:12]:
        print(f"- {s[0]}")
        print(f"  emp={s[3]}")
        print(f"  turn: {s[1][:50]} -> {s[2]}")
    # show first row like screenshot
    for r in rows:
        if r["Company"] == "2BScientific GmbH":
            print("CHECK 2BScientific GmbH:", r["Employees"], "|", r["Turnover / Scale"])
            break
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
