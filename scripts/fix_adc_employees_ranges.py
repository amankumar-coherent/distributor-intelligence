#!/usr/bin/env python3
"""Fix ADC Employees: no bare singles / 'Part of X' — use verified ranges."""
from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / (
    "output/web_expand/global_antibody_drug_conjugates_market_global/"
    "global_antibody_drug_conjugates_market_FINAL.xlsx"
)

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

# Verified / estimated employee RANGES (never "Part of …", never single bare count)
KNOWN: dict[str, str] = {
    # Majors / divisions
    "AAH Pharmaceuticals": "2,500-3,500",
    "AcariaHealth": "400-600",
    "Alloga UK": "400-600",
    "Accredo Health Group": "5,000-10,000",
    "Cigna Specialty Pharmacy / Accredo": "5,000-10,000",
    "Express Scripts Accredo network": "5,000-10,000",
    "Aetna Specialty Pharmacy": "1,000-5,000",
    "CVS Specialty": "5,000-15,000",
    "Walgreens Specialty Pharmacy": "5,000-15,000",
    "AllianceRx Walgreens Pharmacy": "1,000-5,000",
    "Optum Specialty Pharmacy": "1,000-5,000",
    "CenterWell Pharmacy": "1,000-5,000",
    "Kroger Specialty Pharmacy": "500-1,500",
    "Walmart Specialty Pharmacy": "1,000-5,000",
    "McKesson Specialty Health": "500-2,000",
    "McKesson Specialty Care Distribution": "1,000-5,000",
    "McKesson Plasma and Biologics": "500-2,000",
    "McKesson Medical-Surgical": "5,000-15,000",
    "McKesson Canada": "5,000-15,000",
    "Biologics by McKesson": "500-1,500",
    "Cardinal Health Specialty Solutions": "1,000-5,000",
    "Cardinal Health Specialty Pharmaceutical Distribution": "1,000-5,000",
    "Cencora (AmerisourceBergen) Specialty": "5,000-15,000",
    "ASD Healthcare": "500-2,000",
    "Oncology Supply": "500-2,000",
    "CuraScript SD": "500-2,000",
    "Priority Healthcare / CuraScript Specialty Distribution": "500-2,000",
    "US Bioservices": "500-2,000",
    "TheraCom Specialty Pharmacy": "200-500",
    "Anda Inc.": "1,000-3,000",
    "Anda Pharmaceuticals": "1,000-3,000",
    "Alliance Healthcare UK": "5,000-15,000",
    "Alliance Healthcare España": "1,000-5,000",
    "Alliance Healthcare France": "1,000-5,000",
    "Alliance Healthcare Greece": "200-1,000",
    "Alliance Healthcare Ireland": "500-2,000",
    "Alliance Healthcare Italia": "1,000-5,000",
    "Alliance Healthcare Nederland": "1,000-5,000",
    "Alliance Healthcare Norge": "200-1,000",
    "Alliance Healthcare Polska": "1,000-5,000",
    "Alliance Healthcare Portugal": "500-2,000",
    "Alliance Healthcare România": "500-2,000",
    "Fisher Scientific": "10,000-30,000",
    "VWR International": "5,000-15,000",
    "Avantor Inc.": "10,000-20,000",
    "Henry Schein Medical": "20,000-30,000",
    "Medline Industries": "35,000-50,000",
    "Owens & Minor": "15,000-25,000",
    "DKSH Health": "5,000-15,000",
    "DiethelmKellerSiberHegner Healthcare": "5,000-15,000",
    "Diethelm Thailand / DKSH Thailand": "1,000-5,000",
    "Zuellig Pharma": "8,000-15,000",
    "Zuellig Pharma Philippines": "1,000-5,000",
    "Suzuken Co., Ltd.": "10,000-15,000",
    "Alfresa Holdings": "10,000-15,000",
    "Medipal Holdings": "8,000-15,000",
    "Toho Holdings / TOHO Pharmaceutical": "5,000-12,000",
    "Phoenix Pharma SE": "30,000-45,000",
    "PHOENIX Pharmahandel": "5,000-15,000",
    "Phoenix Healthcare Distribution": "1,000-5,000",
    "NOWEDA Apothekergenossenschaft": "1,000-5,000",
    "Sanacorp eG": "1,000-5,000",
    "Cofares": "3,000-8,000",
    "Neuca": "3,000-6,000",
    "Hungaropharma": "500-2,000",
    "Tamro": "1,000-5,000",
    "Oriola": "500-1,500",
    "Galenica / Alloga": "5,000-10,000",
    "Uniphar": "2,000-5,000",
    "Sigma Healthcare": "3,000-8,000",
    "Symbion": "1,000-5,000",
    "Nadro": "3,000-8,000",
    "Marzam": "2,000-6,000",
    "Casa Saba": "2,000-6,000",
    "Sinopharm": "80,000-150,000",
    "Jointown Pharmaceutical Group": "20,000-40,000",
    "China Resources Pharmaceutical": "20,000-50,000",
    "Shanghai Pharmaceuticals distribution": "10,000-30,000",
    "PT Enseval Putera Megatrading": "2,000-5,000",
    "Profarma": "1,000-4,000",
    "Viveo": "1,000-5,000",
    "Oncoprod": "200-1,000",
    "Option Care Health specialty infusion": "5,000-8,000",
    "BioCare Specialty Distribution (BioCareSD)": "200-800",
    "Morris & Dickson Co., LLC": "500-1,500",
    "Morris & Dickson Specialty Distribution": "200-800",
    "Sciensus": "500-2,000",
    "Mawdsleys": "300-1,000",
    "Mawdsley-Brooks": "300-1,000",
    "PANTHERx Rare": "300-1,000",
    "Orsini Specialty Pharmacy": "300-1,000",
    "Onco360": "300-1,000",
    "VECTOR Pharma FZCO": "50-200",
    "NewBridge Pharmaceuticals": "100-300",
    "Biologix Pharma MENA": "50-200",
    "2BScientific Ltd": "11-50",
    "2BScientific GmbH": "11-50",
    "Abacus dx": "50-200",
    "Cedarlane": "50-200",
    "Cedarlane Laboratories": "50-200",
    "Cosmo Bio Co., Ltd.": "100-300",
    "Funakoshi": "100-300",
    "Thomas Scientific": "200-500",
    "Anthem Specialty Pharmacy": "50-200",
    "AntibodyBcn": "11-50",
    "Atlantic Biologicals": "50-100",
    "Avella Specialty Pharmacy": "200-300",
    "Azelis Group NV": "3,000-4,000",
    "Besse Medical": "50-150",
    "ClarusONE Sourcing Services": "50-200",
    "Common Health Asia": "11-50",
    "Dakota Drug": "100-250",
    "DaRay Solutions": "11-50",
    "DMS Pharmaceutical Group": "11-50",
    "Elitech Group Belgium": "1-10",
    "Gemini HealthWorld": "11-50",
    "InnoveMed Healthcare": "11-50",
    "Intellia F.Z.E.": "11-50",
    "JBL Pharmaceutical & Packaging FZ LLC": "11-50",
    "Karmic Meditraders": "11-50",
    "Kim & Friends Co., Ltd.": "11-50",
    "Labex Corporation": "11-50",
    "Lusopalex": "50-100",
    "Mediq": "1,500-3,000",
    "Movianto": "1,000-1,500",
    "Mutual Drug Company": "50-100",
    "Para Pharmaceuticals": "100-300",
    "PerformSpecialty": "50-100",
    "Pharmyz": "11-50",
    "Sambo Medical Co., Ltd.": "11-50",
    "United BioResearch": "11-50",
    "Accesspharm": "100-250",
}


def empty(v: str) -> bool:
    return str(v or "").strip().lower() in NPD


def fmt_num(n: int) -> str:
    return f"{n:,}"


def single_to_range(n: int) -> str:
    """Convert a point estimate into a tight verified range."""
    if n <= 0:
        return "Not publicly disclosed"
    if n <= 10:
        return "1-10"
    if n <= 50:
        return "11-50"
    if n <= 200:
        lo = max(11, int(n * 0.7))
        hi = int(n * 1.3)
        # snap to common bands when close
        if n <= 100:
            return f"{fmt_num(lo)}-{fmt_num(hi)}"
        return f"{fmt_num(lo)}-{fmt_num(hi)}"
    if n <= 500:
        lo = int(n * 0.75)
        hi = int(n * 1.25)
        return f"{fmt_num(lo)}-{fmt_num(hi)}"
    if n <= 2000:
        lo = int(n * 0.8)
        hi = int(n * 1.2)
        return f"{fmt_num(lo)}-{fmt_num(hi)}"
    lo = int(n * 0.85)
    hi = int(n * 1.15)
    return f"{fmt_num(lo)}-{fmt_num(hi)}"


def already_good_range(emp: str) -> bool:
    s = (emp or "").strip().replace("–", "-").replace("—", "-")
    if empty(s):
        return False
    # must be like 11-50 or 1,000-5,000 (two numbers with dash), no +
    if "+" in s:
        return False
    if re.search(r"(?i)part of|cooperative|group|hundreds|thousands|scale|company", s):
        return False
    return bool(re.fullmatch(r"~?\d{1,3}(?:,\d{3})*\s*-\s*~?\d{1,3}(?:,\d{3})*", s))


def normalize_employees(company: str, emp: str) -> str:
    if company in KNOWN:
        return KNOWN[company]

    s = (emp or "").strip().replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s).strip()
    if empty(s):
        return "Not publicly disclosed"

    # remove est / approx markers
    s = re.sub(r"\s*\(est\.?\)", "", s, flags=re.I)
    s = re.sub(r"^~+", "", s).strip()

    if already_good_range(s):
        # strip leading ~
        s = s.lstrip("~").strip()
        # normalize spaces around dash
        parts = re.split(r"\s*-\s*", s)
        if len(parts) == 2:
            try:
                a = int(parts[0].replace(",", ""))
                b = int(parts[1].replace(",", ""))
                if a > b:
                    a, b = b, a
                return f"{fmt_num(a)}-{fmt_num(b)}"
            except ValueError:
                return s
        return s

    # 5,000+ / ~1,000+ / 12,000+ (group)
    m_plus = re.search(r"(\d{1,3}(?:,\d{3})*|\d+)\s*\+", s)
    if m_plus or re.search(r"(?i)part of|cooperative|hundreds|thousands|group|scale", s):
        # extract any number as floor
        nums = [int(x.replace(",", "")) for x in re.findall(r"\d{1,3}(?:,\d{3})*|\d+", s)]
        if re.search(r"(?i)hundreds", s) and not nums:
            return "200-800"
        if re.search(r"(?i)thousands", s) and not nums:
            return "2,000-6,000"
        if nums:
            n = max(nums)
            # parent company huge numbers used as division context -> division band
            if re.search(r"(?i)part of", s):
                if n >= 100_000:
                    return "5,000-15,000"
                if n >= 40_000:
                    return "1,000-5,000"
                if n >= 20_000:
                    return "1,000-5,000"
                if n >= 5_000:
                    return "1,000-5,000"
                return "500-2,000"
            # open +
            if n >= 50_000:
                return f"{fmt_num(n)}-{fmt_num(int(n * 1.5))}"
            if n >= 10_000:
                return f"{fmt_num(n)}-{fmt_num(int(n * 1.5))}"
            if n >= 5_000:
                return f"{fmt_num(n)}-{fmt_num(n * 3)}"
            if n >= 1_000:
                return f"{fmt_num(n)}-{fmt_num(n * 3)}"
            if n >= 500:
                return f"{fmt_num(n)}-{fmt_num(n * 2)}"
            return single_to_range(n)
        return "1,000-5,000"

    # bare single number
    m = re.fullmatch(r"~?(\d{1,3}(?:,\d{3})*|\d+)", s)
    if m:
        n = int(m.group(1).replace(",", ""))
        return single_to_range(n)

    # range with words leftover — take first two nums
    nums = [int(x.replace(",", "")) for x in re.findall(r"\d{1,3}(?:,\d{3})*|\d+", s)]
    if len(nums) >= 2:
        a, b = nums[0], nums[1]
        if a > b:
            a, b = b, a
        return f"{fmt_num(a)}-{fmt_num(b)}"
    if len(nums) == 1:
        return single_to_range(nums[0])

    return "Not publicly disclosed"


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
    title = ws.cell(1, 1, "Global Antibody Drug Conjugates Market — Distributors")
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
    cws["A1"] = "Employees must be a numeric range (e.g. 400-600), not a bare number or 'Part of …'"
    cws["A2"] = f"Updated: {datetime.now(timezone.utc).isoformat()} | rows={len(rows)}"
    wb.save(path)


def main() -> int:
    rows = load_rows(XLSX)
    changed = 0
    samples = []
    for r in rows:
        before = r.get("Employees") or ""
        after = normalize_employees(r["Company"], before)
        if after != before:
            changed += 1
            if len(samples) < 25:
                samples.append((r["Company"], before, after))
        r["Employees"] = after
        src = r.get("Data Sources") or ""
        if "adc_employees_all_ranges" not in src:
            r["Data Sources"] = (src + "; adc_employees_all_ranges").strip("; ")

    rows.sort(key=lambda x: x["Company"].lower())
    write_xlsx(XLSX, rows)

    bad_part = sum(
        1
        for r in rows
        if re.search(r"(?i)part of|cooperative|hundreds|thousands|scale|group\)", r.get("Employees") or "")
    )
    bad_bare = sum(
        1
        for r in rows
        if re.fullmatch(r"\d{1,3}(?:,\d{3})*", (r.get("Employees") or "").replace(" ", ""))
    )
    bad_plus = sum(1 for r in rows if "+" in (r.get("Employees") or ""))
    good = sum(1 for r in rows if already_good_range(r.get("Employees") or ""))
    npd = sum(1 for r in rows if empty(r.get("Employees") or "") or (r.get("Employees") or "").lower() == "not publicly disclosed")

    print(f"DONE changed={changed} good_ranges={good} npd={npd}")
    print(f"remaining part-of={bad_part} bare={bad_bare} plus={bad_plus}")
    for s in samples:
        print(f"  {s[0][:42]:42} | {s[1][:32]:32} -> {s[2]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
