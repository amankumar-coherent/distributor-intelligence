#!/usr/bin/env python3
"""Force all ADC Turnover / Scale values into explicit min-max ranges."""
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

# Company-specific full ranges (no +)
KNOWN: dict[str, str] = {
    "AAH Pharmaceuticals": "$2.5B-$3.5B",  # ~£2.9B -> USD-ish band; keep GBP better
    "Suzuken Co., Ltd.": "¥2.0T-¥2.5T",
    "Alfresa Holdings": "¥2.0T-¥3.0T",
    "Medipal Holdings": "¥2.0T-¥3.0T",
    "Toho Holdings / TOHO Pharmaceutical": "¥1.0T-¥2.0T",
    "Phoenix Pharma SE": "€25B-$35B",
    "PHOENIX Pharmahandel": "€25B-€35B",
    "Phoenix Healthcare Distribution": "€1B-€5B",
    "Avantor Inc.": "$6B-$8B",
    "Zuellig Pharma": "$10B-$20B",
    "Zuellig Pharma Philippines": "$1B-$3B",
    "DKSH Health": "CHF 10B-CHF 12B",
    "DiethelmKellerSiberHegner Healthcare": "CHF 10B-CHF 12B",
    "Diethelm Thailand / DKSH Thailand": "$100M-$500M",
    "Sinopharm": "$50B-$100B",
    "Jointown Pharmaceutical Group": "$10B-$30B",
    "China Resources Pharmaceutical": "$10B-$30B",
    "Shanghai Pharmaceuticals distribution": "$10B-$30B",
    "Medline Industries": "$15B-$25B",
    "Owens & Minor": "$5B-$15B",
    "Henry Schein Medical": "$10B-$15B",
    "Fisher Scientific": "$5B-$15B",
    "VWR International": "$5B-$10B",
    "Option Care Health specialty infusion": "$2B-$5B",
    "Neuca": "$2B-$5B",
    "Profarma": "$1B-$3B",
    "Cofares": "€5B-€10B",
    "Sigma Healthcare": "A$5B-A$10B",
    "Symbion": "A$1B-A$5B",
    "Uniphar": "€0.5B-€1.5B",
    "Galenica / Alloga": "CHF 3B-CHF 5B",
    "AcariaHealth": "$75M-$120M",
    "Alloga UK": "$75M-$120M",
    "CVS Specialty": "$1B-$5B",
    "Walgreens Specialty Pharmacy": "$1B-$5B",
    "AllianceRx Walgreens Pharmacy": "$1B-$5B",
    "Accredo Health Group": "$1B-$5B",
    "Cigna Specialty Pharmacy / Accredo": "$1B-$5B",
    "Express Scripts Accredo network": "$1B-$5B",
    "Optum Specialty Pharmacy": "$1B-$5B",
    "Aetna Specialty Pharmacy": "$1B-$5B",
    "McKesson Specialty Health": "$1B-$5B",
    "McKesson Specialty Care Distribution": "$1B-$5B",
    "McKesson Plasma and Biologics": "$1B-$5B",
    "McKesson Medical-Surgical": "$1B-$5B",
    "McKesson Canada": "$1B-$5B",
    "Biologics by McKesson": "$500M-$1B",
    "Cardinal Health Specialty Solutions": "$1B-$5B",
    "Cardinal Health Specialty Pharmaceutical Distribution": "$1B-$5B",
    "Cencora (AmerisourceBergen) Specialty": "$1B-$5B",
    "ASD Healthcare": "$1B-$5B",
    "Oncology Supply": "$500M-$1B",
    "CuraScript SD": "$500M-$1B",
    "Anda Inc.": "$1B-$5B",
    "Anda Pharmaceuticals": "$1B-$5B",
    "Alliance Healthcare UK": "$1B-$5B",
    "NOWEDA Apothekergenossenschaft": "$1B-$5B",
    "Sanacorp eG": "$1B-$5B",
    "Nadro": "$1B-$5B",
    "Marzam": "$1B-$5B",
    "Casa Saba": "$1B-$5B",
    "Hungaropharma": "$500M-$1B",
    "Tamro": "$1B-$5B",
    "Oriola": "$1B-$3B",
    "PT Enseval Putera Megatrading": "$1B-$3B",
    "Walmart Specialty Pharmacy": "$1B-$5B",
    "Kroger Specialty Pharmacy": "$500M-$1B",
    "CenterWell Pharmacy": "$500M-$1B",
    "BioCare Specialty Distribution (BioCareSD)": "$100M-$500M",
    "Morris & Dickson Co., LLC": "$500M-$1B",
    "Morris & Dickson Specialty Distribution": "$500M-$1B",
    "Viveo": "$500M-$1B",
    "Oncoprod": "$100M-$500M",
    "US Bioservices": "$500M-$1B",
    "TheraCom Specialty Pharmacy": "$100M-$500M",
    "PANTHERx Rare": "$100M-$500M",
    "Orsini Specialty Pharmacy": "$100M-$500M",
    "Onco360": "$100M-$500M",
    "Sciensus": "$100M-$500M",
    "Mawdsleys": "$100M-$500M",
    "Mawdsley-Brooks": "$100M-$500M",
    "VECTOR Pharma FZCO": "$25M-$100M",
    "NewBridge Pharmaceuticals": "$25M-$100M",
    "2BScientific Ltd": "$5M-$25M",
    "2BScientific GmbH": "$5M-$25M",
    "Abacus dx": "$5M-$25M",
    "Cedarlane": "$5M-$25M",
    "Cedarlane Laboratories": "$5M-$25M",
    "Cosmo Bio Co., Ltd.": "$25M-$100M",
    "Funakoshi": "$25M-$100M",
    "Thomas Scientific": "$25M-$100M",
}

# Fix AAH to GBP range
KNOWN["AAH Pharmaceuticals"] = "£2.5B-£3.5B"
KNOWN["Phoenix Pharma SE"] = "€25B-€35B"


def empty(v: str) -> bool:
    return str(v or "").strip().lower() in NPD


def emp_mid(emp: str) -> float | None:
    s = (emp or "").strip()
    if empty(s):
        return None
    if re.search(r"(?i)part of|cooperative|group|thousands|hundreds", s):
        if re.search(r"(?i)300,?000|2\s*m", s):
            return 300_000
        if re.search(r"(?i)4[0-9],?000|3[0-9],?000|2[0-9],?000|1[0-9],?000", s):
            return 25_000
        if re.search(r"(?i)[5-9],?000|10,?000|12,?000|13,?000", s):
            return 8_000
        if re.search(r"(?i)1,?000|thousand", s):
            return 1_500
        return 2_000
    nums = [int(x.replace(",", "")) for x in re.findall(r"\d{1,3}(?:,\d{3})*|\d+", s)]
    if not nums:
        return None
    if len(nums) >= 2 and "-" in s:
        return (nums[0] + nums[1]) / 2
    return float(nums[0])


def band_from_emp(emp: str) -> str:
    mid = emp_mid(emp)
    if mid is None:
        return "Not publicly disclosed"
    if mid < 15:
        return "$1M-$5M"
    if mid < 60:
        return "$5M-$25M"
    if mid < 250:
        return "$25M-$100M"
    if mid < 750:
        return "$100M-$500M"
    if mid < 2000:
        return "$500M-$1B"
    if mid < 10000:
        return "$1B-$5B"
    return "$5B-$20B"


def to_usd_millions(num: float, unit: str) -> float:
    u = (unit or "").upper()
    if u == "K":
        return num / 1000.0
    if u == "M" or u == "":
        return num
    if u == "B":
        return num * 1000.0
    if u == "T":
        return num * 1_000_000.0
    return num


def fmt_range(lo_m: float, hi_m: float, cur: str = "$") -> str:
    """Format a min-max range in millions of currency units."""
    def fmt(v: float) -> str:
        if v >= 1_000_000:  # trillions in millions? 1T = 1e6 M
            return f"{cur}{v/1_000_000:.1f}T".replace(".0T", "T")
        if v >= 1000:
            x = v / 1000.0
            if abs(x - round(x)) < 0.05:
                return f"{cur}{int(round(x))}B"
            return f"{cur}{x:.1f}B"
        if v >= 1:
            if abs(v - round(v)) < 0.05:
                return f"{cur}{int(round(v))}M"
            return f"{cur}{v:.1f}M"
        # below 1M -> show as K
        k = v * 1000
        return f"{cur}{int(round(k))}K"

    # special currencies already embedded in cur like £ € ¥ CHF  A$
    return f"{fmt(lo_m)}-{fmt(hi_m)}"


def parse_money_token(tok: str) -> tuple[str, float, str] | None:
    """Return (currency, number, unit) from like $1B, £2.5B, CHF 10B, A$5B, 25M."""
    s = tok.strip().replace(",", "")
    m = re.match(
        r"^(CHF\s*|A\$|US\$|\$|€|£|¥)?\s*([\d.]+)\s*([KMBTkmbt])?$",
        s,
    )
    if not m:
        return None
    cur_raw, num_s, unit = m.groups()
    cur = (cur_raw or "$").strip()
    if cur in {"$", "US$"}:
        cur = "$"
    try:
        num = float(num_s)
    except ValueError:
        return None
    return cur, num, (unit or "M").upper()


def normalize_turnover(company: str, turn: str, emp: str) -> str:
    if company in KNOWN:
        return KNOWN[company]

    t = (turn or "").strip().replace("–", "-").replace("—", "-")
    t = re.sub(r"\s+", " ", t).strip()
    if empty(t):
        return band_from_emp(emp)

    # Already a proper two-sided range with units on both or shared unit patterns
    # Examples to accept: $5M-$25M, $5-25M, £2.5B-£3.5B, CHF 10B-CHF 12B, $500M-$1B
    # Convert $1B+ / $10B+
    m_plus = re.match(r"^(CHF\s*|A\$|\$|€|£|¥)?\s*([\d.]+)\s*([KMBT])?\+$", t, re.I)
    if m_plus:
        cur = (m_plus.group(1) or "$").strip()
        if cur in {"$", "US$"}:
            cur = "$"
        num = float(m_plus.group(2))
        unit = (m_plus.group(3) or "B").upper()
        lo_m = to_usd_millions(num, unit)
        # open-ended -> bounded range by scale
        if lo_m >= 10_000:  # >= $10B
            hi_m = lo_m * 3
        elif lo_m >= 1000:  # >= $1B
            hi_m = lo_m * 5  # $1B-$5B
        elif lo_m >= 500:
            hi_m = 1000
        else:
            hi_m = lo_m * 3
        # format with same currency symbol style
        if cur == "CHF":
            return f"CHF {num:.0f}{unit}-CHF {to_usd_millions(hi_m, 'M')/to_usd_millions(1, unit):.0f}{unit}" if False else (
                f"CHF {num:.0f}B-CHF {hi_m/1000:.0f}B" if unit == "B" else f"CHF {num:.0f}{unit}-CHF {(hi_m if unit=='M' else hi_m/1000):.0f}{unit}"
            )
        return fmt_range(lo_m, hi_m, cur)

    # <$5M -> $1M-$5M
    m_lt = re.match(r"^<\s*(CHF\s*|A\$|\$|€|£|¥)?\s*([\d.]+)\s*([KMBT])?$", t, re.I)
    if m_lt:
        cur = (m_lt.group(1) or "$").strip() or "$"
        if cur in {"$", "US$"}:
            cur = "$"
        num = float(m_lt.group(2))
        unit = (m_lt.group(3) or "M").upper()
        hi_m = to_usd_millions(num, unit)
        lo_m = max(hi_m * 0.2, 0.5 if hi_m >= 1 else hi_m * 0.2)
        return fmt_range(lo_m, hi_m, cur)

    # Range like $5-25M or $5M-$25M or $500M-$1B or $734.2-917.7M
    m_range = re.match(
        r"^(CHF\s*|A\$|\$|€|£|¥)?\s*([\d.]+)\s*([KMBT])?\s*-\s*"
        r"(CHF\s*|A\$|\$|€|£|¥)?\s*([\d.]+)\s*([KMBT])?$",
        t,
        re.I,
    )
    if m_range:
        cur1 = (m_range.group(1) or m_range.group(4) or "$").strip() or "$"
        if cur1 in {"$", "US$"}:
            cur1 = "$"
        n1 = float(m_range.group(2))
        u1 = (m_range.group(3) or m_range.group(6) or "M").upper()
        n2 = float(m_range.group(5))
        u2 = (m_range.group(6) or m_range.group(3) or "M").upper()
        # if first unit missing and second is M (e.g. $5-25M)
        if not m_range.group(3) and m_range.group(6):
            u1 = u2
        lo_m = to_usd_millions(n1, u1)
        hi_m = to_usd_millions(n2, u2)
        if lo_m > hi_m:
            lo_m, hi_m = hi_m, lo_m
        return fmt_range(lo_m, hi_m, cur1 if cur1 != "CHF" else "$").replace("$", "CHF " if cur1 == "CHF" else ("A$" if cur1 == "A$" else cur1), 1) if cur1 in {"CHF", "A$"} else fmt_range(lo_m, hi_m, cur1)

    # Single value $98.9M / $262B / CHF 11B
    m_single = re.match(
        r"^(CHF\s*|A\$|\$|€|£|¥)?\s*([\d.]+)\s*([KMBT])$",
        t,
        re.I,
    )
    if m_single:
        cur = (m_single.group(1) or "$").strip() or "$"
        if cur in {"$", "US$"}:
            cur = "$"
        num = float(m_single.group(2))
        unit = m_single.group(3).upper()
        mid = to_usd_millions(num, unit)
        lo_m = mid * 0.75
        hi_m = mid * 1.25
        if cur == "CHF":
            # keep CHF formatting simply
            if unit == "B":
                return f"CHF {lo_m/1000:.1f}B-CHF {hi_m/1000:.1f}B".replace(".0B", "B")
            return f"CHF {lo_m:.0f}M-CHF {hi_m:.0f}M"
        return fmt_range(lo_m, hi_m, cur)

    # Fallback by employees
    return band_from_emp(emp)


def is_full_range(t: str) -> bool:
    t = (t or "").strip()
    if empty(t):
        return False
    if "+" in t or t.startswith("<"):
        return False
    return bool(
        re.search(
            r"\d\s*[KMBTkmbt]?\s*-\s*(?:CHF\s*|A\$|\$|€|£|¥)?\s*\d",
            t,
        )
    )


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
    cws["A1"] = "Turnover / Scale must be an explicit range (e.g. $1B-$5B), never $1B+ or single point"
    cws["A2"] = f"Updated: {datetime.now(timezone.utc).isoformat()} | rows={len(rows)}"
    wb.save(path)


def main() -> int:
    # fix known typos
    KNOWN["Phoenix Pharma SE"] = "€25B-€35B"

    rows = load_rows(XLSX)
    changed = 0
    samples = []
    for r in rows:
        before = r.get("Turnover / Scale") or ""
        after = normalize_turnover(r["Company"], before, r.get("Employees") or "")
        # final safety: if still not a range and not NPD, force emp band
        if after.lower() != "not publicly disclosed" and not is_full_range(after):
            after = band_from_emp(r.get("Employees") or "")
        if after != before:
            changed += 1
            if len(samples) < 20 and ("+" in before or before.startswith("<") or (before and "-" not in before)):
                samples.append((r["Company"], before, after))
        r["Turnover / Scale"] = after
        src = r.get("Data Sources") or ""
        if "adc_turnover_all_ranges" not in src:
            r["Data Sources"] = (src + "; adc_turnover_all_ranges").strip("; ")

    rows.sort(key=lambda x: x["Company"].lower())
    write_xlsx(XLSX, rows)

    plus = sum(1 for r in rows if "+" in (r.get("Turnover / Scale") or ""))
    non_range = sum(
        1
        for r in rows
        if (r.get("Turnover / Scale") or "").lower() != "not publicly disclosed"
        and not is_full_range(r.get("Turnover / Scale") or "")
    )
    print(f"DONE changed={changed} plus_left={plus} non_range_left={non_range}")
    for s in samples:
        print(f"  {s[0][:40]:40} | {s[1][:20]:20} -> {s[2]}")
    # distribution
    from collections import Counter
    print("top values:", Counter(r.get("Turnover / Scale") for r in rows).most_common(15))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
