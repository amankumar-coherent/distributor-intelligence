#!/usr/bin/env python3
import re
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fill_adc_employees_turnover import load_rows, write_xlsx, empty

XLSX = Path(__file__).resolve().parents[1] / (
    "output/web_expand/global_antibody_drug_conjugates_market_global/"
    "global_antibody_drug_conjugates_market_FINAL.xlsx"
)
rows = load_rows(XLSX)
fixes = 0
for r in rows:
    t = (r.get("Turnover / Scale") or "").strip()
    e = (r.get("Employees") or "").strip()
    if re.match(r"^\$\d(\.\d)?\s*\(est\.\)$", t) or t in {"$1 (est.)", "$1", "1 (est.)"}:
        print("fix turn", r["Company"], t)
        r["Turnover / Scale"] = "Not publicly disclosed"
        fixes += 1
    if r["Company"] == "Assay Matrix Pty Ltd" and "1,000" in e:
        r["Employees"] = "11-50 (est.)"
        r["Turnover / Scale"] = "$5-25M (est.)"
        fixes += 1
        print("fix Assay Matrix")
    if "Red Oak" in r["Company"]:
        r["Turnover / Scale"] = "Part of McKesson / Red Oak sourcing (scale n/d)"
        fixes += 1
        print("fix Red Oak")

rows.sort(key=lambda x: x["Company"].lower())
write_xlsx(XLSX, rows)
emp_f = sum(
    1
    for r in rows
    if not empty(r.get("Employees")) and r.get("Employees", "").lower() != "not publicly disclosed"
)
turn_f = sum(
    1
    for r in rows
    if not empty(r.get("Turnover / Scale"))
    and r.get("Turnover / Scale", "").lower() != "not publicly disclosed"
)
print(f"fixes={fixes} emp_filled={emp_f} turn_filled={turn_f} total={len(rows)}")
for name in [
    "AAH Pharmaceuticals",
    "AcariaHealth",
    "Alloga UK",
    "Suzuken Co., Ltd.",
    "Zuellig Pharma",
    "2BScientific Ltd",
]:
    for r in rows:
        if r["Company"] == name:
            print(name, "|", r["Employees"], "|", r["Turnover / Scale"])
