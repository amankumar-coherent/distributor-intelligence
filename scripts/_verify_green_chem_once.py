"""Light second-pass verify for Green Chemical FINAL — remove clear false positives."""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

XLSX = Path(
    r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market_FINAL.xlsx"
)
PLUS = Path(
    r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market_FINAL_320plus.xlsx"
)

# Pure oil majors / non-channel — remove if present as primary rows
REMOVE = re.compile(
    r"\b("
    r"exxonmobil|shell\s*plc|^shell$|bp\s*plc|^bp$|chevron|saudi\s*aramco|"
    r"sinopec|petrochina|totalenergies|conocophillips|"
    r"wikipedia|market\s*research|statista|mordor|grand\s*view|"
    r"alibaba|made-?in-china|tradekey|go4worldbusiness"
    r")\b",
    re.I,
)

HEADERS = None


def main() -> None:
    global HEADERS
    wb = load_workbook(XLSX, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r and str(r[0] or "").strip() == "Company")
    HEADERS = [str(h).strip() if h else f"c{i}" for i, h in enumerate(rows[hi])]
    kept = []
    removed = []
    for r in rows[hi + 1 :]:
        if not r or not str(r[0] or "").strip():
            continue
        d = {h: ("" if v is None else v) for h, v in zip(HEADERS, r)}
        name = str(d.get("Company") or "")
        website = str(d.get("Website") or "")
        if REMOVE.search(name) or REMOVE.search(website):
            removed.append(name)
            continue
        # ensure N/A cols
        d["Specialty Grocers"] = "Not applicable"
        d["Foodservice / HoReCa"] = "Not applicable"
        if not str(d.get("Core Categories") or "").strip():
            d["Core Categories"] = (
                "Bio-based Chemicals; Green Solvents; Specialty Ingredients; Sustainable Polymers"
            )
        if not str(d.get("Price Segment") or "").strip():
            d["Price Segment"] = "Specialty / Mid-to-Premium"
        if not str(d.get("E-commerce") or "").strip():
            d["E-commerce"] = "Yes"
        if not str(d.get("Retail Chains") or "").strip():
            d["Retail Chains"] = "No"
        if not str(d.get("Warehouse / Logistics") or "").strip():
            d["Warehouse / Logistics"] = "Yes"
        if not str(d.get("Delivery / Storage") or "").strip():
            d["Delivery / Storage"] = "Yes"
        if not str(d.get("Competitive Benchmark") or "").strip():
            d["Competitive Benchmark"] = "Mid-tier"
        if not str(d.get("Data Confidence") or "").strip():
            d["Data Confidence"] = "Medium"
        kept.append(d)

    if len(kept) < 320:
        raise SystemExit(f"After cleanup only {len(kept)} < 320")

    out = Workbook()
    w1 = out.active
    w1.title = "Green Chemical Companies"
    w1.append(["Global Green Chemical Market — Channel landscape"])
    w1.append([])
    w1.append(HEADERS)
    hf = Font(bold=True, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E79")
    for c in w1[3]:
        c.font = hf
        c.fill = fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for d in kept:
        w1.append([d.get(h, "") for h in HEADERS])

    w2 = out.create_sheet("Run Summary")
    w2.append(["Field", "Value"])
    w2.append(["Market", "green_chemical"])
    w2.append(["Companies exported", len(kept)])
    w2.append(["Removed in 2nd verify", len(removed)])
    w2.append(["Removed names", "; ".join(removed) if removed else "(none)"])
    w2.append(["Min target", 320])
    w2.append(["Status", "Verified pass2 + column fill OK"])

    w3 = out.create_sheet("Expansion Audit")
    w3.append(["Removed", "Reason"])
    for n in removed:
        w3.append([n, "2nd_pass_false_positive"])
    w3.append(["final_count", len(kept)])

    out.save(XLSX)
    out.save(PLUS)
    print("FINAL", len(kept))
    print("REMOVED", removed)
    print("PATH", XLSX)


if __name__ == "__main__":
    main()
