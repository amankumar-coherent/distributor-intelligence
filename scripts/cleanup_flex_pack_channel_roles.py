"""Channel-fit cleanup for Global Flexible Packaging FINAL.

Removes:
- Distribution Type == Manufacturer (OEMs)
- Verified non-channel / wrong-market / converter-OEM blanks
- Lab scientific distributors (not flex-packaging channel partners)

Keeps Specialty Distributor + verified packaging merchants/distributors.
Sets blank Distribution Type to Specialty Distributor for kept channel rows.
"""
from __future__ import annotations

import json
import shutil
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

XLSX = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\channel_role_cleanup_audit.json"
)

# Explicit removals beyond Distribution Type=Manufacturer
REMOVE_BY_NAME = {
    # OEMs / converters among blank Dist Type
    "Flexipol Packaging",
    "Packflex",
    "Bison Bag Co.",
    "Poly Pak America",
    "Pouch Partners",
    "Dynapack Corporation",
    "Flex-Pack Inc.",
    "Al Jazeera Packaging",
    "Arabian Packaging Co.",
    "Emirates Packaging Industries",
    "Shree Ganesh Packaging",
    "Packman Packaging",
    "Packaging Depot",  # website points to TricorBraun Flex manufacturer
    "Global Packaging Inc.",  # flexible film manufacturer
    "Siam Pack Industrial Co. Ltd.",
    "Siam Packaging Co. Ltd.",
    # Wrong market / not packaging materials channel
    "Delta Packaging Systems",  # coding/marking equipment
    "Sanko Co. Ltd.",  # fasteners
    "Toyo Bussan Co. Ltd.",  # safety/uniforms
    "TCI",  # freight/logistics
    "Carr Group",  # technical textiles
    "Packaging Exchange",  # reusable assets / warehouse assets
    "Hitech Packaging",  # dead login URL
    # Lab scientific (not flex packaging channel partners)
    "Fisher Scientific Packaging Dist",
    "Thomas Scientific Packaging Dist",
    "VWR Packaging Lab Dist",
    # Unverified Indian/generic manufacturer-pattern blanks (no channel evidence)
    "Arihant Packaging",
    "Champion Packaging",
    "Gayatri Packaging",
    "Hindustan Packaging",
    "Jain Packaging",
    "Karan Packaging",
    "Mangalam Packaging",
    "Om Packaging",
    "Quality Packaging",
    "Raj Packaging",
    "Shree Krishna Packaging",
    "Shree Packaging",
    "Shree Ram Packaging",
    "Sri Packaging",
    "Tirupati Packaging",
    "Unique Packaging",
    "Vardhman Packaging",
    "Vishnu Packaging",
    # Ambiguous converter / manufacturer-leaning blanks
    "Packaging Resources",
    "Packaging Products Corporation",
    "Nelson Packaging",
    "IPL Packaging",
    "Mega P & L",
    "Bulk Packaging",
    "Great Pacific Packaging",
    "H&H Packaging",
    "Plastic Bag Partners",
    "Kiwi Packaging",
    "Pacific Packaging",
    "Tasman Packaging",
    "Total Packaging",
    "Allied Packaging",
    "Global Packaging",
    "Packaging House",
    "Packaging Partners",  # brokerage/sales org — borderline; keep? User asked channel partner — KEEP actually
}

# Override: keep these as channel even if listed above by mistake
KEEP_ALWAYS = {
    "Antalis Packaging",
    "Associated Bag",
    "Imperial Supplies",
    "Packaging Warehouse",
    "Packaging2Buy",
    "Packaging Wholesalers",
    "The Packaging Company",
    "Packaging Supply Group",
    "Packaging Source",
    "U.S. Packaging & Wrapping LLC",
    "Diversified Packaging Solutions",
    "Packaging Partners",  # packaging sales/distribution org = channel
    "Allied Packaging",  # AU bags/cups merchant
    "Global Packaging",  # CA packaging merchant
    "Packaging House",
    "Plastic Bag Partners",
    "Great Pacific Packaging",
    "H&H Packaging",
    "Bulk Packaging",
    "Nelson Packaging",
    "Pacific Packaging",
    "Tasman Packaging",
    "Total Packaging",
    "Kiwi Packaging",
    "Mega P & L",
    "IPL Packaging",
    "Packaging Products Corporation",
    "Packaging Resources",
}

for name in KEEP_ALWAYS:
    REMOVE_BY_NAME.discard(name)

backup = XLSX.with_name(XLSX.stem + "_pre_channel_cleanup_backup.xlsx")
shutil.copy2(XLSX, backup)

wb = openpyxl.load_workbook(XLSX)
ws = wb.active
headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
ci = {h: i for i, h in enumerate(headers)}

removed = []
kept_blank_fixed = []

for row_idx in range(ws.max_row, 1, -1):
    name = ws.cell(row=row_idx, column=ci["Company"] + 1).value
    if not name:
        # drop empty rows
        vals = [ws.cell(row=row_idx, column=c + 1).value for c in range(len(headers))]
        if all(v is None or str(v).strip() == "" for v in vals):
            ws.delete_rows(row_idx)
        continue

    dt = ws.cell(row=row_idx, column=ci["Distribution Type"] + 1).value
    dt_s = str(dt or "").strip()
    web = ws.cell(row=row_idx, column=ci["Website"] + 1).value

    reason = None
    if dt_s.lower() == "manufacturer":
        reason = "Distribution Type=Manufacturer (OEM)"
    elif name in REMOVE_BY_NAME:
        reason = "non-channel / wrong-market / unverified OEM pattern"

    if reason:
        removed.append(
            {
                "company": name,
                "website": web,
                "distribution_type": dt_s or None,
                "reason": reason,
            }
        )
        ws.delete_rows(row_idx)
        continue

    # Fill blank Distribution Type for kept channel rows
    if not dt_s and "Distribution Type" in ci:
        ws.cell(row=row_idx, column=ci["Distribution Type"] + 1).value = "Specialty Distributor"
        kept_blank_fixed.append(name)

wb.save(XLSX)

# Recount
wb2 = openpyxl.load_workbook(XLSX, data_only=True)
ws2 = wb2.active
headers2 = [c.value for c in next(ws2.iter_rows(min_row=1, max_row=1))]
ci2 = {h: i for i, h in enumerate(headers2)}
rows = []
for r in ws2.iter_rows(min_row=2, values_only=True):
    if not r[ci2["Company"]]:
        continue
    rows.append(
        {
            "Company": r[ci2["Company"]],
            "Website": r[ci2["Website"]],
            "Distribution Type": r[ci2["Distribution Type"]],
            "Founded": r[ci2["Founded"]],
        }
    )
wb2.close()

from collections import Counter

dt_counts = Counter(str(r["Distribution Type"] or "").strip() or "(blank)" for r in rows)
founded_empty = sum(
    1 for r in rows if r["Founded"] is None or str(r["Founded"]).strip() == ""
)

audit = {
    "ts": datetime.now(timezone.utc).isoformat(),
    "backup": str(backup),
    "removed_count": len(removed),
    "removed": removed,
    "blank_dist_type_filled_to_specialty": kept_blank_fixed,
    "total_rows_after": len(rows),
    "distribution_type_counts": dict(dt_counts),
    "founded_empty_after": founded_empty,
}
AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")

print("backup:", backup)
print("removed:", len(removed))
print("kept_blank_fixed:", len(kept_blank_fixed), kept_blank_fixed)
print("total_after:", len(rows))
print("dist types:", dict(dt_counts))
print("founded_empty:", founded_empty)
print("removed sample:")
for x in removed[:15]:
    print(" -", x["company"], "|", x["reason"])
print(" ...")
for x in removed[-10:]:
    print(" -", x["company"], "|", x["reason"])
