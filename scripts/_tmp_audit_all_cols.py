"""Audit blank rates across all FINAL Excel columns for 3 markets."""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
KEY_COLS = [
    "Company",
    "Founded",
    "Headquarters",
    "Cities / Regions",
    "Ownership",
    "Business Type",
    "Employees",
    "Turnover / Scale",
    "Contact Person",
    "Role",
    "Email",
    "Phone / WhatsApp",
    "LinkedIn",
    "Website",
    "Core Categories",
    "Specialty Focus",
    "Price Segment",
    "Key Brands Represented",
]


def find_header(ws):
    for r in range(1, 8):
        vals = [c.value for c in ws[r]]
        if any(v and str(v).strip().lower() == "company" for v in vals):
            return r, vals
    return 1, [c.value for c in ws[1]]


def main() -> None:
    for mid in [
        "global_smartwatch_market_global",
        "global_green_chemical_market_global",
        "global_firewall_market_global",
    ]:
        short = re.sub(r"_global$", "", mid)
        xp = ROOT / mid / f"{short}_FINAL.xlsx"
        wb = load_workbook(xp, read_only=True, data_only=True)
        ws = wb.active
        hr, headers = find_header(ws)
        hmap = {str(h).strip(): i for i, h in enumerate(headers) if h}
        ci = hmap["Company"]
        rows = []
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if row and row[ci]:
                rows.append(row)
        print(f"\n=== {mid} FINAL n={len(rows)} ===")
        for col in KEY_COLS:
            if col not in hmap:
                print(f"  {col}: MISSING COL")
                continue
            idx = hmap[col]
            blank = []
            for row in rows:
                v = row[idx]
                s = "" if v is None else str(v).strip()
                if not s or s.lower() in ("not publicly disclosed", "n/a", "none", "-"):
                    # treat NPD as filled for employees/turnover
                    if s.lower() == "not publicly disclosed" and col in ("Employees", "Turnover / Scale"):
                        continue
                    blank.append(str(row[ci])[:40])
            print(f"  {col}: blank={len(blank)}")
            if blank and col in (
                "Contact Person",
                "Phone / WhatsApp",
                "Core Categories",
                "Specialty Focus",
                "Price Segment",
                "Key Brands Represented",
                "Founded",
                "Headquarters",
                "Website",
            ):
                for b in blank[:8]:
                    print(f"    - {b}")
        # show newly merged if present
        for row in rows:
            name = str(row[ci])
            if any(
                x in name.lower()
                for x in (
                    "arrow",
                    "galaxtecc",
                    "suntsu",
                    "bluelotus",
                    "green chemistry",
                    "integrity biochem",
                    "riddhi",
                    "trans western",
                    "unitedchem",
                    "gchem",
                    "fortis",
                    "invitty",
                    "wijungle",
                    "digitaltech",
                    "importgm",
                    "indiannetwork",
                    "resolution",
                    "servershyderabad",
                )
            ):
                print(
                    f"  NEW {name[:35]:35} | "
                    f"founded={row[hmap.get('Founded')]} | "
                    f"brands={str(row[hmap.get('Key Brands Represented')])[:50]} | "
                    f"price={row[hmap.get('Price Segment')]} | "
                    f"contact={row[hmap.get('Contact Person')]}"
                )
        wb.close()


if __name__ == "__main__":
    main()
