"""Audit price/specialty/brands distributions (PowerShell-safe)."""
from __future__ import annotations

import re
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")


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
        print("\n", mid)
        pi = hmap.get("Price Segment")
        ci = hmap.get("Company")
        bi = hmap.get("Key Brands Represented")
        si = hmap.get("Specialty Focus")
        prices: Counter[str] = Counter()
        brand_problems = []
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row or ci is None or not row[ci]:
                continue
            prices[str(row[pi] or "").strip() or "(blank)"] += 1
            kb = str(row[bi] or "").strip()
            low = kb.lower()
            looks_bad = False
            if not kb:
                looks_bad = True
            elif any(
                x in low
                for x in [
                    "professional source",
                    "private investors",
                    "specialty chemicals",
                    "bio-based",
                    "smartwatches",
                    "wearable technology",
                    "firewalls",
                    "networking",
                    "components;",
                    "systems for",
                ]
            ):
                looks_bad = True
            elif ";" not in kb and "," not in kb and len(kb.split()) > 6:
                looks_bad = True
            if looks_bad:
                brand_problems.append((str(row[ci])[:40], kb[:110] or "(blank)"))
        print(" Price Segment:", dict(prices))
        print(" Bad/blank Key Brands:", len(brand_problems))
        for a, b in brand_problems[:30]:
            print("  ", a, "|", b)
        wb.close()


if __name__ == "__main__":
    main()
