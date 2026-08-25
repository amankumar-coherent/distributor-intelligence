# -*- coding: utf-8 -*-
"""One-shot apply of all agent-verified HQ seeds. No DeepSeek race."""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
SCRIPT = Path(r"D:\vendor-intel\scripts\verify_fill_hq_flexpack.py")

# Extra seeds verified in this session (beyond MANUAL_SEEDS in script)
EXTRA = {
    "Crown Packaging": "Chesterfield, Missouri, United States",
    "Dollies and Boxes Unlimited": "Scarborough, Ontario, Canada",
    "All Packaging Company": "Aurora, Colorado, United States",
    "Fillmore Container": "Lancaster, Pennsylvania, United States",
    "EcoPack": "Dandenong South, Victoria, Australia",
    "Global Industrial": "Port Washington, New York, United States",
    "Imperial Dade": "Jersey City, New Jersey, United States",
    "McMaster-Carr": "Elmhurst, Illinois, United States",
    "Manutan": "Gonesse, France",
    "MonotaRO": "Osaka, Japan",
    "Hisco": "Houston, Texas, United States",
    "Kaufman Container": "Cleveland, Ohio, United States",
    "O.Berk": "Union, New Jersey, United States",
    "HD Supply": "Atlanta, Georgia, United States",
    "Freund Container & Supply": "Chicago, Illinois, United States",
    "Midland Paper": "Wheeling, Illinois, United States",
    "Moglix": "Noida, Uttar Pradesh, India",
    "Igepa Group": "Hamburg, Germany",
    "GWP Group": "Cricklade, Wiltshire, United Kingdom",
    "JPG": "Survilliers, France",
    "Karton.eu": "Spremberg, Germany",
    "Network Packaging": "Wolverhampton, United Kingdom",
    "Novvia Group": "Franklin, Tennessee, United States",
    "Packaging Unlimited": "Louisville, Kentucky, United States",
    "Hillas Packaging": "Fort Worth, Texas, United States",
    "Mondoffice": "Survilliers, France",
    "Gerdmans": "Markaryd, Sweden",
    "IPS Packaging": "Greenville, South Carolina, United States",
    "Industrial Packaging Supplies": "Webster, Massachusetts, United States",
    "MAC Papers": "Jacksonville, Florida, United States",
    "MaxPak": "Huntingdale, Victoria, Australia",
    "Staples": "Framingham, Massachusetts, United States",
    "Quill": "Lincolnshire, Illinois, United States",
    "Zoro Tools": "Buffalo Grove, Illinois, United States",
    "WebstaurantStore": "Lititz, Pennsylvania, United States",
    "Richards Packaging": "Mississauga, Ontario, Canada",
    "Shorr Packaging": "Aurora, Illinois, United States",
    "SKS Bottle & Packaging": "Saratoga Springs, New York, United States",
    "Winc": "Mascot, New South Wales, Australia",
    "Ratioform": "Pliening, Germany",
    "Papyrus": "Molndal, Sweden",
    "Viking Direct": "Dublin, Ireland",
    "Spicers": "Dandenong South, Victoria, Australia",
    "U.S. Plastic Corp.": "Lima, Ohio, United States",
    "The Cary Company": "Addison, Illinois, United States",
    "Aramsco": "Radnor, Pennsylvania, United States",
    "Elkay Plastics": "Commerce, California, United States",
    "Trinity Packaging Supply": "Voorhees, New Jersey, United States",
    "PackagingSupplies.com": "Solon, Ohio, United States",
    "Nashville Wraps": "Hendersonville, Tennessee, United States",
    "The Packaging Company": "Brampton, Ontario, Canada",
    "Paper Mart": "Orange, California, United States",
    "Diversified Packaging Solutions": "St. Louis, Missouri, United States",
    "Delaware Valley Packaging": "Fairfield, California, United States",
}

text = SCRIPT.read_text(encoding="utf-8")
block = text.split("MANUAL_SEEDS", 1)[1].split("COUNTRY_ONLY", 1)[0]
seeds = dict(re.findall(r'"([^"]+)":\s*"([^"]+)"', block))
seeds.update(EXTRA)

COUNTRY_ONLY = {
    "united states", "usa", "us", "united kingdom", "uk", "great britain", "england",
    "germany", "deutschland", "australia", "canada", "france", "india", "mexico",
    "spain", "ireland", "netherlands", "sweden", "japan", "italy", "singapore",
    "new zealand", "south africa", "switzerland", "austria", "belgium", "china",
    "brazil", "poland", "portugal", "denmark", "norway", "finland",
}


def needs(hq: str) -> bool:
    h = (hq or "").strip()
    if not h or "not publicly" in h.lower():
        return True
    if "," in h:
        left = h.split(",")[0].strip()
        if left and left.lower() not in COUNTRY_ONLY and len(left) >= 3:
            return False
    if h.lower() in COUNTRY_ONLY:
        return True
    return "," not in h and len(h.split()) <= 3


wb = load_workbook(FP)
ws = wb.active
applied = missing = 0
for row in ws.iter_rows(min_row=2):
    name = str(row[0].value or "").strip()
    if name in seeds:
        row[3].value = seeds[name]
        applied += 1
    elif name:
        # track if still weak
        pass

weak = good = 0
still = []
for row in ws.iter_rows(min_row=2, values_only=True):
    name = str(row[0] or "").strip()
    hq = str(row[3] or "").strip()
    if needs(hq):
        weak += 1
        still.append(f"{name}\t{row[1]}\t{hq}")
    else:
        good += 1

wb.save(FP)
out = FP.parent / "hq_still_weak.txt"
out.write_text("\n".join(still), encoding="utf-8")
print(f"seed_keys={len(seeds)} applied_rows={applied}")
print(f"cityish={good} weak={weak} total={good+weak}")
print(f"wrote {out}")
