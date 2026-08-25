#!/usr/bin/env python3
"""Fill remaining 45 avocado Key Brands + gap columns from website evidence."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

X = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_OFFICE_SLIM.xlsx"
)
AUDIT = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\_avocado_brands_remaining45_audit.json"
)
NPD = "Not publicly disclosed"

# company -> updates (only set keys provided)
UPDATES: dict[str, dict] = {
    "Adams Australia": {
        "Key Brands Represented": "Adams Australia, Adams Vegetable Oils, others",
    },
    "American International Foods": {
        "Key Brands Represented": "AIFI, others",
    },
    "Aramark Middle East Distribution": {
        "Key Brands Represented": "MEDI, Burger Studio, Bleecker Street, Topio's, P.O.D., Zoca, Grab & Gather, others",
    },
    "BOETTGER Food Ingredients": {
        "Key Brands Represented": "Boettger, others",
    },
    "Cargill Food Distribution Europe": {
        "Key Brands Represented": "Clear Valley, Odyssey, others",
    },
    "Coastal Pacific Food Distributors": {
        "Key Brands Represented": "CPFD, others",
    },
    "Connect The World S.A. de C.V.": {
        "Key Brands Represented": NPD,
    },
    "Core-Mark": {
        "Key Brands Represented": "Mighty Marked, Fresh Take, others",
    },
    "Dhamecha Cash & Carry": {
        "Key Brands Represented": "DMART, Day 1, Green Valley, others",
    },
    "DKSH Indonesia": {
        "Key Brands Represented": NPD,
    },
    "Dot Foods": {
        "Key Brands Represented": NPD,
    },
    "Europagricultproduct": {
        "Key Brands Represented": NPD,
    },
    "Food Supplies Company": {
        "Website": "https://www.foosco.com",
        "Key Brands Represented": "Megachips, XO, Krunchex, others",
        "Office No.": "+973 17 166166",
        "Country Code": "BH",
        "Region Code": "ME",
        "Summary": "FOOSCO (The Food Supply Company Ltd.) — Bahrain food manufacturing and catering; snack/juice brands Megachips, XO, Krunchex.",
    },
    "Foodservice Direct": {
        "Key Brands Represented": NPD,
    },
    "Fresh Produce Consortium Members Distribution": {
        "Website": "https://www.freshproduce.org.uk/",
        "Key Brands Represented": NPD,
        "Summary": "UK Fresh Produce Consortium member distribution channel — trade association / member wholesalers; no single house brand set disclosed.",
    },
    "Gillco Ingredients": {
        "Website": "https://gillco.com/",
        "Key Brands Represented": "Cargill, Corbion, FrieslandCampina, IFF, Cosucra, Bio-Botanica, Syensqo, others",
        "Office No.": NPD,
        "Ownership": "Privately held (Azelis company)",
        "Summary": "Gillco Ingredients (Azelis) — US industrial food/beverage ingredient distributor; oils & shortenings via manufacturing partners including Cargill.",
    },
    "Gourmet Gulf Distribution": {
        "Website": "https://www.gourmetgulf.com/",
        "Key Brands Represented": "YO! Sushi, California Pizza Kitchen, Texas de Brazil, Panda Express, Hummingbird Bakery, Azkadenya, Dalloyau, others",
        "Office No.": "+971-4-2910500",
        "Country Code": "AE",
        "Region Code": "ME",
        "Summary": "Gourmet Gulf — Dubai F&B franchise operator/distributor for international dining brands across UAE, Qatar and Saudi Arabia.",
    },
    "GreenGold USA": {
        "Key Brands Represented": NPD,
    },
    "Henry Broch Foods": {
        "Key Brands Represented": "Henry Broch, Jardins and Broch, others",
        "Summary": "Henry Broch Foods (Jardins and Broch) — ambient dry-blend co-manufacturer and industrial ingredient sales; private-label manufacturing partner.",
    },
    "Horeca Trade": {
        "Key Brands Represented": NPD,
    },
    "Horizon Harvest": {
        "Website": "https://horizonharvests.com",
        "Key Brands Represented": "Horizon Harvest, others",
    },
    "Infinity Ingredients": {
        "Website": "https://infinity-ingredients.co.uk/",
        "Key Brands Represented": "Infinity Ingredients Elements, Provital, Sophim, ROELMI HPC, Laboratoires Expanscience, Evident Ingredients, others",
        "Office No.": "01344 354 900",
        "Summary": "Infinity Ingredients — UK/Ireland cosmetic & personal-care ingredient distributor; own Elements oils/butters range plus global principal brands.",
    },
    "Ingredi": {
        "Key Brands Represented": "Ingredi, others",
    },
    "ITOCHU Food Distribution": {
        "Website": "https://www.itochu-shokuhin.com/",
        "Key Brands Represented": "Karada Switch, TOMIN ICHIBA, others",
        "Summary": "ITOCHU-SHOKUHIN — Japan liquor/food wholesaler; health PB Karada Switch and frozen brand TOMIN ICHIBA plus collab ready meals.",
    },
    "Jan Oskam": {
        "Website": "https://www.janoskam.nl/",
        "Key Brands Represented": "Oskam, MAX Freshness, others",
        "Office No.": "+31 30 232 0420",
        "Office Email": "info@janoskam.nl",
        "Country Code": "NL",
        "Region Code": "EU",
        "Summary": "Jan Oskam B.V. — Dutch hard-fruit grower/exporter (apples/pears) and fresh produce trader; MAX Freshness supplier brand.",
    },
    "Labatt Food Service": {
        "Key Brands Represented": NPD,
    },
    "Maines Paper & Food Service": {
        "Key Brands Represented": NPD,
    },
    "Martin-Brower": {
        "Key Brands Represented": NPD,
    },
    "Matthew Clark": {
        "Key Brands Represented": "Matthew Clark Bibendum, Generation Wine, others",
        "Summary": "Matthew Clark (C&C Group) — UK on-trade drinks wholesaler; exclusive wine selections including Generation Wine and broad third-party portfolio.",
    },
    "Medikonda Australia": {
        "Key Brands Represented": NPD,
    },
    "Mitsubishi Corporation Food Distribution": {
        "Website": "https://www.mitsubishi-shokuhin.com/",
        "Key Brands Represented": "Karada Shift, Motto Nippon, Lilly, Lilly Casual, others",
        "Summary": "Mitsubishi Shokuhin — Japan food wholesaler with house brands Karada Shift, Motto Nippon, Lilly / Lilly Casual plus distributed global brands.",
    },
    "Nutra Food Ingredients": {
        "Website": "https://nutrafoodingredients.com/",
        "Key Brands Represented": "Grand Valley Organics, others",
        "Summary": "Nutra Food Ingredients LLC — Kentwood MI nutritional/functional ingredient supplier; Grand Valley Organics brand; MCT/coconut oil powders.",
    },
    "Olam Food Ingredients Distribution": {
        "Website": "https://www.ofi.com/",
        "Key Brands Represented": "deZaan, Unicao, Joanes, Huysman, Britannia Specialty Fats, Macao, others",
        "Summary": "ofi (Olam Food Ingredients) — global cocoa/coffee/nuts/dairy/spices ingredients; cocoa brands deZaan, Unicao, Joanes, Huysman; Britannia Specialty Fats.",
    },
    "OPW Ingredients": {
        "Website": "https://opw-ingredients.com/",
        "Key Brands Represented": "OPW Ingredients, others",
        "Summary": "OPW Ingredients GmbH — plant-based oils/proteins/flours manufacturer-trader (incl. avocado oil) for food, feed and personal care.",
    },
    "Patties Food Group Distribution": {
        "Website": "https://pattiesfoodgroup.com/",
        "Key Brands Represented": "Four'n Twenty, Patties, Herbert Adams, Nanna's, Chefs Pride, Boscastle, Snowy River, others",
        "Summary": "Patties Food Group — Australia/NZ convenience & snacking food manufacturer; iconic pastry/dessert brands Four'n Twenty, Patties, Herbert Adams, Nanna's.",
    },
    "Pratts Wholesale": {
        "Website": "https://pratts.com/",
        "Key Brands Represented": "Vincenzo's, Pratts, others",
        "Office No.": "+1 204-949-2800",
        "Summary": "Pratts Wholesale Food Service — Winnipeg grocery/foodservice wholesaler; Vincenzo's private-label cheese line among offerings.",
    },
    "Pro*Act": {
        "Key Brands Represented": NPD,
    },
    "Reyes Holdings": {
        "Key Brands Represented": NPD,
    },
    "RFI Ingredients": {
        "Key Brands Represented": "OxyPhyte, ChromaPhyte, Chocamine, GS15-4, VitaKelp, FermaPro, Promulsion, others",
        "Summary": "RFI Ingredients — Blauvelt NY branded botanical/functional ingredients (OxyPhyte, ChromaPhyte, Chocamine, GS15-4, VitaKelp, FermaPro, Promulsion).",
    },
    "Scavuzzo's": {
        "Website": "https://scavuzzos.com/",
        "Key Brands Represented": "S&S Meat Co., 4G Meat Processing, others",
        "Office No.": NPD,
        "Summary": "Scavuzzo's Foodservice — Kansas City independent broadline distributor with proprietary meat plants S&S Meat Co. and 4G Meat Processing.",
    },
    "Sojitz": {
        "Key Brands Represented": NPD,
    },
    "Swire Foods Distribution": {
        "Key Brands Represented": "Coca-Cola, Taikoo Sugar, others",
        "Summary": "Swire Foods / Swire Coca-Cola — manufactures/distributes Coca-Cola system beverages and Taikoo Sugar food products across Greater China and Asia.",
    },
    "Vistar": {
        "Key Brands Represented": NPD,
    },
    "Quimidroga": {
        "Key Brands Represented": "Quimidroga, BiomeNrich, Aquacid, others",
        "Ownership": "Privately held",
        "Employees": "201-500",
        "Country Code": "ES",
        "Region Code": "EU",
        "Summary": "Quimidroga — Barcelona chemical & food-ingredient distributor since 1944; represents specialty lines including BiomeNrich probiotics and Aquacid water-treatment chemicals.",
    },
    "Circle Ingredients": {
        "Key Brands Represented": "Circle Ingredients, Circle Packaging, Circle Trading, others",
        "Ownership": "Privately held",
        "Employees": "11-50",
        "Office No.": "+31 23 303 8502",
        "Country Code": "NL",
        "Region Code": "EU",
        "Summary": "Circle Ingredients B.V. (Heemstede) — B2B chemicals, food ingredients, packaging and machinery sourcing/distribution across NL, Caribbean and Colombia.",
    },
}

wb = openpyxl.load_workbook(X)
ws = wb.active
hdr = [c.value for c in ws[2]]
ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}

applied_brands = 0
applied_fields = 0
missing = []
for r in range(3, ws.max_row + 1):
    co = str(ws.cell(r, ic["Company"]).value or "").strip()
    if co not in UPDATES:
        continue
    for col, val in UPDATES[co].items():
        if col not in ic:
            missing.append(f"{co}:{col}")
            continue
        ws.cell(r, ic[col]).value = val
        applied_fields += 1
        if col == "Key Brands Represented":
            applied_brands += 1

wb.save(X)
wb.close()

# recount
wb = openpyxl.load_workbook(X, data_only=True)
ws = wb.active
hdr = [c.value for c in ws[2]]
ic = {str(h): i for i, h in enumerate(hdr) if h}
real = npd = 0
still = []
for row in ws.iter_rows(min_row=3, values_only=True):
    c = str(row[ic["Company"]] or "").strip()
    if not c:
        continue
    b = str(row[ic["Key Brands Represented"]] or "").strip()
    if b in ("", NPD):
        npd += 1
        still.append(c)
    else:
        real += 1
wb.close()

audit = {
    "ts": datetime.now(timezone.utc).isoformat(),
    "applied_brand_rows": applied_brands,
    "applied_field_writes": applied_fields,
    "total": real + npd,
    "real": real,
    "npd": npd,
    "still_npd": still,
    "missing_cols": missing,
    "note": "Website/agent verified; no invented phones/emails; honest NPD for redistributors without PL",
}
AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
print(json.dumps({k: audit[k] for k in ("applied_brand_rows", "applied_field_writes", "total", "real", "npd")}, indent=2))
print("still_npd:", "|".join(still))
