#!/usr/bin/env python3
"""Clean Core Categories / Specialty Focus / Key Brands — real data only."""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

XLSX = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global"
    r"\global_avocado_oil_market_OFFICE_SLIM.xlsx"
)
NPD = "Not publicly disclosed"

# Placeholder / invented brand language → clear to NPD
PLACEHOLDER = re.compile(
    r"(?i)("
    r"principal .+ brand|"
    r"principal specialty|"
    r"principal fmcg|"
    r"principal edible|"
    r"principal supplier|"
    r"principal food|"
    r"private label & brand|"
    r"private label & regional|"
    r"private[- ]label .+ oil|"
    r"member banner|"
    r"banner private label|"
    r"contract-approved|"
    r"traded principal|"
    r"group ingredient|"
    r"house/?private-label|"
    r"house natural oil|"
    r"house essential|"
    r"specialty avocado .+ brand|"
    r"specialty & import|"
    r"warehouse private label|"
    r"cash & carry private label|"
    r"retailer private label|"
    r"national oil|"
    r"national avocado|"
    r"brand oils?, others|"
    r"brand edible oils|"
    r"brand avocado|"
    r"regional & private-label|"
    r"imported specialty oil brands|"
    r"produce house brands|"
    r"grower avocado|"
    r"organic grower brands|"
    r"cooperative produce brands|"
    r"cosmetic ingredient / oil|"
    r"portfolio principals|"
    r"supplier brands \(oils|"
    r"azelis principal|"
    r"barentz portfolio|"
    r"bidfood private label|"
    r"brakes private label|"
    r"trading-house food|"
    r"chain specs|"
    r"approved oil brands|"
    r"program staple brands|"
    r"dairy brand oils|"
    r"natural ingredient lines|"
    r"bulk private label|"
    r"sourced oil grades|"
    r"supplies brands|"
    r"limited oil brands|"
    r"drinks brands|"
    r"culinary oil skus|"
    r"premium avocado & specialty|"
    r"natural/organic avocado oil brands|"
    r"specialty oil & ingredient suppliers|"
    r"acme-hardesty specialty oleochemical lines"  # not a brand name
    r")"
)

# Web-verified Key Brands only (sources noted in comments)
BRAND_OVERRIDES = {
    # all-organic-treasures.com
    "All Organic Treasures GmbH": "Heliaflor, AOT Food Ingredients, AOT Cosmetic Ingredients, others",
    "AOT GmbH": "Heliaflor, AOT Food Ingredients, AOT Cosmetic Ingredients, others",
    # biona.co.uk
    "Biona Organic Foods": "Biona, others",
    # cho-america.com / bulkbycho.com
    "Bulk by CHO": "Terra Delyssa, Origin 846, Fork & Leaf, Bella Del Sol, MORESH, others",
    # unfi.com private brands / SuperValu heritage
    "SuperValu": "Essential Everyday, Culinary Circle, Wild Harvest, Equaline, Arctic Shores, others",
    "Supervalu Wholesale": "Essential Everyday, Culinary Circle, Wild Harvest, Equaline, others",
    "UNFI (United Natural Foods)": "Essential Everyday, Culinary Circle, Wild Harvest, Equaline, Field Day, Woodstock, others",
    # cobana.com / FreshPlaza
    "Cobana Fruchtring": "Cobana, Ory, others",
    "Cobana GmbH & Co. KG": "Cobana, Ory, others",
    # markon.com
    "Markon Cooperative": "Markon First Crop, Ready-Set-Serve, Markon Essentials, others",
    # bidfood.co.uk
    "Bidfood": "Everyday Favourites, others",
    "Bidvest Australia": "Not publicly disclosed",  # AU Bidfood brands not verified here
    # costco Kirkland Signature avocado oil
    "Costco": "Kirkland Signature, others",
    "Sam's Club Mexico": "Member's Mark, others",
    # sysco foodie brand tiers
    "Sysco": "Sysco Classic, Sysco Reliance, Sysco Premium, Fry-On, others",
    "Sysco Specialty": "Sysco Classic, Sysco Premium, others",
    # dipasa.nl
    "Dipasa Europe": "Dipasa, others",
    "Dipasa Europe B.V.": "Dipasa, others",
    # fyffes
    "Fyffes": "Fyffes, others",
    # company is the brand
    "Gustav Heess": "Gustav Heess, others",
    "Henry Lamotte Oils": "Henry Lamotte Oils, others",
    "Jedwards International": "Jedwards International, others",
    "Nature's Best": "Not publicly disclosed",
}

# Core / Specialty Focus corrections (web-aligned, no country names, end with , others)
FOCUS_CORE_OVERRIDES = {
    "Bulk by CHO": {
        "core": "Bulk edible oils; olive oil; avocado oil; specialty oils, others",
        "focus": "Bulk & private-label edible oil supply for food manufacturing and foodservice, others",
    },
    "Bulk Apothecary": {
        "core": "Natural ingredients wholesale; carrier oils; butters, others",
        "focus": "Cosmetic and personal-care botanical oils and raw materials wholesale, others",
    },
    "Foodservice Direct": {
        "core": "Online foodservice wholesale; edible oils, others",
        "focus": "Foodservice culinary oils and pantry staples distribution, others",
    },
    "foodguys": {
        "core": "Foodservice redistribution; specialty foods; edible oils, others",
        "focus": "Specialty foodservice ingredients and culinary oils redistribution, others",
    },
    "Spinneys Distribution": {
        "core": "Grocery retail distribution; edible oils, others",
        "focus": "Grocery retail edible oils and packaged foods assortment, others",
    },
    "Naissance": {
        "core": "Natural oils wholesale; carrier oils; avocado oil, others",
        "focus": "Natural and organic carrier oils for personal care and wellness, others",
    },
    "All Organic Treasures GmbH": {
        "core": "Organic vegetable oils; plant proteins; cosmetic raw materials, others",
        "focus": "Organic food and natural cosmetic ingredient oils and proteins, others",
    },
    "AOT GmbH": {
        "core": "Organic vegetable oils; plant proteins; cosmetic raw materials, others",
        "focus": "Organic food and natural cosmetic ingredient oils and proteins, others",
    },
    "Biona Organic Foods": {
        "core": "Organic grocery foods; edible oils; avocado oil, others",
        "focus": "Organic retail edible oils including avocado, olive, coconut and seed oils, others",
    },
    "Charkit Chemical": {
        "core": "Specialty chemicals; fine chemicals distribution, others",
        "focus": "Specialty chemical ingredients distribution for personal care and industrial uses, others",
    },
    "Charkit Chemical Company LLC": {
        "core": "Specialty chemicals; fine chemicals distribution, others",
        "focus": "Specialty chemical ingredients distribution for personal care and industrial uses, others",
    },
    "Aromantic": {
        "core": "Cosmetic raw materials; carrier oils, others",
        "focus": "Natural cosmetic raw materials and botanical carrier oils, others",
    },
    "Base Formula": {
        "core": "Cosmetic ingredients; carrier oils; avocado oil, others",
        "focus": "Natural cosmetic and DIY formulation oils and ingredients, others",
    },
    "Camden-Grey Essential Oils": {
        "core": "Essential oils; carrier oils; avocado oil, others",
        "focus": "Bulk essential and carrier oils for personal care and aromatherapy, others",
    },
}


def is_placeholder_brands(val: str) -> bool:
    s = (val or "").strip()
    if not s:
        return True
    if s.lower() in {"n/a", "na", "none", "not publicly disclosed"}:
        return False
    if PLACEHOLDER.search(s):
        return True
    # no proper brand-like token (Capitalized word that isn't Private/Principal/Brand/House)
    tokens = re.findall(r"[A-Z][A-Za-z0-9&''\-]{1,}", s)
    junk = {
        "Private",
        "Principal",
        "Brand",
        "Brands",
        "House",
        "Member",
        "Retailer",
        "National",
        "Specialty",
        "Imported",
        "Organic",
        "Natural",
        "Group",
        "Trading",
        "Contract",
        "Approved",
        "Warehouse",
        "Cash",
        "Carry",
        "Banner",
        "Others",
        "Oil",
        "Oils",
        "Food",
        "Foods",
        "Ingredient",
        "Ingredients",
        "Supplier",
        "Suppliers",
        "Portfolio",
        "Lines",
        "Line",
        "Label",
        "Labels",
        "Regional",
        "Premium",
        "Bulk",
        "Sourced",
        "Grades",
        "Program",
        "Staple",
        "Dairy",
        "Drinks",
        "Culinary",
        "Skus",
        "SKU",
        "Limited",
        "Supplies",
        "Grower",
        "Produce",
        "Cooperative",
        "Cosmetic",
        "Personal",
        "Care",
        "Botanical",
        "Carrier",
        "Essential",
        "Avocado",
        "Edible",
        "FMCG",
        "And",
        "The",
        "For",
        "To",
        "Of",
        "With",
        "Chain",
        "Specs",
        "Independent",
    }
    real = [t for t in tokens if t not in junk]
    return len(real) == 0


def ensure_others(text: str) -> str:
    t = (text or "").strip().rstrip(",")
    if not t:
        return t
    if re.search(r",\s*others\s*$", t, re.I):
        return re.sub(r",\s*others\s*$", ", others", t, flags=re.I)
    return t + ", others"


def main() -> None:
    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}

    cleared = fixed_brands = fixed_focus = 0
    for row in range(3, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        brands = str(ws.cell(row, ic["Key Brands Represented"]).value or "").strip()
        core = str(ws.cell(row, ic["Core Categories"]).value or "").strip()
        focus = str(ws.cell(row, ic["Specialty Focus"]).value or "").strip()

        # 1) brands
        if co in BRAND_OVERRIDES:
            ws.cell(row, ic["Key Brands Represented"]).value = BRAND_OVERRIDES[co]
            fixed_brands += 1
        elif is_placeholder_brands(brands):
            ws.cell(row, ic["Key Brands Represented"]).value = NPD
            cleared += 1
        else:
            # keep real brands; ensure , others
            if brands.lower() != NPD.lower():
                ws.cell(row, ic["Key Brands Represented"]).value = ensure_others(brands)

        # 2) core/focus overrides
        if co in FOCUS_CORE_OVERRIDES:
            hit = FOCUS_CORE_OVERRIDES[co]
            if hit.get("core"):
                ws.cell(row, ic["Core Categories"]).value = hit["core"]
            if hit.get("focus"):
                ws.cell(row, ic["Specialty Focus"]).value = hit["focus"]
            fixed_focus += 1
        else:
            # ensure specialty focus ends with , others; strip accidental country if any (none expected)
            if focus and focus.lower() != NPD.lower():
                ws.cell(row, ic["Specialty Focus"]).value = ensure_others(focus)
            if core and core.lower() != NPD.lower():
                ws.cell(row, ic["Core Categories"]).value = ensure_others(core)

    wb.save(XLSX)
    wb.close()

    # recount
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c or "") for c in rows[1]]
    ic2 = {h: i for i, h in enumerate(hdr) if h}
    npd_b = sum(
        1
        for r in rows[2:]
        if str(r[ic2["Key Brands Represented"]] or "").strip() == NPD
    )
    real_b = len(rows) - 2 - npd_b
    print(
        f"cleared_placeholder_brands={cleared} verified_brand_overrides={fixed_brands} "
        f"core_focus_fixes={fixed_focus}"
    )
    print(f"Key Brands real={real_b} NPD={npd_b} total={len(rows)-2}")
    wb.close()


if __name__ == "__main__":
    main()
