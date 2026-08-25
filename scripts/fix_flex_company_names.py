"""Normalize company names to proper trade/legal names (no artificial Dist/geo tags)."""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = PATH.parent / "fix_company_names_audit.json"

# Explicit proper-name overrides (old -> proper)
RENAMES: dict[str, str] = {
    "Abco Kovex UK": "Abco Kovex",
    "All American Containers Dist": "All American Containers",
    "All Packaging Company Denver": "All Packaging Company",
    "Askul Japan": "Askul",
    "Atlantic Packaging Corp Dist": "Atlantic Packaging Corp",
    "Bags Plus Australia": "Bags Plus",
    "Boxmart UK": "Boxmart",
    "Boxworks Australia": "Boxworks",
    "CPS Supply Line (Crawford Provincial)": "CPS Supply Line",
    "Emballage Direct France": "Emballage Direct",
    "Flexi-Films UK": "Flexi-Films",
    "FlexPRO Europe": "FlexPRO",
    "GM Packaging UK": "GM Packaging",
    "Hughes Enterprises (United Hughes / BradyPLUS)": "Hughes Enterprises",
    "Individual FoodService (BradyPLUS)": "Individual FoodService",
    "Kalamazoo Spain": "Kalamazoo",
    "MaxPak Australia": "MaxPak",
    "MonotaRO Japan": "MonotaRO",
    "Multiplastics Europe": "Multiplastics",
    "Online Packaging Australia": "Online Packaging",
    "Packaging Distributors South Africa": "Packaging Distributors",
    "Packaging Express UK": "Packaging Express",
    "Packaging Products NZ": "Packaging Products",
    "Packaging Supplies UK": "Pack Supplies",  # pack-supplies.co.uk brand
    "Packaging World UK": "Packaging World",
    "Pakonline Australia": "Pakonline",
    "Pluspak Australia": "Pluspak",
    "Poly Pak Canada": "Poly Pak",
    "Polythene UK": "Polythene UK",  # official brand includes UK
    "Premier Packaging NZ": "Premier Packaging",
    "Seow Khim Polythelene (SKP)": "Seow Khim Polythene",
    "Spicers Australia": "Spicers",
    "Spicers Paper Canada": "Spicers Paper",
    "Transpack UK": "Transpack",
    "UPack India": "UPack",
    "Waltons ZA": "Waltons",
    "Winc Australia": "Winc",
    "Global Trading UK": "Global Trading",
    "Corrugated Packaging Inc Dist": "Corrugated Packaging Inc",
    "Packaging Solutions France": "Packaging Solutions",
    "Imballaggi Italia Dist": "Imballaggi Italia",
    "Packaging NZ": "Packaging NZ",
    "NZ Packaging": "NZ Packaging",
    "Packaging Direct NZ": "Packaging Direct New Zealand",
    "Packaging Direct BV": "Packaging Direct B.V.",
    "Boxes UK": "Boxes UK",
    "The Packaging Company Canada": "The Packaging Company Canada",
    "Global Industrial Canada": "Global Industrial Canada",
    "QPack UK Merchant": "QPack",
    "Horna Packaging DE": "Horna",
    "New Century Packaging USA": "New Century Packaging USA Inc",
    "UK Packaging Online": "UK Packaging",
    "Office Packaging UK": "Office Packaging",
    "Packaging Products Corp Dist": "Packaging Products Corp",
    "NNZ Germany": "NNZ GmbH",
    "Quill.com Packaging": "Quill",
    "IndustryBuying Packaging": "IndustryBuying",
    "Moglix Packaging": "Moglix",
    "All American Packaging Supply": "All American Packaging",
    "Empaques y Envases de M�xico": "Empaques y Envases de Mexico",
    "Empaques del Pac�fico": "Empaques del Pacifico",
    "TransPack Krumbach": "TransPack Krumbach",
    "Packaging24 GmbH": "Packaging24 GmbH",
    "Packaging Service GmbH": "Packaging Service GmbH",
    "Verpackungszentrale GmbH": "Verpackungszentrale GmbH",
    "Papier-Mettler KG": "Papier-Mettler",
    "Verpackungsprofi GmbH": "Verpackungsprofi",
    "Verpackungsbedarf GmbH": "Verpackungsbedarf",
    "ESWE Versandpack": "ESWE Versandpack",
    "Paul Hildebrandt": "Paul Hildebrandt",
    "Knueppel Verpackung": "Knueppel Verpackung",
    "Baas Verpakkingen": "Baas Verpakkingen",
    "Tillmann Verpackungen": "Tillmann Verpackungen",
    "Wagener Verpackung": "Wagener Verpackung",
    "Brangs Heinrich": "Brangs + Heinrich",
    "Ehlken Verpackungen": "Ehlken Verpackungen",
    "Packpack.de": "Packpack",
    "Karton.eu": "Karton.eu",
    "PackagingSupplies.com": "PackagingSupplies.com",
    "MrBoxOnline": "Mr Box Online",
    "R.J. Schinner Co": "R.J. Schinner Co.",
    "Schermerhorn Bros.": "Schermerhorn Bros.",
    "U.S. Packaging & Wrapping": "U.S. Packaging & Wrapping",
    "U.S. Plastic Corp.": "U.S. Plastic Corp.",
    "O.Berk": "O.Berk",
    "WB Mason": "W.B. Mason",
    "The Home Depot Pro": "The Home Depot",
    "Bidvest Paperplus": "Bidvest Paperplus",
    "MAC Papers and Packaging": "MAC Papers",
    "Polyair Inter Pack": "Polyair",
    "Sigma Supply of North America": "Sigma Supply",
    "Southern Packaging LP": "Southern Packaging",
    "Tantraco Enterprise": "Tantraco",
    "Robert McCabe Packaging": "Robert McCabe Packaging",
    "Industrial & Shippers Supply": "Industrial & Shippers Supply",
    "Package Design & Supply": "Package Design & Supply",
    "Box and Packaging Supply": "Box and Packaging Supply",
    "Container and Packaging Supply": "Container and Packaging Supply",
    "Acme Paper & Supply": "Acme Paper & Supply",
    "Landmark Paper & Packaging": "Landmark Paper & Packaging",
    "Preferred Plastics & Packaging": "Preferred Plastics & Packaging",
    "Burch Bottle & Packaging": "Burch Bottle & Packaging",
    "SKS Bottle & Packaging": "SKS Bottle & Packaging",
    "Dayton Bag & Burlap": "Dayton Bag & Burlap",
    "Imperial Bag & Paper": "Imperial Bag & Paper",
    "Braden Packaging & Supply": "Braden Packaging & Supply",
    "Huckster Packaging & Supply": "Huckster Packaging & Supply",
    "Service Box & Tape": "Service Box & Tape",
    "Freund Container & Supply": "Freund Container & Supply",
    "R+R Packaging": "R+R Packaging",
    "BB-Verpackungen": "BB-Verpackungen",
    "BVS Verpacken": "BVS Verpacken",
    "Hoffmann Verpackung": "Hoffmann Verpackung",
    "Verpacken Online": "Verpacken Online",
    "Spar-Pack": "Spar-Pack",
    "PackQueen": "PackQueen",
    "Surepak": "Surepak",
    "PacPlus": "PacPlus",
    "Packen24": "Packen24",
    "Serwo": "Serwo",
    "ANW Packaging": "ANW Packaging",
    "Maidstone Packaging Supplies": "Maidstone Packaging Supplies",
    "Sovereign Partners": "Sovereign Partners",
    "De Wild Group": "De Wild",
    "Furodet Emballages": "Furodet Emballages",
    "Embageant": "Embageant",
    "Emballage Discount": "Emballage Discount",
    "Emballages Direct": "Emballages Direct",
    "Emballages Pro": "Emballages Pro",
    "Emballages Carrousel": "Emballages Carrousel",
    "Empaques del Norte": "Empaques del Norte",
    "EmpaSoluciones": "EmpaSoluciones",
    "Kyana Packaging Solutions": "Kyana Packaging Solutions",
    "Lancaster Packaging": "Lancaster Packaging",
    "PaperPack": "PaperPack",
    "Partners Packaging": "Partners Packaging",
    "Ace Paper & Packaging": "Ace Paper & Packaging",
    "Gillis & Lane": "Gillis & Lane",
    "Pro Pac Materials": "Pro Pac Materials",
    "Skybox Packaging": "Skybox Packaging",
    "PakFactory": "PakFactory",
    "Box and Wrap": "Box and Wrap",
    "Davenport Packaging": "Davenport Packaging",
    "The Packaging People": "The Packaging People",
    "Packaging Online": "Packaging Online",
    "Packaging Plus": "Packaging Plus",
    "Eco Packaging": "Eco Packaging",
    "Packaging Supplies Direct": "Packaging Supplies Direct",
    "Complete Packaging Solutions": "Complete Packaging Solutions",
    "Direct Packaging Supplies": "Direct Packaging Supplies",
    "Australian Packaging Solutions": "Australian Packaging Solutions",
    "Action Packaging": "Action Packaging",
    "Able Packaging": "Able Packaging",
    "Aeropak": "Aeropak",
    "Bags and Boxes": "Bags and Boxes",
    "Bespoke Packaging": "Bespoke Packaging",
    "Diamond Packaging": "Diamond Packaging",
    "Fast Packaging": "Fast Packaging",
    "First Packaging": "First Packaging",
    "Ernest Packaging": "Ernest Packaging",
    "Stickel Packaging Supply": "Stickel Packaging Supply",
    "SSI Packaging Group": "SSI Packaging Group",
    "Elkay Plastics": "Elkay Plastics",
    "Pioneer Packaging": "Pioneer Packaging",
    "MLR Packaging": "MLR Packaging",
    "Discount Packaging": "Discount Packaging",
    "Pittsburgh Packaging Company": "Pittsburgh Packaging Company",
    "Box Shop": "The Box Shop",
    "Direct Supply Packaging": "Direct Supply",
    "OSI Packaging": "OSI Packaging",
    "The Packaging Source": "The Packaging Source",
    "The Packaging Depot": "The Packaging Depot",
    "Packaging Depot": "Packaging Depot",
    "Aramsco": "Aramsco",
    "Network Distribution": "Network Distribution",
    "Pollock Orora": "Pollock",
    "Boxed": "Boxed",
    "Verpackungsdiscount": "Verpackungsdiscount",
}


def clean_mojibake(name: str) -> str:
    # Fix common Excel/cp1252 mojibake for accented Spanish names
    fixes = {
        "M�xico": "Mexico",
        "Pac�fico": "Pacifico",
        "MÃ©xico": "Mexico",
        "PacÃ­fico": "Pacifico",
    }
    out = name
    for a, b in fixes.items():
        out = out.replace(a, b)
    # Also try latin-1/cp1252 roundtrip for leftover replacement chars
    if "�" in out:
        out = (
            out.replace("Envases de M�xico", "Envases de Mexico")
            .replace("del Pac�fico", "del Pacifico")
        )
    return out


def main() -> None:
    wb = load_workbook(PATH)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ci = headers.index("Company")
    wi = headers.index("Website")

    changes = []
    for row in range(2, ws.max_row + 1):
        old = str(ws.cell(row, ci + 1).value or "").strip()
        website = str(ws.cell(row, wi + 1).value or "").strip()
        new = RENAMES.get(old, old)
        new = clean_mojibake(new).strip()
        # strip leftover trailing Dist if any missed
        new = re.sub(r"\s+Dist$", "", new).strip()
        if new != old:
            ws.cell(row, ci + 1).value = new
            changes.append({"row": row - 1, "old": old, "new": new, "website": website})

    # uniqueness check
    names = [str(ws.cell(r, ci + 1).value or "").strip() for r in range(2, ws.max_row + 1)]
    by_name: dict[str, list[int]] = defaultdict(list)
    for i, n in enumerate(names, start=1):
        by_name[n.lower()].append(i)

    dups = {k: v for k, v in by_name.items() if len(v) > 1}
    if dups:
        # resolve remaining dups with website hostname token (last resort, still proper-ish)
        for key, idxs in dups.items():
            for idx in idxs:
                r = idx + 1
                cur = str(ws.cell(r, ci + 1).value or "").strip()
                web = str(ws.cell(r, wi + 1).value or "").strip().lower()
                # Specific known collisions
                if "globalindustrial.ca" in web:
                    cur2 = "Global Industrial Canada"
                elif "globalindustrial.com" in web:
                    cur2 = "Global Industrial"
                elif "thepackagingcompany.ca" in web:
                    cur2 = "The Packaging Company Canada"
                elif "thepackagingcompany.com" in web and "canada" not in cur.lower():
                    cur2 = "The Packaging Company"
                elif "packagingproducts.co.nz" in web:
                    cur2 = "Packaging Products New Zealand"
                elif "packagingproductscorp.com" in web:
                    cur2 = "Packaging Products Corp"
                elif "packagingsupplies.com.au" in web:
                    cur2 = "Packaging Supplies Australia"
                elif "pack-supplies.co.uk" in web:
                    cur2 = "Pack Supplies"
                else:
                    cur2 = cur
                if cur2 != cur:
                    changes.append(
                        {
                            "row": idx,
                            "old": cur,
                            "new": cur2,
                            "website": web,
                            "note": "collision_fix",
                        }
                    )
                    ws.cell(r, ci + 1).value = cur2

    names2 = [str(ws.cell(r, ci + 1).value or "").strip() for r in range(2, ws.max_row + 1)]
    by_name2 = defaultdict(list)
    for i, n in enumerate(names2, start=1):
        by_name2[n.lower()].append(i)
    dups2 = {k: v for k, v in by_name2.items() if len(v) > 1}

    wb.save(PATH)
    audit = {
        "changed": len(changes),
        "total": len(names2),
        "unique_names": len(set(x.lower() for x in names2)),
        "duplicate_name_keys": dups2,
        "changes": changes,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"changed={len(changes)} total={len(names2)} unique={audit['unique_names']}")
    if dups2:
        print("REMAINING_DUPS", json.dumps(dups2))
    else:
        print("ALL_NAMES_UNIQUE")
    for c in changes[:40]:
        print(f"  {c['old']} -> {c['new']}")
    if len(changes) > 40:
        print(f"  ... +{len(changes)-40} more")


if __name__ == "__main__":
    main()
