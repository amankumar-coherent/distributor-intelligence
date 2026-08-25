# -*- coding: utf-8 -*-
"""
Fill / verify Ownership for global flexible packaging FINAL.
Only web-evidenced values. Never invent owners.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "ownership_verify_audit.json"

# Hand-verified from official IR / press / company pages (this session + prior).
HAND: dict[str, str] = {
    # --- Corrections / normalize existing ---
    "Bunzl": "Public (LSE: BNZL)",
    "Fastenal": "Public (NASDAQ: FAST)",
    "Global Industrial": "Public (NYSE: GIC)",
    "The Home Depot": "Public (NYSE: HD)",
    "Uline": "Private (family-owned)",
    "Lyreco": "Private (Lyreco Group / family)",
    "RAJA": "Private (RAJA Group / family)",
    "Antalis": "Subsidiary of KPP Group Holdings (TSE: 9274)",
    "Igepa Group": "Private (Igepa group)",
    "Inapa": "Public (Euronext Lisbon: INA)",
    "Papyrus": "Subsidiary of OptiGroup (PE: Altor / Ardian / FSN Capital)",
    "Papier Union": "Subsidiary of Inapa (Germany operations / Inapa Deutschland)",
    "Spicers": "Subsidiary of KPP Group Holdings (TSE: 9274)",
    "Spicers Paper": "Subsidiary of Central National-Gottesman (CNG)",
    "Signet": "Subsidiary of Spicers Group (KPP Group Holdings)",
    "Veritiv": "Private (Clayton, Dubilier & Rice)",
    "Cospak": "Subsidiary of San Miguel Yamamura Australasia (SMYPC JV)",
    "Packaging2Buy": "Subsidiary of Bunzl plc",
    "Kite Packaging": "Private (employee-owned / ESOP)",
    "BOX Partners": "Private (Charlesbank Capital Partners-backed)",
    "Kiwi Plastic Packaging": "Private (New Zealand-owned)",
    "SupplyOne": "Private",
    "Shippers Supply": "Subsidiary of Spicers Canada / CNG",
    "Dollies and Boxes Unlimited": "Subsidiary of SupplyOne Canada (Crownhill family)",
    "Mr Box Online": "Subsidiary of SupplyOne",
    "Hughes Enterprises": "Subsidiary of BradyPLUS / Imperial Brady",
    "Individual FoodService": "Subsidiary of BradyPLUS / Imperial Brady",
    "Sigma Supply": "Private (WBENC; BradyPLUS network)",
    "Adelaide Packaging Supplies": "Private (family-owned)",

    # --- Empty / NPD filled from web evidence ---
    "Ahlsell": "Private (CVC Capital Partners-backed)",
    "Askul": "Public (TSE: 2678); major shareholder LY Corporation",
    "MonotaRO": "Public (TSE: 3064); majority Grainger Global Holdings",
    "IndustryBuying": "Subsidiary of MonotaRO (IB MONOTARO)",
    "Manutan": "Private (Guichard family; delisted Euronext)",
    "Officeworks": "Subsidiary of Wesfarmers (ASX: WES)",
    "Blackwoods": "Subsidiary of Wesfarmers (ASX: WES)",
    "Staples": "Private (Sycamore Partners)",
    "Quill": "Subsidiary of Staples (Sycamore Partners)",
    "Zoro Tools": "Subsidiary of W.W. Grainger (NYSE: GWW)",
    "HD Supply": "Subsidiary of The Home Depot (NYSE: HD)",
    "Imperial Dade": "Private (Imperial Brady; Bain / Kelso / Advent / Warburg / FEMSA / Tillis)",
    "BradyPLUS": "Private (Imperial Brady; Bain / Kelso / Advent / Warburg / FEMSA / Tillis)",
    "Imperial Bag & Paper": "Private (Imperial Dade / Imperial Brady group)",
    "All American Containers": "Subsidiary of TricorBraun (ex-Veritiv Containers / AAC)",
    "Richards Packaging": "Public (TSX: RIC) — Richards Group Inc.",
    "Novvia Group": "Private (Kelso & Company-backed)",
    "Viking Direct": "Subsidiary of RAJA Group",
    "JPG": "Subsidiary of RAJA Group (Mondoffice France)",
    "Kalamazoo": "Subsidiary of RAJA Group (Mondoffice Spain)",
    "Mondoffice": "Subsidiary of RAJA Group",
    "Ratioform": "Subsidiary of TAKKT AG (kaiserkraft / Haniel)",
    "Shorr Packaging": "Private (100% employee-owned / ESOP)",
    "MAC Papers": "Private (Monomoy Capital Partners-backed)",
    "Midland Paper": "Private (management-owned)",
    "McMaster-Carr": "Private (family-owned)",
    "WebstaurantStore": "Subsidiary of Clark Associates (family-owned)",
    "Hoffmann Verpackung": "Private (family-owned / Coenen)",
    "Papier-Mettler": "Private (family-owned)",
    "Gerdmans": "Subsidiary of TAKKT AG (Haniel)",
    "Moglix": "Private (VC-backed; Tiger Global / Accel / Alpha Wave)",
    "Office Depot": "Subsidiary of Grupo Gigante (Office Depot de Mexico license)",
    "Waltons": "Subsidiary of Bidvest Group",
    "Packaging House": "Related to NZ Safety Blackwoods / Wesfarmers Industrial & Safety",
    "Tantraco": "Private",
    "Seow Khim Polythene": "Private",
    "Atlantic Packaging Products": "Private",
    "Brangs + Heinrich": "Private (family-owned / Brook-Risse Gruppe)",
    "NNZ": "Private (family-owned)",
    "NNZ GmbH": "Subsidiary of NNZ",
    "Associated Bag": "Private",
    "Crawford Packaging": "Private",
    "Polyair": "Subsidiary of Intertape Polymer Group (Clearlake Capital)",
    "Hisco": "Subsidiary of Distribution Solutions Group (NASDAQ: DSGR)",
    "Macfarlane Packaging": "Subsidiary of Macfarlane Group plc (LSE: MACF)",
    "PRODINGER Verpackung": "Private (PRODINGER Gruppe)",
    "Abco Kovex": "Subsidiary of Bunzl plc",
    "Global Industrial Canada": "Subsidiary of Global Industrial Company (NYSE: GIC)",
    "GWP Group": "Subsidiary of Macfarlane Group plc (LSE: MACF)",
    "Davpack": "Subsidiary of TAKKT AG (Ratioform / Davpack)",
    "O.Berk": "Subsidiary of Berlin Packaging (Oak Hill / CPP Investments)",
    "Delaware Valley Packaging": "Subsidiary of Berlin Packaging (Diablo Valley Packaging)",
    "SKS Bottle & Packaging": "Subsidiary of Pipeline Packaging",
    "Unicorr Packaging": "Private (family-owned)",
    "United One Source": "Private",
    "The Cary Company": "Private",
    "Springpack": "Private",
    "Kaufman Container": "Private",
    "The Boxery": "Private",
    "Karton.eu": "Private",
}


def is_empty(v) -> bool:
    if v is None:
        return True
    s = str(v).strip().lower()
    return s in ("", "not publicly disclosed", "n/a", "na", "none", "-")


def main() -> None:
    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ci = headers.index("Company")
    oi = headers.index("Ownership")

    updated = []
    missing = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, ci + 1).value
        if not name or name not in HAND:
            continue
        new_v = HAND[name]
        old = ws.cell(row, oi + 1).value
        if str(old or "").strip() == new_v:
            continue
        ws.cell(row, oi + 1).value = new_v
        updated.append({"company": name, "old": old, "new": new_v})

    for name in HAND:
        if not any(ws.cell(r, ci + 1).value == name for r in range(2, ws.max_row + 1)):
            missing.append(name)

    # Stats
    filled = empty = 0
    for row in range(2, ws.max_row + 1):
        if is_empty(ws.cell(row, oi + 1).value):
            empty += 1
        else:
            filled += 1

    wb.save(FP)
    audit = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "hand_keys": len(HAND),
        "updated": len(updated),
        "missing_in_sheet": missing,
        "filled_after": filled,
        "empty_after": empty,
        "rows": updated,
        "note": "Evidence-only ownership; blanks left where not verified",
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"updated={len(updated)} filled={filled} empty={empty} missing={missing}")
    print(f"saved {FP}")
    print(f"audit {AUDIT}")


if __name__ == "__main__":
    main()
