# -*- coding: utf-8 -*-
"""Pass2: drop remaining fabricated clones; restore verified websites only."""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_unique_website_fix_pass2.json"

# Verified this session via web search — keep even if local DNS flaky.
RESTORE: dict[str, str] = {
    "Cofares": "https://www.cofares.es",
    "Comifar": "https://www.phoenixpharmaitalia.it",
    "Asahi Kasei Medical Asia Pacific": "https://www.asahi-kasei.co.jp/medical/en/",
    "Kenya Medical Supplies Authority (KEMSA)": "https://kemsa.go.ke",
    "Healthcare 21": "https://healthcare21.eu",
    "Farmacias Guadalajara": "https://www.farmaciasguadalajara.com",
    "Matsumotokiyoshi Holdings": "https://www.matsukiyococokara.com",
    "Imperial Health Sciences": "https://www.imperiallogistics.com/imperial-health-sciences.php",
    "Surgipharm": "https://surgipharm.com",
    "Bidvest Medical": "https://www.bidvest.co.za",
    "Japan Medical Supply Co., Ltd.": "https://www.jms.cc",
    "BML Inc.": "https://www.bml.co.jp",
    "Suzuken Co., Ltd.": "https://www.suzuken.co.jp",
    "Toho Holdings Co., Ltd.": "https://www.tohohd.co.jp",
    "Medipal Holdings Corporation": "https://www.medipal.co.jp",
    "Alfresa Holdings Corporation": "https://www.alfresa.com",
    "Mitsui & Co.": "https://www.mitsui.com",
    "Welch Allyn Australia": "https://www.hillrom.com",
    "Wellness Warehouse": "https://www.wellnesswarehouse.com",
    "Anixter": "https://www.anixter.com",
    "Owens & Minor": "https://www.owens-minor.com",
    "Saudi Pharmaceutical Industries & Medical Appliances Corporation (SPIMACO)": "https://spimaco.com.sa",
}

# Clear parked / wrong domains
CLEAR_ALWAYS = {
    "Al-Zahraa Medical",
}

# Fabricated / unverifiable name patterns to drop (country-stamped clones, empty web).
DROP_EXACT = {
    "Al-Ahli Medical Supplies",
    "Al-Faisal Medical Systems",
    "Al-Saya Medical Company",
    "Al-Sharq Medical",
    "Al-Tayyar Medical",
    "Al-Zahraa Medical",
    "Bionics UK",
    "Care Innovations",
    "Carena Co., Ltd.",
    "CareTech Vertrieb",
    "Cirúrgica Brasil Central",
    "Cirúrgica Campelo",
    "Cirúrgica Nacional",
    "Cirúrgica Rio Preto",
    "Cirúrgica São Judas",
    "Cirúrgica São Lucas",
    "Cirúrgica Vida",
    "Cyprus Medical Supplies",
    "DACH Medical Group",
    "Droguería Central",
    "Droguería del Pacífico",
    "Egypt Medical Supplies",
    "Equipamentos Médicos",
    "Farmacias Arrocha",
    "Farmacias del Dr. Simoni",
    "Farmacias del Este",
    "Farmacias del Pacífico",
    "Farmacias Económicas",
    "Farmacias El Amal",
    "Farmacias La Salle",
    "Farmacias Mutuales",
    "Farmacias Peruanas",
    "Farmacias San Cayetano",
    "Farmacias San Jorge",
    "Farmacias San Nicolás",
    "Farmacias San Rafael",
    "Farmacias San Roque",
    "Farmacias Santa Lucía",
    "Farmacias Unidas",
    "Farmacéutica del Mediterráneo",
    "Greece Medical Supplies",
    "Grupo Farmacéutico Mexicano",
    "Gulf Medical Equipment",
    "Gulf Medical Supplies",
    "Gulf Medical Trading",
    "Hanmi Healthcare",
    "Health Care Equipment & Services",
    "Health Care Wholesalers",
    "Health Equipment Distributors",
    "Health Equipment Retailers",
    "Health Supplies Africa",
    "HealthCare Equipment Centre",
    "Healthcare Pro",
    "Healthcare Solutions UK",
    "Healthcare Wholesale UK",
    "HealthMed Brasil",
    "HealthTech Distribution",
    "Iran Medical Supplies",
    "Iraq Medical Supplies",
    "Israel Medical Supplies",
    "Jordan Medical Supplies",
    "Kawanishi Holdings",
    "Lebanon Medical Supplies",
    "Malta Medical Supplies",
    "Medi Arabia",
    "Medi-Rad Associates",
    "Medi-Rad Malaysia",
    "Medic Distribuidora",
    "Medical Depot Australia",
    "Medical Distributors",
    "Medical Equipment Supplies Australia",
    "Medical Supplies Australia",
    "Medical Wholesale UK",
    "Medicall Brasil",
    "Medicall Distribuidora",
    "Medicall Equipamentos",
    "Medicall Home Care",
    "Medicall Produtos Hospitalares",
    "Medicina e Saúde",
    "MediImport Africa",
    "MediMark",
    "MediPro GmbH",
    "MediSource Gulf",
    "MediSupplies Direct",
    "MediTech Africa",
    "MediTrade Gulf",
    "MedTech Distribuidora",
    "MedVida",
    "Middle East Medical Supplies",
    "Modern Medical Equipment",
    "Morocco Medical Supplies",
    "Oman Medical Supplies",
    "Palestine Medical Supplies",
    "Prodimed",
    "Sanitätshaus Becker",
    "Sanitätshaus Weber",
    "Saudi Medical Devices",
    "Saudi Medical Supplies",
    "Shenzhen Aeon Technology",
    "Shenzhen Creative Medical",
    "Surgical House",
    "Syria Medical Supplies",
    "The Medical Equipment Centre",
    "The Medical Equipment Company",
    "The Medical Supply Company",
    "Tunisia Medical Supplies",
    "Turkey Medical Supplies",
    "UK Health Supplies",
    "UK Medical Distributors",
    "UK Medical Equipment",
    "UK Medical Wholesale",
    "Wellness Forever",
    "Wellness Pharmacy",
    "Yemen Medical Supplies",
}


def main() -> None:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(XLSX.stem + f"_BEFORE_PASS2_{ts}.xlsx")
    shutil.copy2(XLSX, backup)

    wb = load_workbook(XLSX)
    ws = wb.active
    hdr = [str(c.value or "").strip() for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(hdr) if h}

    dropped = []
    restored = []
    renamed = []

    # Collect rows
    rows = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        row = {h: ("" if ws.cell(r, ic[h]).value is None else str(ws.cell(r, ic[h]).value).strip()) for h in ic}
        rows.append(row)

    before = len(rows)
    keep = []
    for row in rows:
        co = row["Company"]
        if co in DROP_EXACT or co in CLEAR_ALWAYS and co in DROP_EXACT:
            dropped.append(co)
            continue
        if co in DROP_EXACT:
            dropped.append(co)
            continue
        keep.append(row)

    # also drop CLEAR_ALWAYS if still present with parked site
    rows2 = []
    for row in keep:
        co = row["Company"]
        if co in CLEAR_ALWAYS:
            dropped.append(co)
            continue
        if co in RESTORE:
            old = row.get("Website", "")
            row["Website"] = RESTORE[co]
            if old != RESTORE[co]:
                restored.append({"company": co, "from": old, "to": RESTORE[co]})
        # rename Matsumotokiyoshi Holdings display optional
        if co == "Matsumotokiyoshi Holdings":
            row["Company"] = "MatsukiyoCocokara & Co. (Matsumotokiyoshi)"
            renamed.append({"from": co, "to": row["Company"]})
            # re-key restore already applied
        rows2.append(row)

    # dedupe by company
    seen = set()
    final = []
    for row in rows2:
        key = re.sub(r"\s+", " ", row["Company"]).strip().lower()
        if key in seen:
            dropped.append(row["Company"] + " (dup)")
            continue
        seen.add(key)
        final.append(row)

    final.sort(key=lambda r: r["Company"].lower())

    # rewrite sheet
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    for i, row in enumerate(final, start=2):
        for h, col in ic.items():
            ws.cell(i, col).value = row.get(h, "")
    wb.save(XLSX)
    wb.close()

    empty = sum(1 for r in final if not r.get("Website"))
    audit = {
        "timestamp": ts,
        "backup": str(backup),
        "before": before,
        "after": len(final),
        "dropped": dropped,
        "restored": restored,
        "renamed": renamed,
        "empty_websites": empty,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("before", before, "after", len(final), "dropped", len(dropped))
    print("restored", len(restored), "empty", empty)
    print("backup", backup)


if __name__ == "__main__":
    main()
