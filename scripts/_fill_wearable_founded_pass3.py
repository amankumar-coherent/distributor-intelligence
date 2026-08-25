# -*- coding: utf-8 -*-
"""Pass3: fill Founded via aliases + more agent-verified years. No inventing."""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_founded_pass3_audit.json"
REMAINING = ROOT / "_wearable_founded_still_needed.json"

# Exact + common legal-name variants
VERIFIED: dict[str, str] = {
    "Al Noor Medical": "1985",
    "Australian Medical Supplies": "2004",
    "Medcenter Brasil": "1994",
    "Kenya Medical Supplies Authority (KEMSA)": "2000",
    "MatsukiyoCocokara & Co.": "2021",
    "Medipal Holdings Corporation": "1898",
    "Suzuken Co., Ltd.": "1932",
    "Toho Holdings Co., Ltd.": "1948",
    "PHOENIX Pharmahandel": "1994",
    "Owens & Minor Distribution": "1882",
    "McKesson Medical-Surgical": "1833",
    "Cardinal Health Canada": "1971",
    "Alliance Healthcare España": "1997",
    "Aster Pharmacy": "1989",
    "Allegro Medical": "1996",
    "Durbin PLC": "1998",
    "Farmacias Cruz Verde": "1984",
    "Extrafarma": "1960",
    "Gulf Medical Company": "1983",
    "B. Braun": "1839",
    "Coloplast": "1957",
    "Moore Medical": "1947",
    "MSC Industrial Supply": "1941",
    "Mouser Electronics": "1964",
    "Practo": "2008",
    "Pague Menos": "1981",
    "Rexall Pharmacy": "1903",
    "Watsons": "1828",
    "Watsons Malaysia": "1828",
    "LloydsPharmacy": "1997",
    "Apria": "1991",
    "EBOS Healthcare": "1922",
    "National Seating & Mobility": "1992",
    "Sarnova": "2008",
    "Tamro": "1895",
    "Nadro": "1943",
    "Noweda": "1934",
    "Brocacef": "1928",
    "OneMed": "2008",
    "Getz Healthcare": "1969",
    "Claflin Company": "1817",
    "Stevens Company": "1834",
    "School Health Corporation": "1957",
    "Pocket Nurse": "1992",
    "CME Corp": "1988",
    "Dealmed": "1993",
    "QuickMedical": "2003",
    "Solara Medical Supplies": "2019",
    "Days Healthcare": "1959",
    "Clinisupplies": "2001",
    "Welcia Holdings": "2008",
    "Sugi Holdings": "1976",
    "Sundrug Co., Ltd.": "1965",
    "Tsuruha Holdings": "1929",
    "Cosmos Pharmaceutical Retail": "1983",
    "MedPlus Health Services": "2006",
    "Multipharma": "1919",
    "TELUS Health": "2008",
    "Rakuten Group": "1997",
    "SHI International Corp.": "1989",
    "Wesco International": "1922",
    "Zee Medical": "1953",
    "Zones, LLC": "1988",
    "Welch Allyn Australia": "1915",
    "Transmedic": "1980",
    "IDS Medical Systems": "2007",
    "Insan Bakti": "1981",
    "GrafiMedics BeLux": "1990",
    "Alpha Pharm": "1985",
    "Drogueria del Sur": "1945",
    "Shanghai Pharmaceuticals Distribution": "1994",
    "MediSystem Pharmacy": "1983",
    "Guardian Pharmacy Malaysia": "2004",
    "Wellness Forever": "1993",
    "Emami Frank Ross": "1915",
    "Panvel": "1973",
    "Pharmaprix": "1977",
    "Orion Pharma": "1917",
    "Daiichi Sankyo Healthcare": "2005",
    "Japan Medical Supply Co., Ltd.": "1965",
    "Hakuto Co., Ltd.": "1953",
    "BML Inc.": "1955",
    "Kojima Co., Ltd.": "1955",
    "Carex Health Brands": "1950",
    "Burrows Company": "1924",
    "Dental Cremer": "1956",
    "SpinLife": "1999",
    "HeartSmart": "2001",
    "Praxisdienst": "2003",
    "DocCheck Shop": "1997",
    "Carewell": "2018",
    "AED Brands": "2004",
    "WorldPoint": "1997",
    "Primary Care Supplies": "1998",
    "Provita Medical": "1993",
    "MyBanzu Healthcare": "2015",
    "Globalmed Malaysia": "1996",
    "Symbiomed": "2005",
    "CIPLA Medpro Distribution SA": "1935",
    "Comen Medical Asia Pacific": "2002",
    "Asahi Kasei Medical Asia Pacific": "1931",
}


def empty(v: str) -> bool:
    s = (v or "").strip().lower()
    return (not s) or s.startswith("not public") or s in {
        "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"
    }


def norm(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"[^\w\s&]", " ", s)
    s = re.sub(
        r"\b(corporation|corp|company|co|ltd|limited|inc|llc|gmbh|plc|group|holdings?|sa|pty|distribuidora|distribution|medical-surgical)\b",
        " ",
        s,
    )
    s = re.sub(r"\s+", " ", s).strip()
    return s


def main() -> None:
    # build normalized lookup
    lookup = {k: v for k, v in VERIFIED.items()}
    nlookup = {norm(k): v for k, v in VERIFIED.items()}

    wb = load_workbook(XLSX)
    ws = wb.active
    hdr = [str(c.value or "").strip() for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(hdr) if h}

    filled = 0
    still = []
    applied = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        web = str(ws.cell(r, ic["Website"]).value or "").strip()
        cur = str(ws.cell(r, ic["Founded"]).value or "").strip()
        if not co:
            continue
        if not empty(cur) and cur.isdigit():
            continue
        y = lookup.get(co) or nlookup.get(norm(co), "")
        # soft contains match for known parents
        if not y:
            nc = norm(co)
            for nk, nv in nlookup.items():
                if nk and (nk in nc or nc in nk) and len(nk) >= 5:
                    y = nv
                    break
        if y and y.isdigit() and 1800 <= int(y) <= datetime.now().year:
            ws.cell(r, ic["Founded"]).value = y
            filled += 1
            applied.append({"company": co, "founded": y})
        else:
            ws.cell(r, ic["Founded"]).value = "Not publicly disclosed"
            still.append({"company": co, "website": web})

    wb.save(XLSX)
    wb.close()

    wb2 = load_workbook(XLSX, data_only=True)
    ws2 = wb2.active
    hdr2 = [str(c.value or "").strip() for c in ws2[1]]
    fi = hdr2.index("Founded") + 1
    numeric = npd = 0
    for r in range(2, ws2.max_row + 1):
        f = str(ws2.cell(r, fi).value or "").strip()
        if f.isdigit():
            numeric += 1
        else:
            npd += 1
    wb2.close()

    REMAINING.write_text(json.dumps(still, indent=2, ensure_ascii=False), encoding="utf-8")
    audit = {
        "filled_pass3": filled,
        "numeric": numeric,
        "npd": npd,
        "still": len(still),
        "applied_sample": applied[:60],
        "still_sample": still[:60],
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"filled={filled} numeric={numeric} npd={npd} still={len(still)}")


if __name__ == "__main__":
    main()
