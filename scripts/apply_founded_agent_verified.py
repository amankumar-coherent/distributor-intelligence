"""One-shot apply of agent-verified Founded years (no Google AI scraper).
Only fills empty/non-digit cells. Never overwrites existing digit years unless FORCE.
"""
from __future__ import annotations

import json
from pathlib import Path

from openpyxl import load_workbook

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "founded_agent_verified_apply.json"

# High-confidence only: About pages, Companies House/ACRA, Wikipedia, official history
VERIFIED: dict[str, str] = {
    "Abco Kovex": "1987",
    "Acme Paper & Supply": "1946",
    "All Packaging Company": "1946",
    "Aropack": "2016",
    "Atlantic Packaging Products": "1945",
    "Auspackaging": "2011",
    "Bags Plus": "1985",
    "Ball & Doggett": "2017",
    "BigDug": "2004",
    "Box Brothers": "1985",
    "Boxworks": "2019",
    "Brown Paper Packaging": "2003",
    "Burch Bottle & Packaging": "1978",
    "BVS Verpacken": "1997",
    "CPS Supply Line": "1963",
    "Davpack": "1968",
    "Dollies and Boxes Unlimited": "1986",
    "Durable Packaging": "1944",
    "Empaques del Norte": "1966",
    "Ernest Packaging": "1946",
    "Excel Pac": "1997",
    "Fillmore Container": "1995",
    "Fastpak Systems": "1986",
    "Gerdmans": "1947",
    "GM Packaging": "2001",
    "GWP Group": "1990",
    "Hoffmann Verpackung": "1920",
    "Hughes Enterprises": "1950",
    "Hisco": "1970",
    "Igepa Group": "1960",
    "Inapa": "1965",
    "Innovative Packaging Group": "2006",
    "Kiwi Plastic Packaging": "2001",
    "Knueppel Verpackung": "1919",
    "Lesters Packaging": "1983",
    "Luv2Pak": "1917",
    "Lyreco": "1926",
    "MaxPak": "1979",
    "McLernon Packaging": "1932",
    "National Packaging Services": "1987",
    "Network Packaging": "1997",
    "Packaging Chimp": "2005",
    "Packaging Panda": "2005",
    "Packaging Unlimited": "1975",
    "Papier Union": "1988",
    "Papyrus": "1895",
    "Paul Hildebrandt": "1889",
    "Pluspak": "1981",
    "Premier Packaging": "2011",
    "Seal Packaging": "1990",
    "Seow Khim Polythene": "1979",
    "Shardlows Packaging": "1965",
    "Spicers": "1896",
    "Springpack": "1975",
    "Swiftpak": "1977",
    "The Bottle Crew": "2000",
    "The Box Depot": "1986",
    "The Boxery": "2010",
    "Trinity Packaging Supply": "2010",
    "Uboxes": "2004",
    "Winc": "1986",
    "BB-Verpackungen": "2006",
    "Karton.eu": "2007",
    "Transpak": "1975",
    "W.B. Mason": "1898",
    "TigerPak": "2002",
    "Aramsco": "1966",
    "PackQueen": "2007",
    "Surepak": "1986",
    "Baas Verpakkingen": "1935",
    "Weller Packaging": "1975",
    "UPack": "2018",
    "Elkay Plastics": "1968",
    "Pioneer Packaging": "1976",
    "Direct Supply": "1985",
    "Tillmann Verpackungen": "1952",
    "Horna": "1976",
    "ESWE Versandpack": "1985",
    "Tantraco": "1978",
    "Global Industrial Canada": "1949",
    "Verpacken Online": "1875",
    "ANW Packaging": "2002",
    "The Packaging Centre": "1966",
    "Ozpack": "2010",
    "Furodet Emballages": "2016",
    "De Wild": "2021",
    "Brangs + Heinrich": "1875",
    "Packen24": "2005",
    # well-known already often filled
    "Uline": "1980",
    "Bunzl": "1854",
    "RAJA": "1954",
    "Manutan": "1966",
    "Ratioform": "1979",
    "Antalis": "2000",
    "Staples": "1986",
    "Office Depot": "1986",
    "Quill": "1956",
    "Zoro Tools": "2011",
    "MSC Industrial Supply": "1941",
    "Global Industrial": "1949",
    "The Home Depot": "1979",
    "WebstaurantStore": "2004",
    "HD Supply": "2007",
    "Officeworks": "1994",
    "Askul": "1993",
    "MonotaRO": "2000",
    "IndustryBuying": "2012",
    "Moglix": "2015",
    "Ahlsell": "1877",
    "Cospak": "1975",
    "Paper Mart": "1921",
    "Associated Bag": "1938",
    "Nashville Wraps": "1988",
    "SKS Bottle & Packaging": "1984",
    "O.Berk": "1910",
    "Freund Container & Supply": "1938",
    "IPS Packaging": "1982",
    "Kaufman Container": "1910",
    "Dayton Bag & Burlap": "1910",
    "Midland Paper": "1907",
    "The Cary Company": "1895",
    "Allpack": "1974",
    "Viking Direct": "1985",
    "Blackwoods": "1878",
    "Waltons": "1942",
    "Shorr Packaging": "1922",
    "Schermerhorn Bros.": "1893",
    "Richards Packaging": "1912",
    "Crawford Packaging": "1953",
    "Preferred Plastics & Packaging": "1983",
    "Kite Packaging": "2001",
    "Fastenal": "1967",
    "Grainger": "1927",
    "McMaster-Carr": "1901",
    "SupplyOne": "1998",
    "Imperial Dade": "1935",
    "Imperial Bag & Paper": "1935",
    "Veritiv": "2014",
    "NNZ": "1922",
    "NNZ GmbH": "1922",
    "Ehlken Verpackungen": "1953",
    "Kyana Packaging Solutions": "1976",
    "PaperPack": "1999",
    "Partners Packaging": "2009",
    "Lancaster Packaging": "1989",
    "Associated Packaging": "1977",
}


def main() -> None:
    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ci = headers.index("Company") + 1
    fi = headers.index("Founded") + 1
    wi = headers.index("Website") + 1

    applied = []
    kept = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, ci).value or "").strip()
        cur = str(ws.cell(r, fi).value or "").strip()
        if name not in VERIFIED:
            continue
        year = VERIFIED[name]
        if cur.isdigit():
            kept.append({"company": name, "year": cur})
            continue
        ws.cell(r, fi).value = year
        applied.append({"company": name, "year": year})

    wb.save(FP)

    filled = sum(
        1
        for r in range(2, ws.max_row + 1)
        if str(ws.cell(r, fi).value or "").strip().isdigit()
    )
    miss = []
    for r in range(2, ws.max_row + 1):
        v = str(ws.cell(r, fi).value or "").strip()
        if not v.isdigit():
            miss.append(
                {
                    "company": ws.cell(r, ci).value,
                    "website": ws.cell(r, wi).value,
                    "founded": v,
                }
            )

    (FP.parent / "founded_missing_152.json").write_text(
        json.dumps(miss, indent=2), encoding="utf-8"
    )
    audit = {
        "applied_now": len(applied),
        "already_had": len(kept),
        "filled_total": filled,
        "missing_total": len(miss),
        "applied": applied,
        "missing_companies": [m["company"] for m in miss],
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"applied={len(applied)} filled={filled} missing={len(miss)}")
    for a in applied[:30]:
        print("SET", a["company"], a["year"])


if __name__ == "__main__":
    main()
