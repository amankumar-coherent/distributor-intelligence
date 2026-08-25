"""Batch B: more web-verified Founded years (empty cells only)."""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "founded_fill_audit_batch_b.json"

VERIFIED: dict[str, tuple[str, str]] = {
    "Paper Mart": ("1921", "papermart.com about / 102-year press"),
    "Lindenmeyr Munroe": ("1859", "lindenmeyrmunroe.com about"),
    "Midland Paper": ("1907", "midlandco.com history"),
    "Midland Paper Packaging": ("1907", "midlandco.com history"),
    "CCL Industries": ("1951", "CCL Industries founding Gordon Lang 1951"),
    "Cosmo Films": ("1976", "Wikipedia Cosmo Films founded 1976 (plant 1981)"),
    "UFlex": ("1985", "uflexltd company presentation"),
    "Polyplex": ("1984", "Polyplex Corporation established 1984"),
    "Transcontinental Inc.": ("1976", "TC Transcontinental founded 1976"),
    "IPS Packaging": ("1976", "ipack.com our-story"),
    "Kaufman Container": ("1910", "Kaufman Container LinkedIn / company since 1910"),
    "O.Berk Company": ("1910", "oberk.com history Osias Berk 1910"),
    "SKS Bottle & Packaging": ("1986", "SKS Bottle founding Paul/Barbara Horan 1986"),
    "MSC Industrial Supply Co.": ("1941", "Wikipedia MSC Industrial"),
    "Global Industrial": ("1949", "Global Industrial / Systemax corporate history"),
    "Quill.com Packaging": ("1956", "Quill Corporation founded 1956"),
    "Zoro Tools": ("2011", "Zoro launched by Grainger 2011"),
    "HD Supply": ("2007", "HD Supply spun from Home Depot 2007"),
    "WebstaurantStore": ("2004", "WebstaurantStore founding"),
    "Amazon Business Packaging": ("2015", "Amazon Business launch 2015"),
    "Cole-Parmer Packaging Dist": ("1955", "Cole-Parmer founded 1955"),
    "Fisher Scientific Packaging Dist": ("1902", "Fisher Scientific founded 1902"),
    "Ahlsell": ("1877", "Ahlsell founded 1877"),
    "MISUMI Japan": ("1963", "MISUMI founded 1963"),
    "Trusco Nakayama": ("1959", "Trusco Nakayama founded 1959"),
    "IndustryBuying Packaging": ("2012", "IndustryBuying founded 2012"),
    "Blackwoods AU": ("1878", "Blackwoods Australia established 1878"),
    "NZ Safety Blackwoods": ("1878", "Blackwoods heritage 1878"),
    "Packlane": ("2014", "Packlane founded 2014"),
    "Noissue": ("2017", "noissue founded 2017"),
    "Viking Direct UK": ("1985", "Viking Direct founded 1985"),
    "Viking Direct Ireland": ("1985", "Viking Direct 1985"),
    "Cormack Packaging": ("1957", "Cormack Packaging AU history"),
    "Cospak Pty Ltd.": ("1975", "Cospak company history"),
    "Detpak": ("1945", "Detpak / Detmold Group"),
    "Plastpack Australia": ("1928", "Plastpack Churchill & Coombes 1928"),
    "Waltons ZA": ("1942", "Waltons SA history"),
    "Springpack": ("1994", "springpack.co.uk history"),
    "Davpack": ("1985", "davpack history"),
    "Cleveland Containers": ("1984", "clevelandcontainers history"),
    "Polybags Limited": ("1986", "polybags.co.uk about"),
    "Swiftpak": ("1984", "swiftpak history"),
    "Allpack": ("1974", "Allpack UK / Macfarlane history"),
    "Allpack Packaging": ("1974", "Allpack 1974"),
    "The Cary Company": ("1895", "thecarycompany about"),
    "Royal Bag": ("1976", "royalbag history"),
    "TapeCase": ("1969", "tapecase history"),
    "Nashville Wraps": ("1988", "nashvillewraps history"),
    "Associated Bag": ("1938", "associatedbag.com — confirm; may already filled"),
    "International Plastics": ("1968", "interplas history"),
    "Interplas": ("1968", "International Plastics"),
    "Laddawn Packaging Dist": ("1977", "Laddawn founded 1977"),
    "Wrap-Tite": ("1986", "Wrap-Tite history"),
    "Schermerhorn Bros. Co.": ("1893", "Schermerhorn Bros history"),
    "Richards Packaging": ("1912", "Richards Packaging history"),
    "Crawford Packaging": ("1953", "Crawford Packaging history"),
    "Dayton Bag & Burlap": ("1910", "Dayton Bag history"),
    "Freund Container & Supply": ("1938", "Freund Container history"),
    "Ratioform": ("1979", "Ratioform founded 1979"),
    "Ratioform Italy": ("1979", "Ratioform 1979"),
    "Network Packaging": ("1988", "Network Packaging / Macfarlane"),
    "Constantia Flexibles": ("2004", "Constantia Flexibles formation 2004"),
    "Suedpack": ("1963", "Suedpack history"),
    "Walki": ("1930", "Walki history"),
    "Wipak": ("1969", "Wipak / Wihuri"),
    "Papier-Mettler": ("1957", "Papier-Mettler history"),
    "Schoeller Allibert": ("1959", "Schoeller Allibert history"),
    "SIG": ("1853", "SIG company history — industrial group; packaging later"),
    "Airgas Inc.": ("1982", "Airgas founded 1982"),
    "Fastenal Packaging Supplies": ("1967", "Fastenal 1967"),
    "Grainger Packaging Supplies": ("1927", "Grainger 1927"),
    "MSC Direct Packaging": ("1941", "MSC 1941"),
    "MSC Industrial Packaging": ("1941", "MSC 1941"),
    "Global Industrial Packaging": ("1949", "Global Industrial"),
    "Global Industrial Corp": ("1949", "Global Industrial"),
    "Staples Business": ("1986", "Staples 1986"),
    "Staples.com Business Packaging": ("1986", "Staples 1986"),
    "Office Depot Business": ("1986", "Office Depot 1986"),
    "Quill": ("1956", "Quill 1956"),
    "Zoro": ("2011", "Zoro 2011"),
    "Lyreco Germany": ("1926", "Lyreco 1926"),
    "Lyreco Corporate": ("1926", "Lyreco 1926"),
    "Antalis UK": ("2000", "Antalis 2000"),
    "Antalis Packaging UK": ("2000", "Antalis 2000"),
    "DaklaPack NL": ("1975", "DaklaPack 1975"),
    "Packhelp Poland": ("2015", "Packhelp 2015"),
    "EcoEnclose Packaging": ("2010", "EcoEnclose 2010"),
    "Preferred Plastics": ("1983", "Preferred Plastics 1983"),
    "Noissue Packaging": ("2017", "noissue 2017"),
    "Packlane Custom": ("2014", "Packlane 2014"),
    "Askul": ("1993", "Askul 1993"),
    "MonotaRO": ("2000", "MonotaRO 2000"),
    "MISUMI JP": ("1963", "MISUMI 1963"),
    "Trusco JP": ("1959", "Trusco 1959"),
    "IndustryBuying": ("2012", "IndustryBuying 2012"),
    "IndustryBuying IN": ("2012", "IndustryBuying 2012"),
    "Moglix": ("2015", "Moglix 2015"),
    "Moglix IN": ("2015", "Moglix 2015"),
    "Officeworks": ("1994", "Officeworks 1994"),
    "Signet AU": ("1968", "Signet Australia — verify before keep"),
    "Makro ZA": ("1971", "Makro SA operations"),
    "Makro South Africa": ("1971", "Makro"),
    "BradyPLUS": ("1957", "Brady Industries Las Vegas 1957 roots"),
    "United Packaging Group": ("SKIP", ""),
}


# Keep only strongly re-confirmed in this batch from live sources;
# drop ones not confirmed in this conversation's web results.
KEEP = {
    "Paper Mart",
    "Lindenmeyr Munroe",
    "Midland Paper",
    "Midland Paper Packaging",
    "CCL Industries",
    "Cosmo Films",
    "UFlex",
    "Polyplex",
    "Transcontinental Inc.",
    "IPS Packaging",
    "Kaufman Container",
    "O.Berk Company",
    "SKS Bottle & Packaging",
    "MSC Industrial Supply Co.",
    "Global Industrial",
    "Quill.com Packaging",
    "Zoro Tools",
    "HD Supply",
    "WebstaurantStore",
    "Amazon Business Packaging",
    "Cole-Parmer Packaging Dist",
    "Fisher Scientific Packaging Dist",
    "Ahlsell",
    "MISUMI Japan",
    "Trusco Nakayama",
    "IndustryBuying Packaging",
    "Blackwoods AU",
    "NZ Safety Blackwoods",
    "Packlane",
    "Noissue",
    "Viking Direct UK",
    "Viking Direct Ireland",
    "Laddawn Packaging Dist",
}


def main() -> None:
    clean = {
        k: v
        for k, v in VERIFIED.items()
        if k in KEEP and re.fullmatch(r"(18|19|20)\d{2}", v[0])
    }

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_founded_b_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    col = {str(h): i + 1 for i, h in enumerate(headers) if h}

    filled = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, col["Company"]).value or "").strip()
        cur = ws.cell(r, col["Founded"]).value
        cur_s = "" if cur is None else str(cur).strip()
        if name not in clean:
            continue
        if cur_s and cur_s not in ("None", "N/A", "Not publicly disclosed", "-"):
            continue
        year, src = clean[name]
        ws.cell(r, col["Founded"]).value = year
        filled.append({"company": name, "founded": year, "source": src})

    wb.save(FP)
    empty = has = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, col["Founded"]).value
        if v is None or str(v).strip() in ("", "None", "N/A", "Not publicly disclosed", "-"):
            empty += 1
        else:
            has += 1

    audit = {
        "backup": str(backup),
        "filled_now": filled,
        "filled_count": len(filled),
        "founded_filled_total": has,
        "founded_empty_total": empty,
        "total_rows": ws.max_row - 1,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"FILLED_NOW={len(filled)} TOTAL_FILLED={has} STILL_EMPTY={empty}")
    for x in filled:
        print(f"  {x['company']}: {x['founded']}")


if __name__ == "__main__":
    main()
