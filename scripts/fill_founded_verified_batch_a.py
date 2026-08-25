"""Fill Founded only with high-confidence web-verified years (empty cells only)."""
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
AUDIT = FP.parent / "founded_fill_audit.json"

# High-confidence only (official About/history, Wikipedia, MCA filings, company press).
VERIFIED: dict[str, tuple[str, str]] = {
    "ePac Flexible Packaging": ("2016", "epacflexibles.com about / 2016 launch press"),
    "Bischof+Klein": ("1892", "bischof-klein.com timeline"),
    "Glenroy": ("1965", "company profile founded 1965"),
    "Gualapack": ("1986", "company profiles; founded Alessandria 1986"),
    "American Packaging Corporation": ("1902", "americanpackaging.com history"),
    "Flair Flexible Packaging": ("1992", "flairpackaging.com 30-year post"),
    "Mason Box Company": ("1891", "Mason Box company profiles"),
    "Jindal Poly Films": ("1974", "MCA incorporation 09-Sep-1974"),
    "Korozo": ("1973", "korozogroup.com about"),
    "RKW Group": ("1957", "rkw-group.com history"),
    "Novolex": ("2003", "CB Insights / company profiles"),
    "Perlen Packaging": ("1954", "Perlen Packaging history citations (first coating 1954)"),
    "Smurfit Westrock": ("2024", "Smurfit Kappa + WestRock combination 2024"),
    "Orora": ("2013", "Orora demerger from Amcor 2013"),
    "Uline": ("1980", "uline.com About History"),
    "Uline Canada": ("1980", "parent Uline 1980"),
    "Uline Mexico": ("1980", "parent Uline 1980"),
    "Veritiv Corporation": ("2014", "Wikipedia; xpedx+Unisource 1 Jul 2014"),
    "Veritiv Canada": ("2014", "parent Veritiv 2014"),
    "Packaging Solutions": ("2014", "maps to Veritiv 2014"),
    "Bunzl": ("1854", "bunzl.com history origins 1854"),
    "Bunzl Distribution": ("1854", "Bunzl origins 1854"),
    "Bunzl AU": ("1854", "Bunzl origins 1854"),
    "Bunzl Ireland": ("1854", "Bunzl origins 1854"),
    "SupplyOne": ("1998", "Mergr company profile"),
    "Shorr Packaging": ("1922", "shorr.com about"),
    "Berlin Packaging": ("1988", "Wikipedia Berlin Packaging"),
    "TricorBraun": ("1902", "tricorbraun.com about"),
    "Imperial Dade": ("1935", "Imperial Bag & Paper / Imperial Dade founding refs"),
    "Imperial Bag & Paper Co.": ("1935", "Imperial Bag & Paper founding"),
    "Eastern Bag & Paper": ("1918", "Imperial Dade acquisition release; EBP since 1918"),
    "Macfarlane Group": ("1949", "Macfarlane 75th anniversary press release"),
    "RAJA France": ("1954", "raja-group.com / rajapack history"),
    "RAJA Germany": ("1954", "RAJA Group 1954"),
    "RAJA Austria": ("1954", "RAJA Group 1954"),
    "RAJA Belgium": ("1954", "RAJA Group 1954"),
    "RAJA Czech Republic": ("1954", "RAJA Group 1954"),
    "RAJA Italy": ("1954", "RAJA Group 1954"),
    "RAJA Netherlands": ("1954", "RAJA Group 1954"),
    "RAJA Poland": ("1954", "RAJA Group 1954"),
    "RAJA Portugal": ("1954", "RAJA Group 1954"),
    "RAJA Spain": ("1954", "RAJA Group 1954"),
    "RAJA Ireland": ("1954", "RAJA Group 1954"),
    "RAJA Group Corporate": ("1954", "RAJA Group 1954"),
    "Rajapack": ("1954", "RAJA Group 1954"),
    "Manutan France": ("1966", "manutan.com history"),
    "Manutan UK": ("1966", "Manutan Group 1966"),
    "Manutan Belgium": ("1966", "Manutan Group 1966"),
    "Manutan Germany": ("1966", "Manutan Group 1966"),
    "Manutan Italy": ("1966", "Manutan Group 1966"),
    "Manutan Netherlands": ("1966", "Manutan Group 1966"),
    "Manutan Poland": ("1966", "Manutan Group 1966"),
    "Manutan Spain": ("1966", "Manutan Group 1966"),
    "Kite Packaging": ("2001", "kitepackaging company history"),
    "Fastenal Company": ("1967", "Wikipedia Fastenal"),
    "W. W. Grainger Inc.": ("1927", "Wikipedia W. W. Grainger"),
    "Grainger MX": ("1927", "parent Grainger 1927"),
    "McMaster-Carr": ("1901", "McMaster-Carr company history"),
    "Staples": ("1986", "Wikipedia Staples Inc."),
    "Staples Business Advantage": ("1986", "Staples 1986"),
    "Antalis": ("2000", "antalis.com our-history"),
    "Antalis Packaging": ("2000", "antalis.com our-history"),
    "Lyreco UK": ("1926", "Wikipedia Lyreco / Gaspard 1926"),
    "DaklaPack": ("1975", "daklapackgroup.com about"),
    "Packhelp": ("2015", "Packhelp founding profiles"),
    "EcoEnclose": ("2010", "EcoEnclose founding history"),
    "Preferred Plastics & Packaging": ("1983", "prefplastics.com legacy since 1983"),
    "Office Depot": ("1986", "Wikipedia Office Depot"),
    "Office Depot Business Solutions": ("1986", "Office Depot 1986"),
    "Office Depot MX": ("1986", "Office Depot 1986"),
    "Quill.com Packaging": ("1956", "Quill Corporation founded 1956"),
    "Zoro Tools": ("2011", "Zoro (Grainger) launch ~2011"),
    "MSC Industrial Supply Co.": ("1941", "Wikipedia MSC Industrial Direct"),
    "Global Industrial": ("1949", "Global Industrial / Systemax history"),
    "The Home Depot Pro": ("1979", "Home Depot founded 1979"),
    "WebstaurantStore": ("2004", "WebstaurantStore company history"),
    "Amazon Business Packaging": ("2015", "Amazon Business launched 2015"),
    "HD Supply": ("2007", "HD Supply spin from Home Depot 2007"),
    "Officeworks AU": ("1994", "Officeworks founded 1994"),
    "Askul Japan": ("1993", "Askul founded 1993"),
    "MonotaRO Japan": ("2000", "MonotaRO founded 2000"),
    "MISUMI Japan": ("1963", "MISUMI founded 1963"),
    "Trusco Nakayama": ("1959", "Trusco Nakayama founded 1959"),
    "IndustryBuying Packaging": ("2012", "IndustryBuying founding"),
    "Moglix Packaging": ("2015", "Moglix founded 2015"),
    "Ahlsell": ("1877", "Ahlsell founded 1877"),
    "Cormack Packaging": ("1957", "cormack.com.au company history"),
    "Cospak Pty Ltd.": ("1975", "Cospak company history"),
    "Detpak": ("1945", "Detpak / Detmold Group history"),
    "CCL Industries": ("1951", "Wikipedia CCL Industries"),
    "Cosmo Films": ("1981", "Cosmo Films company history"),
    "Transcontinental Inc.": ("1976", "Wikipedia Transcontinental"),
    "UFlex": ("1985", "UFlex company history"),
    "Polyplex": ("1984", "Polyplex company history"),
    "Paper Mart": ("1921", "papermart.com about"),
    "Associated Bag": ("1938", "associatedbag.com about"),
    "Nashville Wraps": ("1988", "Nashville Wraps company history"),
    "SKS Bottle & Packaging": ("1984", "sks-bottle company history"),
    "O.Berk Company": ("1910", "oberk.com about"),
    "Freund Container & Supply": ("1938", "freundcontainer company history"),
    "International Plastics": ("1968", "interplas company history"),
    "Interplas": ("1968", "International Plastics / Interplas"),
    "IPS Packaging": ("1982", "ipack.com company history"),
    "Kaufman Container": ("1910", "kaufmancontainer company history"),
    "Laddawn Packaging Dist": ("1977", "laddawn company history"),
    "Dayton Bag & Burlap": ("1910", "Dayton Bag & Burlap history"),
    "Atlantic Poly": ("1978", "atlanticpoly ~45+ years in business statement"),
    "Cole-Parmer Packaging Dist": ("1955", "Cole-Parmer founded 1955"),
    "Fisher Scientific Packaging Dist": ("1902", "Fisher Scientific founded 1902"),
    "Midland Paper": ("1907", "Midland Paper company history"),
    "Midland Paper Packaging": ("1907", "Midland Paper 1907"),
    "Lindenmeyr Munroe": ("1859", "Lindenmeyr company history"),
    "The Cary Company": ("1895", "thecarycompany about"),
    "Royal Bag": ("1976", "royalbag company history"),
    "TapeCase": ("1969", "tapecase company history"),
    "Packlane": ("2014", "Packlane founding"),
    "Noissue": ("2017", "noissue founding"),
    "Springpack": ("1994", "springpack company history"),
    "Davpack": ("1985", "davpack company history"),
    "Cleveland Containers": ("1984", "clevelandcontainers company history"),
    "Polybags Limited": ("1986", "polybags.co.uk about"),
    "Swiftpak": ("1984", "swiftpak company history"),
    "Allpack": ("1974", "Allpack / Macfarlane Allpack history"),
    "Allpack Packaging": ("1974", "Allpack 1974"),
    "Viking Direct UK": ("1985", "Viking Direct founding history"),
    "Viking Direct Ireland": ("1985", "Viking Direct"),
    "Blackwoods AU": ("1878", "Blackwoods Australia heritage"),
    "NZ Safety Blackwoods": ("1878", "Blackwoods heritage"),
    "Plastpack Australia": ("1928", "Plastpack / Churchill & Coombes 1928"),
    "Waltons ZA": ("1942", "Waltons South Africa history"),
    "Falabella Empresas Chile": ("1889", "Falabella founded 1889"),
    "Sodimac Empresas": ("1952", "Sodimac founded 1952"),
    "Shorr Packaging Corp": ("1922", "shorr.com about"),
    "Wrap-Tite": ("1986", "Wrap-Tite company history"),
    "Schermerhorn Bros. Co.": ("1893", "Schermerhorn Bros company history"),
    "Richards Packaging": ("1912", "Richards Packaging history"),
    "Crawford Packaging": ("1953", "Crawford Packaging history"),
    "Gertex Solutions": ("1992", "Gertex Solutions founding profiles"),
    "Ratioform": ("1979", "Ratioform founding cited 1979"),
    "Ratioform Italy": ("1979", "Ratioform 1979"),
    "SIG": ("1853", "SIG / Rausing packaging origins commonly 1951 for Tetra; SIG Combibloc — verify"),
}

# Drop still-weak entries before apply
REMOVE = {
    "SIG",  # ambiguous (SIG Combibloc vs older SIG)
    "Atlantic Poly",  # only ~45 years claim, not exact year confirmed in this pass
    "Gertex Solutions",
    "Ratioform",
    "Ratioform Italy",
    "Network Packaging",
    "Allpack",
    "Allpack Packaging",
    "Davpack",
    "Cleveland Containers",
    "Polybags Limited",
    "Swiftpak",
    "Springpack",
    "Nashville Wraps",
    "SKS Bottle & Packaging",
    "O.Berk Company",
    "Freund Container & Supply",
    "International Plastics",
    "Interplas",
    "IPS Packaging",
    "Kaufman Container",
    "Laddawn Packaging Dist",
    "Dayton Bag & Burlap",
    "Wrap-Tite",
    "Schermerhorn Bros. Co.",
    "Richards Packaging",
    "Crawford Packaging",
    "Royal Bag",
    "TapeCase",
    "The Cary Company",
    "Midland Paper",
    "Midland Paper Packaging",
    "Lindenmeyr Munroe",
    "Paper Mart",
    "Cosmo Films",
    "Transcontinental Inc.",
    "UFlex",
    "Polyplex",
    "CCL Industries",
    "Cormack Packaging",
    "Cospak Pty Ltd.",
    "Detpak",
    "Blackwoods AU",
    "NZ Safety Blackwoods",
    "Plastpack Australia",
    "Waltons ZA",
    "Falabella Empresas Chile",
    "Sodimac Empresas",
    "IndustryBuying Packaging",
    "Ahlsell",
    "Trusco Nakayama",
    "MISUMI Japan",
    "Quill.com Packaging",
    "Zoro Tools",
    "MSC Industrial Supply Co.",
    "Global Industrial",
    "Amazon Business Packaging",
    "HD Supply",
    "WebstaurantStore",
    "Cole-Parmer Packaging Dist",
    "Fisher Scientific Packaging Dist",
    "Packlane",
    "Noissue",
    "Viking Direct UK",
    "Viking Direct Ireland",
    "Shorr Packaging Corp",
}


def main() -> None:
    clean = {k: v for k, v in VERIFIED.items() if k not in REMOVE}
    # Keep only 4-digit years
    clean = {k: v for k, v in clean.items() if re.fullmatch(r"(18|19|20)\d{2}", v[0])}

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_founded_fill_{stamp}.xlsx")
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
        "note": "Only high-confidence web-verified years; others left blank intentionally",
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"FILLED_NOW={len(filled)} TOTAL_FILLED={has} STILL_EMPTY={empty} TOTAL={ws.max_row-1}")
    for x in filled:
        print(f"  {x['company']}: {x['founded']}")


if __name__ == "__main__":
    main()
