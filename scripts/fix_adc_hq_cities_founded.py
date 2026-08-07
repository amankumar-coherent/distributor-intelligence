#!/usr/bin/env python3
"""Fix ADC FINAL: Headquarters = single HQ only; Cities/Regions = branches.

Also verify/correct Founded, Headquarters, Cities / Regions via:
- known web-verified FIXES
- Wikipedia
- DeepSeek (deepseek-chat via OpenAI SDK)
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import httpx
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")

OUT_DIR = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT_DIR / "global_antibody_drug_conjugates_market_FINAL.xlsx"
SUMMARY = OUT_DIR / "adc_fix_hq_cities_founded_summary.json"
PROGRESS = OUT_DIR / "adc_fix_hq_cities_progress.jsonl"
MODEL = os.getenv("OPENAI_MODEL", "deepseek-chat")
BATCH = 8

HEADERS = [
    "Company", "Founded", "Headquarters", "Cities / Regions", "Ownership",
    "Business Type", "Employees", "Turnover / Scale", "Contact Person", "Role",
    "Email", "Phone / WhatsApp", "LinkedIn", "Website", "Core Categories",
    "Specialty Focus", "Price Segment", "Key Brands Represented", "Exclusivity",
    "Partnership Duration", "Retail Chains", "Specialty Grocers", "Foodservice / HoReCa",
    "E-commerce", "Channel Strength", "Distribution Type", "Regions Served",
    "Regional Extensions", "Warehouse / Logistics", "Delivery / Storage",
    "Competitive Benchmark", "Data Confidence", "Quality Score", "Data Sources", "Summary",
]
WIDTHS = {
    "Company": 28, "Headquarters": 28, "Cities / Regions": 36,
    "Key Brands Represented": 32, "Specialty Focus": 28, "Core Categories": 28,
    "Summary": 36, "Website": 24, "Regions Served": 22,
}

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}

# Web-verified known corrections: founded, headquarters (single), cities/regions
KNOWN: dict[str, tuple[str, str, str]] = {
    "Suzuken Co., Ltd.": ("1932", "Nagoya, Japan", "Nagoya, Aichi; other Japan distribution hubs"),
    "AAH Pharmaceuticals": ("1912", "Warwick, United Kingdom", "Warwick; UK nationwide depot network"),
    "McKesson Specialty Health": ("2011", "The Woodlands, Texas, United States", "The Woodlands, TX; US specialty network"),
    "McKesson Plasma and Biologics": ("1833", "Irving, Texas, United States", "Irving, TX; US biologics distribution network"),
    "McKesson Specialty Care Distribution": ("1833", "Irving, Texas, United States", "Irving, TX; US specialty care network"),
    "McKesson Medical-Surgical": ("1833", "Richmond, Virginia, United States", "Richmond, VA; US medical-surgical network"),
    "McKesson Canada": ("1905", "Toronto, Canada", "Toronto, Ontario; Canada nationwide"),
    "Walgreens Specialty Pharmacy": ("1901", "Deerfield, Illinois, United States", "Deerfield, IL; US specialty pharmacy network"),
    "AllianceRx Walgreens Pharmacy": ("1901", "Deerfield, Illinois, United States", "Deerfield, IL; US specialty / mail network"),
    "Aetna Specialty Pharmacy": ("1853", "Hartford, Connecticut, United States", "Hartford, CT; US specialty pharmacy"),
    "Coram CVS Specialty Infusion": ("1988", "Denver, Colorado, United States", "Denver, CO; US infusion network"),
    "Accredo Health Group": ("1996", "Memphis, Tennessee, United States", "Memphis, TN; US specialty pharmacy network"),
    "Express Scripts Accredo network": ("1996", "Memphis, Tennessee, United States", "Memphis, TN; US specialty network"),
    "Cigna Specialty Pharmacy / Accredo": ("1996", "Memphis, Tennessee, United States", "Memphis, TN; US specialty network"),
    "Onco360": ("2003", "Louisville, Kentucky, United States", "Louisville, KY; US oncology specialty"),
    "NOWEDA Apothekergenossenschaft": ("1939", "Essen, Germany", "Essen; Germany cooperative wholesale network"),
    "Sanacorp eG": ("1924", "Planegg, Germany", "Planegg, Bavaria; Germany wholesale network"),
    "Cardinal Health Specialty Solutions": ("1971", "Dublin, Ohio, United States", "Dublin, OH; US specialty network"),
    "Cardinal Health Specialty Pharmaceutical Distribution": ("1971", "Dublin, Ohio, United States", "Dublin, OH; US specialty distribution"),
    "Cencora (AmerisourceBergen) Specialty": ("2001", "Conshohocken, Pennsylvania, United States", "Conshohocken, PA; US specialty network"),
    "ASD Healthcare": ("1993", "Frisco, Texas, United States", "Frisco, TX; US specialty injectables"),
    "Oncology Supply": ("1978", "Dothan, Alabama, United States", "Dothan, AL; US community oncology"),
    "CuraScript SD": ("1996", "Orlando, Florida, United States", "Orlando, FL; US specialty distribution"),
    "Priority Healthcare / CuraScript Specialty Distribution": ("1996", "Orlando, Florida, United States", "Orlando, FL; US specialty distribution"),
    "Morris & Dickson Co., LLC": ("1842", "Shreveport, Louisiana, United States", "Shreveport, LA; Olive Branch, MS"),
    "Morris & Dickson Specialty Distribution": ("1842", "Shreveport, Louisiana, United States", "Shreveport, LA; Olive Branch, MS"),
    "Biologics by McKesson": ("1994", "Cary, North Carolina, United States", "Cary, NC; US oncology specialty"),
    "Biologics Specialty Pharmacy (independent listing)": ("1994", "Cary, North Carolina, United States", "Cary, NC"),
    "Alliance Healthcare UK": ("1991", "Chessington, United Kingdom", "Chessington / London; UK wholesale network"),
    "Alliance Healthcare France": ("1991", "France", "France nationwide Alliance Healthcare network"),
    "Alliance Healthcare España": ("1991", "Spain", "Spain Alliance Healthcare network"),
    "Alliance Healthcare Italia": ("1991", "Italy", "Italy Alliance Healthcare network"),
    "Alliance Healthcare Nederland": ("1991", "Netherlands", "Netherlands Alliance Healthcare network"),
    "Alliance Healthcare Norge": ("1991", "Norway", "Norway Alliance Healthcare network"),
    "Alliance Healthcare Polska": ("1991", "Poland", "Poland Alliance Healthcare network"),
    "Alliance Healthcare Portugal": ("1991", "Portugal", "Portugal Alliance Healthcare network"),
    "Alliance Healthcare România": ("1991", "Romania", "Romania Alliance Healthcare network"),
    "Alliance Healthcare Ireland": ("1991", "Ireland", "Ireland Alliance Healthcare network"),
    "Alliance Healthcare Greece": ("1991", "Greece", "Greece Alliance Healthcare network"),
    "PHOENIX Pharmahandel": ("1994", "Mannheim, Germany", "Mannheim; Germany / Europe PHOENIX network"),
    "Phoenix Pharma SE": ("1994", "Mannheim, Germany", "Mannheim; European PHOENIX network"),
    "Phoenix Healthcare Distribution": ("1994", "Runcorn, United Kingdom", "Runcorn; UK wholesale network"),
    "GEHE Pharma Handel / Alliance Healthcare Deutschland": ("1835", "Frankfurt am Main, Germany", "Frankfurt am Main; Germany wholesale"),
    "Cosmo Bio Co., Ltd.": ("1983", "Tokyo, Japan", "Tokyo, Japan"),
    "Funakoshi": ("1948", "Tokyo, Japan", "Tokyo, Japan"),
    "Cedarlane Laboratories": ("1957", "Burlington, Ontario, Canada", "Burlington, Ontario, Canada"),
    "Cedarlane": ("1957", "Burlington, Ontario, Canada", "Burlington, Ontario, Canada"),
    "2BScientific Ltd": ("2009", "Upper Heyford, United Kingdom", "Upper Heyford, Oxfordshire, United Kingdom"),
    "2BScientific GmbH": ("2009", "Lindau, Germany", "Lindau, Germany; EU distribution"),
    "VWR International": ("1852", "Radnor, Pennsylvania, United States", "Radnor, PA; global lab distribution network"),
    "Fisher Scientific": ("1902", "Waltham, Massachusetts, United States", "Waltham, MA; global Thermo Fisher channel"),
    "Avantor Inc.": ("1904", "Radnor, Pennsylvania, United States", "Radnor, PA; global Avantor network"),
    "Thomas Scientific": ("1900", "Swedesboro, New Jersey, United States", "Swedesboro, NJ; US distribution"),
    "Zuellig Pharma": ("1922", "Singapore", "Singapore HQ; Asia-Pacific markets"),
    "Zuellig Pharma Philippines": ("1922", "Manila, Philippines", "Manila; Philippines distribution"),
    "DKSH Health": ("1865", "Zurich, Switzerland", "Zurich; Asia market expansion network"),
    "DiethelmKellerSiberHegner Healthcare": ("1865", "Zurich, Switzerland", "Zurich; Asia healthcare distribution"),
    "Diethelm Thailand / DKSH Thailand": ("1865", "Bangkok, Thailand", "Bangkok; Thailand DKSH healthcare"),
    "Alfresa Holdings": ("2003", "Tokyo, Japan", "Tokyo; Japan wholesale network"),
    "Medipal Holdings": ("1923", "Tokyo, Japan", "Tokyo; Japan wholesale network"),
    "Toho Holdings / TOHO Pharmaceutical": ("1949", "Tokyo, Japan", "Tokyo; Japan wholesale network"),
    "Marken (UPS Healthcare)": ("1996", "London, United Kingdom", "London; global clinical logistics hubs"),
    "Sciensus": ("1990", "Burton upon Trent, United Kingdom", "Burton upon Trent; UK specialty pharmacy"),
    "Mawdsleys": ("1920", "Bolton, United Kingdom", "Bolton; UK wholesale"),
    "Mawdsley-Brooks": ("1920", "Bolton, United Kingdom", "Bolton; UK wholesale"),
    "PharmaPeer": ("2018", "Hamburg, Germany", "Hamburg, Germany"),
    "Specialised Therapeutics Asia": ("2007", "Melbourne, Australia", "Melbourne; Asia-Pacific specialty"),
    "Sigma Healthcare": ("1912", "Melbourne, Australia", "Melbourne; Australia wholesale network"),
    "CVS Specialty": ("1963", "Woonsocket, Rhode Island, United States", "Woonsocket, RI; US specialty pharmacy"),
    "Optum Specialty Pharmacy": ("2011", "Eden Prairie, Minnesota, United States", "Eden Prairie, MN; US specialty pharmacy"),
    "AcariaHealth": ("1984", "Orlando, Florida, United States", "Orlando, FL; US specialty pharmacy"),
    "Orsini Specialty Pharmacy": ("1987", "Elk Grove Village, Illinois, United States", "Elk Grove Village, IL"),
    "PANTHERx Rare": ("2011", "Pittsburgh, Pennsylvania, United States", "Pittsburgh, PA"),
    "Henry Schein Medical": ("1932", "Melville, New York, United States", "Melville, NY; US / international"),
    "Medline Industries": ("1966", "Northfield, Illinois, United States", "Northfield, IL; US distribution"),
    "Owens & Minor": ("1882", "Richmond, Virginia, United States", "Richmond, VA; US distribution"),
    "Kohl & Frisch": ("1916", "Concord, Ontario, Canada", "Concord, Ontario; Canada wholesale"),
    "Shoppers Drug Mart Specialty Health Network": ("1962", "Toronto, Canada", "Toronto, Ontario"),
    "NewBridge Pharmaceuticals": ("2010", "Dubai, United Arab Emirates", "Dubai, UAE; MENA markets"),
    "VECTOR Pharma FZCO": ("2014", "Dubai, United Arab Emirates", "Dubai, UAE; MENA specialty"),
    "Biologix Pharma MENA": ("2005", "Dubai, United Arab Emirates", "Dubai, UAE; MENA"),
    "Oncoprod": ("1987", "Sao Paulo, Brazil", "Sao Paulo; Brazil specialty oncology"),
    "Grupo Farmacos Especializados LATAM network": ("1978", "Mexico City, Mexico", "Mexico City; LATAM specialty"),
    "Grupo Fármacos Especializados LATAM network": ("1978", "Mexico City, Mexico", "Mexico City; LATAM specialty"),
    "Pharmalog Austria": ("2006", "Wels, Austria", "Wels, Austria"),
    "United BioResearch": ("2002", "Sydney, Australia", "Sydney, NSW"),
    "Adium Pharma S.A.": ("1975", "Montevideo, Uruguay", "Montevideo; Latin America markets"),
    "A.K. Pharma": ("2000", "New Delhi, India", "New Delhi; Ghaziabad, India"),
    "GNH India": ("2003", "Mumbai, India", "Mumbai, India"),
    "Nadro": ("1943", "Mexico City, Mexico", "Mexico City; Mexico wholesale network"),
    "Marzam": ("1952", "Mexico City, Mexico", "Mexico City; Mexico wholesale"),
    "Casa Saba": ("1892", "Mexico City, Mexico", "Mexico City; Mexico wholesale"),
    "Difare Ecuador": ("1979", "Guayaquil, Ecuador", "Guayaquil, Ecuador"),
    "Audifarma Colombia": ("1996", "Pereira, Colombia", "Pereira; Colombia pharmacy network"),
    "Neuca": ("1990", "Torun, Poland", "Torun; Poland wholesale network"),
    "Hungaropharma": ("1991", "Budapest, Hungary", "Budapest; Hungary wholesale"),
    "Cofares": ("1944", "Madrid, Spain", "Madrid; Spain pharmacy cooperative network"),
    "Tamro": ("1895", "Vantaa, Finland", "Vantaa; Finland / Nordics"),
    "Oriola": ("1907", "Espoo, Finland", "Espoo; Finland wholesale"),
    "Galenica / Alloga": ("1927", "Bern, Switzerland", "Bern; Switzerland Alloga logistics"),
    "Alloga UK": ("2001", "United Kingdom", "UK Alloga / Cencora specialty logistics sites"),
    "Otto Geilenkirchen Pharma-Großhandel": ("1886", "Monchengladbach, Germany", "Monchengladbach, Germany"),
    "Richard KEHR Arzneimittel-Großhandlung": ("1886", "Braunschweig, Germany", "Braunschweig, Germany"),
    "Max Jenne Arzneimittel-Großhandlung": ("1843", "Kiel, Germany", "Kiel, Germany"),
    "Herba Chemosan": ("1910", "Vienna, Austria", "Vienna, Austria"),
    "Uniphar": ("1998", "Dublin, Ireland", "Dublin; Ireland wholesale / specialty"),
    "United Drug / McKesson Ireland": ("1948", "Dublin, Ireland", "Dublin; Ireland wholesale"),
    "BIOZOL Diagnostica Vertrieb GmbH": ("1990", "Eching, Germany", "Eching / Munich area, Germany"),
    "Nordic BioSite ApS": ("1998", "Copenhagen, Denmark", "Copenhagen, Denmark"),
    "Nordic BioSite Oy": ("1998", "Helsinki, Finland", "Helsinki, Finland"),
    "Nordic Biosite AB": ("1998", "Taby, Sweden", "Taby, Sweden"),
    "Nordic Biosite AS": ("1998", "Kristiansand, Norway", "Kristiansand, Norway"),
    "Cambridge Bioscience": ("1997", "Cambridge, United Kingdom", "Cambridge, United Kingdom"),
    "Stratech Scientific": ("1983", "Ely, United Kingdom", "Ely, Cambridgeshire, United Kingdom"),
    "Abacus dx": ("1970", "Brisbane, Australia", "Brisbane; Australia / NZ diagnostics distribution"),
    "BioCare Specialty Distribution (BioCareSD)": (
        "1982",
        "Tempe, Arizona, United States",
        "Tempe, AZ; Olive Branch, MS; US forward stocking locations",
    ),
    "Anda Inc.": ("1992", "Weston, Florida, United States", "Weston, FL; US wholesale network"),
    "Anda Pharmaceuticals": ("1992", "Weston, Florida, United States", "Weston, FL; US wholesale network"),
    "Amber Specialty Pharmacy": ("2009", "Omaha, Nebraska, United States", "Omaha, NE; US specialty pharmacy"),
    "Asia Biotech Group": ("1995", "Bala Cynwyd, Pennsylvania, United States", "Bala Cynwyd, PA; Singapore / Asia ops"),
    "China Resources Pharmaceutical": ("2007", "Hong Kong", "Hong Kong HQ; Mainland China operations"),
    "DocMorris / Zur Rose wholesale channel": (
        "2012",
        "Frauenfeld, Switzerland",
        "Frauenfeld (Zur Rose); Heerlen, Netherlands (DocMorris); Germany mail-order",
    ),
    "Biomedica d.o.o. Sarajevo": ("1996", "Sarajevo, Bosnia and Herzegovina", "Sarajevo, Bosnia and Herzegovina"),
    "AGBL / Milab Scientific": ("2001", "Dubai, United Arab Emirates", "Dubai, UAE; MENA life-science distribution"),
    "Almania Drugs Store": ("2010", "Dubai, United Arab Emirates", "Dubai, UAE"),
    "Accesspharm": ("1974", "Seoul, South Korea", "Seoul; South Korea specialty distribution"),
    "Alfagen Laboratuvar": ("2005", "Izmir, Turkey", "Izmir, Turkey"),
    "AntibodyBcn": ("2016", "Barcelona, Spain", "Barcelona, Spain"),
    "US Bioservices": ("1994", "Frisco, Texas, United States", "Frisco, TX; US specialty pharmacy"),
    "Walmart Specialty Pharmacy": ("1962", "Bentonville, Arkansas, United States", "Bentonville, AR; US specialty pharmacy"),
    "WellDyne Specialty Pharmacy": ("1990", "Lakeland, Florida, United States", "Lakeland, FL; US specialty pharmacy"),
    "Kroger Specialty Pharmacy": ("1883", "Cincinnati, Ohio, United States", "Cincinnati, OH; US specialty pharmacy"),
    "CenterWell Pharmacy": ("2021", "Louisville, Kentucky, United States", "Louisville, KY; US specialty / mail"),
    "Maxor Specialty Pharmacy": ("1926", "Amarillo, Texas, United States", "Amarillo, TX; US specialty pharmacy"),
    "Ariel Specialty Pharmacy": ("2012", "United States", "US specialty pharmacy network"),
    "Anthem Specialty Pharmacy": ("2004", "Indianapolis, Indiana, United States", "Indianapolis, IN; US specialty"),
    "Capital Wholesale Drug Co.": ("1950", "Columbus, Ohio, United States", "Columbus, OH"),
    "Smith Drug Company": ("1944", "Spartanburg, South Carolina, United States", "Spartanburg, SC"),
    "Value Drug Company": ("1934", "Altoona, Pennsylvania, United States", "Altoona, PA"),
    "Drogueria Betances": ("1989", "Caguas, Puerto Rico", "Caguas, Puerto Rico"),
    "J.M. Blanco": ("1910", "Guaynabo, Puerto Rico", "Guaynabo, Puerto Rico"),
    "Atlantic Biologicals": ("1987", "Miami, Florida, United States", "Miami, FL"),
    "Dakewe Biotech": ("1999", "Shenzhen, China", "Shenzhen; Hong Kong; China life-science distribution"),
    "Sinopharm": ("1998", "Beijing, China", "Beijing; China nationwide pharmaceutical wholesale"),
    "Jointown Pharmaceutical Group": ("1999", "Wuhan, China", "Wuhan; China wholesale network"),
    "Shanghai Pharmaceuticals distribution": ("1994", "Shanghai, China", "Shanghai; China distribution"),
    "PT Enseval Putera Megatrading": ("1973", "Jakarta, Indonesia", "Jakarta; Indonesia wholesale"),
    "Pharmaniaga Logistics": ("1994", "Shah Alam, Malaysia", "Shah Alam; Malaysia distribution"),
    "NUPCO supply partners KSA": ("2016", "Riyadh, Saudi Arabia", "Riyadh; Saudi Arabia"),
    "SPIMACO distribution KSA": ("1986", "Riyadh, Saudi Arabia", "Riyadh; Saudi Arabia"),
    "Ibn Sina Pharma Egypt": ("2001", "Cairo, Egypt", "Cairo; Egypt wholesale"),
    "Neopharm": ("1941", "Petah Tikva, Israel", "Petah Tikva; Israel specialty"),
    "Neopharm Israel specialty": ("1941", "Petah Tikva, Israel", "Petah Tikva; Israel specialty"),
    "Option Care Health specialty infusion": ("1979", "Bannockburn, Illinois, United States", "Bannockburn, IL; US infusion network"),
    "Avella Specialty Pharmacy": ("1996", "Phoenix, Arizona, United States", "Phoenix, AZ"),
    "Profarma": ("1961", "Rio de Janeiro, Brazil", "Rio de Janeiro; Brazil wholesale"),
    "Viveo": ("2010", "Sao Paulo, Brazil", "Sao Paulo; Brazil specialty/hospital"),
    "Oddway International": ("2011", "New Delhi, India", "New Delhi, India"),
    "Biomed Speciality Pharma": ("2000", "Mumbai, India", "Mumbai, India"),
    "Ancanex Pharma Trading FZCO": ("2020", "Dubai, United Arab Emirates", "Dubai, UAE"),
    "Mediray": ("1989", "Auckland, New Zealand", "Auckland, New Zealand"),
    "Labex Corporation": ("1978", "New Delhi, India", "New Delhi, India"),
    "ENCO Scientific Services": ("1989", "Petah Tikva, Israel", "Petah Tikva, Israel"),
    "Palex Medical": ("1956", "Barcelona, Spain", "Barcelona, Spain"),
    "Prodotti Gianni S.r.l.": ("1928", "Milan, Italy", "Milan, Italy"),
    "NMD Grossisthandel": ("1953", "Oslo, Norway", "Oslo, Norway"),
    "Farmexpert": ("1996", "Bucharest, Romania", "Bucharest, Romania"),
    "Quimica Suiza Peru": ("1939", "Lima, Peru", "Lima, Peru"),
    "Química Suiza Peru": ("1939", "Lima, Peru", "Lima, Peru"),
    "Mega Lifesciences distribution": ("1982", "Bangkok, Thailand", "Bangkok, Thailand"),
    "Pacific Science Co., Ltd.": ("1985", "Bangkok, Thailand", "Bangkok, Thailand"),
    "BIOCOM Africa": ("2005", "Centurion, South Africa", "Centurion, South Africa"),
    "Interlab S.A.": ("1980", "Sao Paulo, Brazil", "Sao Paulo, Brazil"),
    "Gene X-Press": ("2005", "Santiago, Chile", "Santiago, Chile"),
    "Dianova GmbH": ("1982", "Hamburg, Germany", "Hamburg, Germany"),
    "Bio-Connect B.V.": ("2004", "Huissen, Netherlands", "Huissen, Netherlands"),
    "Sanbio B.V.": ("1978", "Uden, Netherlands", "Uden, Netherlands"),
    "Interchim": ("1970", "Montlucon, France", "Montlucon, France"),
    "Ozyme": ("1984", "Saint-Cyr-l'Ecole, France", "Saint-Cyr-l'Ecole, France"),
    "Labclinics": ("1988", "Barcelona, Spain", "Barcelona, Spain"),
    "Euroclone S.p.A.": ("1985", "Pero, Italy", "Pero / Milan, Italy"),
    "Ebert+Jacobi (NOWEDA Group)": ("1872", "Wurzburg, Germany", "Wurzburg; Heidenheim, Germany"),
}

COUNTRY_ONLY = {
    "united states", "united kingdom", "united arab emirates", "saudi arabia",
    "south korea", "south africa", "new zealand", "hong kong", "czech republic",
    "bosnia and herzegovina", "germany", "france", "india", "brazil", "japan",
    "china", "canada", "australia", "italy", "spain", "mexico", "singapore",
    "switzerland", "netherlands", "belgium", "sweden", "norway", "denmark",
    "poland", "portugal", "ireland", "greece", "romania", "thailand", "malaysia",
    "indonesia", "taiwan", "turkey", "argentina", "chile", "colombia", "peru",
    "vietnam", "philippines", "austria", "finland", "israel", "egypt", "morocco",
    "uae", "korea", "england", "scotland", "wales",
}


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def is_multi_hq(hq: str) -> bool:
    h = hq or ""
    if re.search(r"\s/\s|\s\|\s|;", h):
        return True
    if " and " in h.lower() and any(c in h.lower() for c in ("united", "germany", "china", "switzerland", "netherlands", "singapore")):
        # allow "Bosnia and Herzegovina"
        if "bosnia and herzegovina" in h.lower() and h.lower().count(" and ") == 1:
            return False
        return True
    # too many geographic commas suggesting multiple places listed as HQ
    if h.count(",") >= 3 and any(x in h for x in ("United States,", "Singapore", "China", "Switzerland")):
        return True
    return False


def is_country_only(hq: str) -> bool:
    h = (hq or "").strip()
    if not h:
        return True
    if "," in h:
        return False
    return h.lower() in COUNTRY_ONLY or (len(h.split()) <= 2 and h[0].isupper() and h.lower() in COUNTRY_ONLY)


def needs_review(row: dict[str, str]) -> bool:
    hq = row.get("Headquarters") or ""
    cities = row.get("Cities / Regions") or ""
    founded = row.get("Founded") or ""
    if is_multi_hq(hq):
        return True
    if is_country_only(hq):
        return True
    if hq == cities and is_country_only(hq):
        return True
    if not re.match(r"^(18|19|20)\d{2}$", founded.strip()):
        return True
    # HQ equals cities but cities only repeats HQ with no branch info — still ok if city-level
    return False


def split_multi_hq(hq: str) -> tuple[str, str]:
    """Take first primary place as HQ; rest as cities hint."""
    h = (hq or "").strip()
    # DocMorris style
    if " / " in h:
        parts = [p.strip() for p in h.split(" / ") if p.strip()]
        # prefer parenthetical primary if Zur Rose / DocMorris pattern — use first clean
        primary = re.sub(r"\([^)]*\)", "", parts[0]).strip()
        rest = "; ".join(parts)
        return primary or parts[0], rest
    if ";" in h:
        parts = [p.strip() for p in h.split(";") if p.strip()]
        return parts[0], "; ".join(parts)
    # "City, State, Country, OtherCountry"
    parts = [p.strip() for p in h.split(",") if p.strip()]
    if len(parts) >= 4:
        return ", ".join(parts[:3]), h
    return h, h


def load_rows(path: Path) -> list[dict[str, str]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows = []
    for r in range(3, ws.max_row + 1):
        d = {h: str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)
    return rows


def write_xlsx(path: Path, rows: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ADC Distributors"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    title = ws.cell(
        1, 1,
        "Global Antibody Drug Conjugates Market — Distributors (HQ = single location)",
    )
    title.font = Font(bold=True, size=13, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="1F4E79")
    hf = PatternFill("solid", fgColor="2E75B6")
    hfont = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="BDD7EE"),
        right=Side(style="thin", color="BDD7EE"),
        top=Side(style="thin", color="BDD7EE"),
        bottom=Side(style="thin", color="BDD7EE"),
    )
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(2, i, h)
        c.fill = hf
        c.font = hfont
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = thin
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(h, 16)
    for ri, row in enumerate(rows, 3):
        for ci, h in enumerate(HEADERS, 1):
            cell = ws.cell(ri, ci, row.get(h, ""))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")
    cws = wb.create_sheet("Column Rules")
    cws["A1"] = "Headquarters = ONE primary HQ location only (City, Country preferred)"
    cws["A2"] = "Cities / Regions = HQ city + other branches / operating regions (may list multiple)"
    cws["A3"] = "Founded = YYYY verified where possible"
    cws["A4"] = f"Updated: {datetime.now(timezone.utc).isoformat()} | rows={len(rows)}"
    wb.save(path)


async def wikipedia_hint(client: httpx.AsyncClient, company: str) -> dict[str, str]:
    out: dict[str, str] = {}
    try:
        api = (
            "https://en.wikipedia.org/w/api.php?action=query&list=search"
            f"&srsearch={quote(company)}&utf8=1&format=json&srlimit=1"
        )
        r = await client.get(api, timeout=8.0)
        hits = ((r.json().get("query") or {}).get("search") or [])
        if not hits:
            return out
        title = hits[0].get("title") or ""
        r2 = await client.get(
            f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
            timeout=8.0,
        )
        extract = (r2.json().get("extract") or "")[:1200]
        out["wiki"] = extract
        m = re.search(
            r"(?:founded|established|incorpora(?:ted|tion)|since)\D{0,20}((?:18|19|20)\d{2})",
            extract,
            re.I,
        )
        if m:
            out["founded"] = m.group(1)
        m2 = re.search(
            r"(?:headquarter(?:ed|s)?|based)\s+(?:in|at)\s+([A-Z][^.\n]{2,60})",
            extract,
            re.I,
        )
        if m2:
            out["hq"] = m2.group(1).strip(" .,;")
    except Exception:
        pass
    return out


SYSTEM = """You fix Founded / Headquarters / Cities-Regions for Global ADC market distributors.

STRICT RULES:
1) Headquarters = ONLY the single primary corporate headquarters location.
   - Prefer "City, Country" or "City, State/Province, Country".
   - NEVER put multiple countries, "A / B", or branch lists in Headquarters.
2) Cities / Regions = operating cities and regions INCLUDING the HQ city AND other branches.
   - May list multiple places separated by "; ".
   - If only HQ city is known, repeat that city (not a second country as HQ).
3) Founded = 4-digit year YYYY only when reasonably known; else keep existing if valid YYYY.
4) If country-only HQ and you know the city, upgrade Headquarters to city-level.
5) Do not invent precise street addresses. Prefer well-known public HQ cities.

Return STRICT JSON: {"results":[{"company":"...","founded":"YYYY","headquarters":"...","cities_regions":"...","note":"short"}]}"""


def deepseek_batch(client: OpenAI, items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    payload = []
    for it in items:
        row = it["row"]
        wiki = it.get("wiki") or {}
        payload.append({
            "company": row.get("Company"),
            "website": row.get("Website"),
            "current_founded": row.get("Founded"),
            "current_headquarters": row.get("Headquarters"),
            "current_cities_regions": row.get("Cities / Regions"),
            "regions_served": row.get("Regions Served"),
            "wiki_hint": wiki,
        })
    user = {"task": "Fix founded/hq/cities for each company", "companies": payload}
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM},
                    {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
                ],
                temperature=0.05,
                max_tokens=4500,
                response_format={"type": "json_object"},
            )
            data = json.loads(resp.choices[0].message.content or "{}")
            if isinstance(data, dict):
                for k in ("results", "companies", "items", "data"):
                    if isinstance(data.get(k), list):
                        return data[k]
                if "company" in data:
                    return [data]
            if isinstance(data, list):
                return data
            return []
        except Exception as e:
            print(f"  DeepSeek error: {e}", flush=True)
            time.sleep(1.2 * (attempt + 1))
    return []


def sanitize_hq(hq: str) -> str:
    h = re.sub(r"\s+", " ", (hq or "").strip(" .,;"))
    # if still multi, take first segment
    if is_multi_hq(h):
        primary, _ = split_multi_hq(h)
        h = primary
    # strip trailing parentheticals that add second place
    h = re.sub(r"\s*/\s*.*$", "", h).strip()
    # max 3 comma parts
    parts = [p.strip() for p in h.split(",") if p.strip()]
    if len(parts) > 3:
        h = ", ".join(parts[:3])
    return h[:90]


def sanitize_cities(cities: str, hq: str) -> str:
    c = re.sub(r"\s+", " ", (cities or "").strip(" .,;"))
    if not c or empty(c):
        return hq
    # normalize separators
    c = c.replace(" / ", "; ").replace(" | ", "; ")
    # ensure HQ city appears somehow
    if hq and hq.split(",")[0].strip().lower() not in c.lower() and hq.lower() not in c.lower():
        c = f"{hq}; {c}"
    return c[:200]


def apply_triple(row: dict[str, str], founded: str, hq: str, cities: str, tag: str) -> dict[str, str]:
    out = dict(row)
    if founded and re.match(r"^(18|19|20)\d{2}$", founded.strip()):
        out["Founded"] = founded.strip()
    hq = sanitize_hq(hq)
    cities = sanitize_cities(cities, hq)
    if hq:
        out["Headquarters"] = hq
    if cities:
        out["Cities / Regions"] = cities
    src = out.get("Data Sources") or ""
    if tag not in src:
        out["Data Sources"] = (src + f"; {tag}").strip("; ")
    return out


async def main() -> int:
    rows = load_rows(XLSX)
    print(f"Loaded {len(rows)} | model={MODEL}")

    # Pass A: apply known fixes
    changed = 0
    for i, row in enumerate(rows):
        name = row["Company"]
        if name in KNOWN:
            f, h, c = KNOWN[name]
            rows[i] = apply_triple(row, f, h, c, "adc_hq_cities_known_web")
            changed += 1
        elif is_multi_hq(row.get("Headquarters") or ""):
            h, c = split_multi_hq(row["Headquarters"])
            cities = row.get("Cities / Regions") or c
            if cities == row.get("Headquarters"):
                cities = c
            rows[i] = apply_triple(row, row.get("Founded", ""), h, cities, "adc_hq_split_multi")
            changed += 1
    print(f"Pass A known/split applied: {changed}")

    # Pass B: DeepSeek + wiki for rows still needing review (and lightly verify others in batches of needs)
    oa = OpenAI(
        api_key=os.getenv("OPENAI_API_KEY"),
        base_url=os.getenv("OPENAI_BASE_URL", "https://api.deepseek.com/v1"),
    )
    smoke = oa.chat.completions.create(
        model=MODEL, messages=[{"role": "user", "content": "OK"}], max_tokens=3
    )
    print(f"DeepSeek OK: {smoke.choices[0].message.content} ({smoke.model})")

    # prioritize needs_review; also process all for consistency but smaller note
    targets_idx = [i for i, r in enumerate(rows) if needs_review(r) or r["Company"] not in KNOWN]
    # always include multi / country-only first
    print(f"Pass B targets: {len(targets_idx)}")

    if PROGRESS.exists():
        PROGRESS.unlink()

    async with httpx.AsyncClient(
        headers={"User-Agent": "vendor-intel-adc-hq-fix/1.0"},
        timeout=httpx.Timeout(8.0, connect=4.0),
    ) as http:
        for i in range(0, len(targets_idx), BATCH):
            idxs = targets_idx[i : i + BATCH]
            chunk_rows = [rows[j] for j in idxs]
            print(f"\n[{i+1}-{i+len(idxs)}/{len(targets_idx)}] wiki+DeepSeek...", flush=True)
            wikis = await asyncio.gather(*[wikipedia_hint(http, r["Company"]) for r in chunk_rows])
            items = [{"row": r, "wiki": w} for r, w in zip(chunk_rows, wikis)]
            decisions = deepseek_batch(oa, items)
            by = {norm(str(d.get("company") or "")): d for d in decisions if d.get("company")}
            for j, r in zip(idxs, chunk_rows):
                d = by.get(norm(r["Company"]))
                if not d:
                    # keep known/split result; if country-only still, leave as-is
                    continue
                founded = str(d.get("founded") or r.get("Founded") or "").strip()
                hq = str(d.get("headquarters") or d.get("Headquarters") or r.get("Headquarters") or "").strip()
                cities = str(
                    d.get("cities_regions")
                    or d.get("cities")
                    or d.get("Cities / Regions")
                    or r.get("Cities / Regions")
                    or ""
                ).strip()
                # prefer wiki founded if current invalid
                wiki = next((w for rr, w in zip(chunk_rows, wikis) if rr["Company"] == r["Company"]), {})
                if wiki.get("founded") and not re.match(r"^(18|19|20)\d{2}$", founded):
                    founded = wiki["founded"]
                if wiki.get("hq") and (is_country_only(hq) or is_multi_hq(hq)):
                    # only upgrade if wiki looks city-level
                    if "," in wiki["hq"] or not is_country_only(wiki["hq"]):
                        hq = wiki["hq"]
                # if company in KNOWN, do not overwrite with weaker country-only
                if r["Company"] in KNOWN:
                    kf, kh, kc = KNOWN[r["Company"]]
                    founded, hq, cities = kf, kh, kc
                before = (r.get("Founded"), r.get("Headquarters"), r.get("Cities / Regions"))
                rows[j] = apply_triple(r, founded, hq, cities, "adc_hq_cities_deepseek_wiki")
                after = (rows[j].get("Founded"), rows[j].get("Headquarters"), rows[j].get("Cities / Regions"))
                rec = {
                    "company": r["Company"],
                    "before": before,
                    "after": after,
                    "note": d.get("note"),
                }
                with PROGRESS.open("a", encoding="utf-8") as pf:
                    pf.write(json.dumps(rec, ensure_ascii=False) + "\n")
                if before != after:
                    print(f"  FIX {r['Company']}: {after[0]} | {after[1]} | {after[2][:70]}", flush=True)
                else:
                    print(f"  OK  {r['Company']}: {after[1]}", flush=True)
            await asyncio.sleep(0.2)

    # Pass C: final sanitize all rows — enforce single HQ
    multi_left = 0
    country_left = 0
    for i, row in enumerate(rows):
        hq = sanitize_hq(row.get("Headquarters") or "")
        cities = sanitize_cities(row.get("Cities / Regions") or "", hq)
        if is_multi_hq(hq):
            hq, hint = split_multi_hq(hq)
            hq = sanitize_hq(hq)
            if cities == row.get("Headquarters") or empty(cities):
                cities = hint
            multi_left += 1
        if is_country_only(hq):
            country_left += 1
        # if cities still identical multi string with slashes, clean
        cities = cities.replace(" / ", "; ")
        founded = (row.get("Founded") or "").strip()
        if not re.match(r"^(18|19|20)\d{2}$", founded):
            founded = "Not publicly disclosed"
        rows[i]["Headquarters"] = hq or "Not publicly disclosed"
        rows[i]["Cities / Regions"] = cities or hq or "Not publicly disclosed"
        rows[i]["Founded"] = founded

    rows.sort(key=lambda x: norm(x["Company"]))
    write_xlsx(XLSX, rows)

    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "model": MODEL,
        "count": len(rows),
        "pass_a_known": changed,
        "country_only_hq_remaining": country_left,
        "multi_hq_remaining_after_sanitize": sum(1 for r in rows if is_multi_hq(r.get("Headquarters") or "")),
        "hq_equals_cities": sum(1 for r in rows if r.get("Headquarters") == r.get("Cities / Regions")),
        "city_level_hq": sum(1 for r in rows if not is_country_only(r.get("Headquarters") or "")),
        "sample": [
            {
                "company": r["Company"],
                "founded": r["Founded"],
                "hq": r["Headquarters"],
                "cities": r["Cities / Regions"],
            }
            for r in rows[:15]
        ],
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nDONE rows={len(rows)} city-level HQ={summary['city_level_hq']} country-only={country_left}")
    print(f"Wrote {XLSX}")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
