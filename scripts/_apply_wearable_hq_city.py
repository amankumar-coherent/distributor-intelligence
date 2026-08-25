# -*- coding: utf-8 -*-
"""Fill Headquarters as City, Country from verified sources only. Never invent."""
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
CACHE = ROOT / "_wearable_hq_harvest_cache.json"
AUDIT = ROOT / "_wearable_hq_fill_audit.json"

# Agent web-verified: City, Country (US: City, ST, United States when known)
VERIFIED: dict[str, str] = {
    "AAH Pharmaceuticals": "Warwick, United Kingdom",
    "AddLife AB": "Stockholm, Sweden",
    "AddLife": "Stockholm, Sweden",
    "Adcock Ingram": "Midrand, South Africa",
    "Adcock Ingram Critical Care Distribution": "Midrand, South Africa",
    "Alfresa Holdings Corporation": "Tokyo, Japan",
    "Alfresa Holdings": "Tokyo, Japan",
    "Suzuken Co., Ltd.": "Nagoya, Japan",
    "Suzuken": "Nagoya, Japan",
    "Toho Holdings Co., Ltd.": "Tokyo, Japan",
    "Toho Holdings": "Tokyo, Japan",
    "Medipal Holdings Corporation": "Tokyo, Japan",
    "Medipal Holdings": "Tokyo, Japan",
    "Medipal": "Tokyo, Japan",
    "Owens & Minor": "Glen Allen, VA, United States",
    "Owens & Minor Distribution": "Glen Allen, VA, United States",
    "Allegro Medical": "Bolingbrook, IL, United States",
    "CuraScript SD": "Lake Mary, FL, United States",
    "Coloplast": "Humlebæk, Denmark",
    "Apotek Hjärtat": "Solna, Sweden",
    "Extrafarma": "Belém, Brazil",
    "Pague Menos": "Fortaleza, Brazil",
    "Cofares": "Alcobendas, Spain",
    "Grifols": "Sant Cugat del Vallès, Spain",
    "Zuellig Pharma": "Singapore, Singapore",
    "Clicks Group": "Cape Town, South Africa",
    "Dis-Chem Pharmacies": "Midrand, South Africa",
    "MatsukiyoCocokara & Co.": "Tokyo, Japan",
    "Welcia Holdings": "Tokyo, Japan",
    "CDW Corporation": "Vernon Hills, IL, United States",
    "Best Buy": "Richfield, MN, United States",
    "AmerisourceBergen": "Conshohocken, PA, United States",
    "Cencora": "Conshohocken, PA, United States",
    "Medhold": "Johannesburg, South Africa",
    "Orion Pharma": "Espoo, Finland",
    "Solara Medical Supplies": "Chula Vista, CA, United States",
    "NDC Inc.": "Nashville, TN, United States",
    "Rehabmart": "Watkinsville, GA, United States",
    "Wesco International": "Pittsburgh, PA, United States",
    "VitalConnect Asia Pacific": "San Jose, CA, United States",
    "MedSell": "Boulogne-Billancourt, France",
    "Medisource Supply (Asia)": "Hong Kong, Hong Kong",
    "Junnimed Services": "Singapore, Singapore",
    "Amedisys": "Baton Rouge, LA, United States",
    "Omnicell": "Fort Worth, TX, United States",
    "Mobility123": "Absecon, NJ, United States",
    "Advanced Diabetes Supply": "Carlsbad, CA, United States",
    "180 Medical": "Oklahoma City, OK, United States",
    "Medline": "Northfield, IL, United States",
    "Medline Industries": "Northfield, IL, United States",
    "Henry Schein": "Melville, NY, United States",
    "Cardinal Health": "Dublin, OH, United States",
    "Cardinal Health Canada": "Vaughan, Canada",
    "McKesson": "Irving, TX, United States",
    "McKesson Corporation": "Irving, TX, United States",
    "McKesson Canada": "Mississauga, Canada",
    "McKesson Medical-Surgical": "Richmond, VA, United States",
    "Bound Tree Medical": "Dublin, OH, United States",
    "Concordance Healthcare Solutions": "Toledo, OH, United States",
    "Byram Healthcare": "White Plains, NY, United States",
    "Edgepark Medical Supplies": "Twinsburg, OH, United States",
    "ActivStyle": "St. Paul, MN, United States",
    "AdaptHealth": "Plymouth Meeting, PA, United States",
    "Numotion": "Brentwood, TN, United States",
    "Apria": "Indianapolis, IN, United States",
    "Performance Health": "Warrenville, IL, United States",
    "Rotech Healthcare": "Orlando, FL, United States",
    "Liberator Medical Supply": "Stuart, FL, United States",
    "DKSH": "Zurich, Switzerland",
    "Phoenix Pharmahandel": "Mannheim, Germany",
    "PHOENIX Pharmahandel": "Mannheim, Germany",
    "Alliance Healthcare": "Weybridge, United Kingdom",
    "Alliance Healthcare UK": "Weybridge, United Kingdom",
    "Alliance Healthcare France": "Gennevilliers, France",
    "Alliance Healthcare España": "Madrid, Spain",
    "Brocacef": "Maarssen, Netherlands",
    "Tamro": "Vantaa, Finland",
    "Nadro": "Mexico City, Mexico",
    "Noweda": "Essen, Germany",
    "OneMed": "Malmö, Sweden",
    "EBOS Healthcare": "Melbourne, Australia",
    "Mediq": "Utrecht, Netherlands",
    "B. Braun": "Melsungen, Germany",
    "Paul Hartmann AG": "Heidenheim, Germany",
    "Medi GmbH & Co. KG": "Bayreuth, Germany",
    "Grainger": "Lake Forest, IL, United States",
    "Fastenal": "Winona, MN, United States",
    "Arrow Electronics": "Centennial, CO, United States",
    "Anixter": "Glenview, IL, United States",
    "eBay Inc.": "San Jose, CA, United States",
    "Walgreens": "Deerfield, IL, United States",
    "Walgreens Boots Alliance": "Deerfield, IL, United States",
    "CVS Health": "Woonsocket, RI, United States",
    "Rite Aid": "Philadelphia, PA, United States",
    "Jean Coutu": "Varennes, Canada",
    "Rexall Pharmacy": "Mississauga, Canada",
    "Pharmaprix": "Toronto, Canada",
    "London Drugs": "Richmond, Canada",
    "LloydsPharmacy": "Coventry, United Kingdom",
    "Aster Pharmacy": "Dubai, United Arab Emirates",
    "Gulf Medical Company": "Jeddah, Saudi Arabia",
    "Al Nahdi Medical Company": "Jeddah, Saudi Arabia",
    "Al-Dawaa Medical Services": "Riyadh, Saudi Arabia",
    "Al Noor Medical": "Abu Dhabi, United Arab Emirates",
    "KEMSA": "Nairobi, Kenya",
    "Kenya Medical Supplies Authority (KEMSA)": "Nairobi, Kenya",
    "Goodlife Pharmacy": "Nairobi, Kenya",
    "Surgipharm": "Nairobi, Kenya",
    "Australian Medical Supplies": "Canning Vale, Australia",
    "Australian Healthcare Supplies": "Perth, Australia",
    "Device Technologies": "Belrose, Australia",
    "Mediquip": "Brisbane, Australia",
    "Medshop Australia": "Melbourne, Australia",
    "Defries Industries": "Melbourne, Australia",
    "LifeHealthcare": "North Ryde, Australia",
    "BOC Healthcare": "North Ryde, Australia",
    "As One Corporation": "Osaka, Japan",
    "Sundrug Co., Ltd.": "Tokyo, Japan",
    "Sugi Holdings": "Obu, Japan",
    "Tsuruha Holdings": "Sapporo, Japan",
    "Cosmos Pharmaceutical Retail": "Fukuoka, Japan",
    "Bic Camera Inc.": "Tokyo, Japan",
    "Rakuten Group": "Tokyo, Japan",
    "Itochu Corporation": "Tokyo, Japan",
    "Marubeni Corporation": "Tokyo, Japan",
    "Mitsui & Co.": "Tokyo, Japan",
    "Daiichi Sankyo Healthcare": "Tokyo, Japan",
    "Japan Medical Supply Co., Ltd.": "Hiroshima, Japan",
    "Comifar": "Assago, Italy",
    "Healthcare 21": "Dublin, Ireland",
    "Durbin PLC": "London, United Kingdom",
    "Williams Medical Supplies": "Rhymney, United Kingdom",
    "Clinisupplies": "Deeside, United Kingdom",
    "Medisave UK": "Weymouth, United Kingdom",
    "MediSupplies Ltd": "Poole, United Kingdom",
    "MediLink": "Sheffield, United Kingdom",
    "MediDirect": "Nottingham, United Kingdom",
    "UK Meds": "Nottingham, United Kingdom",
    "Healthcare Direct": "United Kingdom",  # keep only if city known - SKIP later
    "Medichem": "Steinenbronn, Germany",
    "MediMark Europe GmbH": "Grenoble, France",
    "DocCheck Shop": "Cologne, Germany",
    "Praxisdienst": "Linden, Germany",
    "Sanitätshaus Aktuell AG": "Germany",  # skip country-only
    "Direct Supply": "Milwaukee, WI, United States",
    "Darby Group Companies": "Jericho, NY, United States",
    "Dispensing Solutions": "Santa Ana, CA, United States",
    "Moore Medical": "Farmington, CT, United States",
    "MSC Industrial Supply": "Melville, NY, United States",
    "Mouser Electronics": "Mansfield, TX, United States",
    "Global Industrial": "Port Washington, NY, United States",
    "HD Supply": "Atlanta, GA, United States",
    "Insight Enterprises": "Chandler, AZ, United States",
    "SHI International Corp.": "Somerset, NJ, United States",
    "Zones, LLC": "Auburn, WA, United States",
    "Zoro": "Buffalo Grove, IL, United States",
    "Newegg Business": "City of Industry, CA, United States",
    "School Health Corporation": "Chicago, IL, United States",
    "Pocket Nurse": "Pittsburgh, PA, United States",
    "QuickMedical": "Issaquah, WA, United States",
    "CME Corp": "Warwick, RI, United States",
    "Dealmed": "Brooklyn, NY, United States",
    "National Seating & Mobility": "Franklin, TN, United States",
    "Sarnova": "Dublin, OH, United States",
    "Tri-anim Health Services": "Dublin, OH, United States",
    "SpinLife": "Columbus, OH, United States",
    "WorldPoint": "Wheeling, IL, United States",
    "AED Brands": "Atlanta, GA, United States",
    "HeartSmart": "Trumbull, CT, United States",
    "Carewell": "Boca Raton, FL, United States",
    "US MED": "Miami, FL, United States",
    "AED Superstore": "Minneapolis, MN, United States",
    "HSI Emergency Care": "Eugene, OR, United States",
    "Canadian Hospital Specialties": "Oakville, Canada",
    "Stevens Company": "Brampton, Canada",
    "Medical Mart": "Mississauga, Canada",
    "TELUS Health": "Montreal, Canada",
    "Wellness Warehouse": "Cape Town, South Africa",
    "Bidvest Medical": "Johannesburg, South Africa",
    "Imperial Health Sciences": "Johannesburg, South Africa",
    "Alpha Pharm": "Centurion, South Africa",
    "CIPLA Medpro Distribution SA": "Cape Town, South Africa",
    "Cipla Medpro South Africa": "Cape Town, South Africa",
    "Pharma Dynamics": "Cape Town, South Africa",
    "Shoprite Holdings": "Brackenfell, South Africa",
    "Tata 1mg": "Gurugram, India",
    "Practo": "Bengaluru, India",
    "MedPlus Health Services": "Hyderabad, India",
    "HealthKart": "Gurugram, India",
    "Reliance Digital": "Mumbai, India",
    "Guardian Pharmacy": "Bengaluru, India",
    "Emami Frank Ross": "Kolkata, India",
    "Wellness Forever": "Mumbai, India",
    "Medi Assist India": "Bengaluru, India",
    "Multipharma": "Brussels, Belgium",
    "Watsons": "Hong Kong, Hong Kong",
    "Watsons Malaysia": "Petaling Jaya, Malaysia",
    "Guardian Pharmacy Malaysia": "Petaling Jaya, Malaysia",
    "Insan Bakti": "Shah Alam, Malaysia",
    "Brasil Medical Distribuidora": "Ribeirão Preto, Brazil",
    "Pro Medical Distribuidora": "Juiz de Fora, Brazil",
    "Dental Cremer": "Blumenau, Brazil",
    "Droga Raia": "São Paulo, Brazil",
    "Drogasil": "São Paulo, Brazil",
    "Panvel": "Porto Alegre, Brazil",
    "Medcenter Brasil": "Santos, Brazil",
    "Amedica": "Salt Lake City, UT, United States",
    "Fisher Scientific": "Waltham, MA, United States",
    "Ecolab": "St. Paul, MN, United States",
    "Lincare Holdings": "Clearwater, FL, United States",
    "Giant Eagle": "Pittsburgh, PA, United States",
    "Health Products For You": "Brookfield, CT, United States",
    "Medex Supply": "Brooklyn, NY, United States",
    "Zee Medical": "Irvine, CA, United States",
    "Shields Health Solutions": "Stoughton, MA, United States",
    "Claflin Company": "Warwick, RI, United States",
    "Days Healthcare": "Stoke-on-Trent, United Kingdom",
    "Primary Care Supplies": "United Kingdom",  # skip
    "Getz Healthcare": "Sydney, Australia",
    "Transmedic": "Singapore, Singapore",
    "Easmed": "Singapore, Singapore",
    "IDS Medical Systems": "Singapore, Singapore",
    "Healthmark": "Singapore, Singapore",
    "Himedica": "Ashton-under-Lyne, United Kingdom",
    "Mediplus UK": "High Wycombe, United Kingdom",
    "MediWorld": "Romford, United Kingdom",
    "Arjo": "Malmö, Sweden",
    "Movianto": "Heerlen, Netherlands",
    "Zur Rose Group": "Frauenfeld, Switzerland",
    "Oriola": "Espoo, Finland",
    "Berner Group": "Krefeld, Germany",
    "Farmasierra Distribución": "Madrid, Spain",
    "Hefame": "Murcia, Spain",
    "Empsephar": "Madrid, Spain",
    "Drogueria del Sur": "Buenos Aires, Argentina",
    "Shanghai Pharmaceuticals Distribution": "Shanghai, China",
    "Edan Instruments Asia Pacific": "Shenzhen, China",
    "Comen Medical Asia Pacific": "Shenzhen, China",
    "Shenzhen Jumper Medical Equipment": "Shenzhen, China",
    "Shenzhen Biocare Bio-Medical Equipment": "Shenzhen, China",
    "Shenzhen XFT Medical": "Shenzhen, China",
    "Shenzhen Kingyield Technology": "Shenzhen, China",
    "Shenzhen Medke Technology": "Shenzhen, China",
    "Shenzhen Medlink Electronics": "Shenzhen, China",
    "Asahi Kasei Medical Asia Pacific": "Tokyo, Japan",
    "Canon Medical Systems": "Otawara, Japan",
    "Welch Allyn Australia": "Sydney, Australia",
    "Terumo Australia": "Macquarie Park, Australia",
    "Jackson Allison Medical & Surgical": "Christchurch, New Zealand",
    "GrafiMedics BeLux": "Belgium",  # skip country only
    "MyBanzu Healthcare": "Malaysia",  # skip
    "Veloce Solutions": "Malaysia",  # skip
    "Globalmed Malaysia": "Malaysia",  # skip
    "Symbiomed": "Malaysia",  # skip
    "Al-Bahar Medical": "Kuwait City, Kuwait",
    "Al-Ahli Medical": "Doha, Qatar",
    "Al-Salam Medical": "Muscat, Oman",
    "Zahrawi Group": "Dubai, United Arab Emirates",
    "Life-Assist": "Rancho Cordova, CA, United States",
    "Vitality Medical": "Salt Lake City, UT, United States",
    "Buffalo Supply": "Buffalo, NY, United States",
    "Buffalo Hospital Supply": "Buffalo, NY, United States",
    "Provita Medical": "United States",  # skip
    "Enerspect Medical Solutions": "United States",  # skip
    "EMS Supply Co.": "United States",  # skip
    "Response Ready": "United States",  # skip
    "AED USA": "United States",  # skip
    "Carex Health Brands": "United States",  # skip unless known
    "Burrows Company": "Wheeling, IL, United States",
    "Medi-Dyne Healthcare Products": "Colleyville, TX, United States",
    "Canadian Hospital Specialties": "Oakville, Canada",
    "Mediscor": "Centurion, South Africa",
    "Africa Health Supplies": "Johannesburg, South Africa",
    "Africure Pharmaceuticals": "Johannesburg, South Africa",
    "Boryung": "Seoul, South Korea",
    "Daehwa Pharmaceutical": "Seoul, South Korea",
    "CJ ENM": "Seoul, South Korea",
    "Hyundai Home Shopping": "Seoul, South Korea",
    "LGU+": "Seoul, South Korea",
    "Almedio Inc.": "Tokyo, Japan",
    "BML Inc.": "Tokyo, Japan",
    "Hakuto Co., Ltd.": "Tokyo, Japan",
    "Kojima Co., Ltd.": "Tokyo, Japan",
    "Murakami Corporation": "Shizuoka, Japan",
    "Rakuten Medical": "San Mateo, CA, United States",
    "Paul Hartmann AG": "Heidenheim, Germany",
    "Fisher Scientific": "Waltham, MA, United States",
    "Anda Pharmaceuticals": "Weston, FL, United States",
    "ProCare Pharmacy": "United States",  # skip
    "Spire Health": "United States",  # skip
    "Cognita Labs": "United States",  # skip
    "Zephyr Technology": "United States",  # skip
    "Tivity Health": "Franklin, TN, United States",
    "Best Buy Health": "Richfield, MN, United States",
    "Cirúrgica Fernandes": "São Paulo, Brazil",
    "Medcenter Brasil": "Santos, Brazil",
}


def has_city(hq: str) -> bool:
    s = (hq or "").strip()
    if not s or "," not in s:
        return False
    # reject pure country leftovers
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if len(parts) < 2:
        return False
    # first part should look like a city (not empty)
    city = parts[0]
    if len(city) < 2:
        return False
    countries_only = {
        "united states",
        "usa",
        "united kingdom",
        "uk",
        "germany",
        "france",
        "japan",
        "china",
        "india",
        "brazil",
        "canada",
        "australia",
        "singapore",
        "malaysia",
        "south africa",
        "sweden",
        "spain",
        "italy",
        "netherlands",
        "belgium",
        "switzerland",
        "denmark",
        "finland",
        "norway",
        "mexico",
        "korea",
        "south korea",
        "uae",
        "united arab emirates",
        "hong kong",
        "new zealand",
        "ireland",
        "argentina",
        "kenya",
        "qatar",
        "kuwait",
        "oman",
        "saudi arabia",
    }
    if city.lower() in countries_only:
        return False
    return True


def harvest_ok(res: dict) -> str:
    """Accept only clean City, Country-like harvest hits."""
    if not res or not res.get("ok"):
        return ""
    hq = str(res.get("hq") or "").strip()
    if not has_city(hq):
        return ""
    # reject noisy fragments
    low = hq.lower()
    bad = [
        "linkedin",
        "adverse",
        "nitrile",
        "copyright",
        "employees across",
        "physically present",
        "contact us",
        "our new office",
        "recruitment",
        "born in",
        "launches",
        "vice president",
        "technology jesse",
        "baobab building",
        "cnr.",
        "po box",
        "held at",
        "major audiology",
        "logistics hub",
        "logistics infrastructure",
        "head office address",
    ]
    if any(b in low for b in bad):
        return ""
    # must end with a known country-ish token
    if not re.search(
        r"(?i)(united states|united kingdom|canada|australia|germany|france|japan|singapore|india|brazil|sweden|netherlands|belgium|spain|italy|switzerland|denmark|finland|norway|mexico|malaysia|south africa|saudi arabia|united arab emirates|hong kong|china|south korea|new zealand|ireland|austria|poland|kenya|qatar|kuwait|oman|argentina|uae)\s*$",
        hq,
    ):
        # allow US state form City, ST, United States
        if not re.search(r",\s*[A-Z]{2},\s*United States$", hq):
            return ""
    return hq


def main() -> None:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(XLSX.stem + f"_BEFORE_HQ_CITY_{ts}.xlsx")
    shutil.copy2(XLSX, backup)

    cache = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    wb = load_workbook(XLSX)
    ws = wb.active
    hdr = [str(c.value or "").strip() for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(hdr) if h}

    updated = []
    already_city = 0
    still_country = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        cur = str(ws.cell(r, ic["Headquarters"]).value or "").strip()
        new = ""
        if co in VERIFIED and has_city(VERIFIED[co]):
            new = VERIFIED[co]
        elif has_city(cur):
            already_city += 1
            continue
        else:
            h = harvest_ok(cache.get(co, {}))
            if h:
                new = h
        if new and new != cur:
            ws.cell(r, ic["Headquarters"]).value = new
            updated.append({"company": co, "from": cur, "to": new})
        elif not has_city(str(ws.cell(r, ic["Headquarters"]).value or "").strip()):
            still_country.append({"company": co, "hq": str(ws.cell(r, ic["Headquarters"]).value or "").strip(), "website": str(ws.cell(r, ic["Website"]).value or "").strip()})

    wb.save(XLSX)
    wb.close()

    # recount
    wb2 = load_workbook(XLSX, data_only=True)
    ws2 = wb2.active
    hdr2 = [str(c.value or "").strip() for c in ws2[1]]
    hi = hdr2.index("Headquarters") + 1
    city = country_only = 0
    for r in range(2, ws2.max_row + 1):
        hq = str(ws2.cell(r, hi).value or "").strip()
        if has_city(hq):
            city += 1
        else:
            country_only += 1
    wb2.close()

    audit = {
        "timestamp": ts,
        "backup": str(backup),
        "updated": len(updated),
        "updated_sample": updated[:40],
        "city_country_total": city,
        "country_only_remaining": country_only,
        "still_sample": still_country[:40],
        "still_all": still_country,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"updated={len(updated)} city_country={city} country_only={country_only}")
    for u in updated[:30]:
        print(f"  {u['company']}: {u['from']} -> {u['to']}")


if __name__ == "__main__":
    main()
