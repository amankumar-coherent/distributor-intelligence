#!/usr/bin/env python3
"""Add new, unique companies from a pasted research table to the FINAL sheet.

Every candidate was checked against the existing 185 rows by company name
AND by website domain before being included here -- true duplicates
(AmerisourceBergen/Cencora, Cardinal Health, Edgepark Medical Supplies,
AdaptHealth, Advanced Diabetes Supply, McKesson Corporation, Byram
Healthcare, Surgipharm, AddLife AB->AddLife, OneMed AB->OneMed, Zee Medical,
NDC Inc., Bound Tree Medical) were excluded, along with internal duplicates
within the pasted tables themselves (kept the more complete of the two
versions where a company appeared twice).
"""
from pathlib import Path
import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "output/web_expand/global_wearable_medical_devices_market_global/global_wearable_medical_devices_market_global_FINAL.xlsx"
SHEET = "Wearable Medical Companies"

COLUMNS = [
    "Company", "Website", "Founded", "Headquarters", "Continent / Geography",
    "Operational Presence", "Ownership", "Employees", "Core Categories",
    "Specialty Focus", "Key Brands Represented", "Online / Offline / Both",
    "Contact Person", "Role", "Email", "LinkedIn", "Office No.",
    "Office Email", "Country Code", "Region Code", "Summary",
]

NEW_ROWS = [
    dict(Company="Smith Drug Company", Website="https://www.smithdrug.com", Founded=1944,
         Headquarters="Spartanburg, SC, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (Southeast)"}, Ownership="Private (JM Smith)", Employees="unspecified",
         **{"Core Categories": "Pharmaceuticals; Home Medical Equipment; others"},
         **{"Specialty Focus": "Independent Pharmacy Wholesale; others"},
         **{"Key Brands Represented": "Dexcom; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/smith-drug-company", **{"Office No.": "(864) 583-3100"},
         **{"Office Email": "info@smithdrug.com"}, **{"Country Code": "+1"}, **{"Region Code": "864"},
         Summary="Smith Drug Co. is a private pharmaceutical wholesaler founded in 1944, headquartered in Spartanburg, SC. It supplies independent pharmacies and HME providers across the southeastern United States and is an authorized Dexcom CGM distributor."),

    dict(Company="Independence Medical", Website="https://www.indemed.com", Founded=1990,
         Headquarters="Twinsburg, OH, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (DME/Pharmacies)"}, Ownership="Private (Edgepark/Cardinal)", Employees="unspecified",
         **{"Core Categories": "Disposable Medical Supplies; Home Medical Equipment; others"},
         **{"Specialty Focus": "DME, Pharmacy Distribution; others"},
         **{"Key Brands Represented": "Dexcom; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="customerservice@edgepark.com",
         LinkedIn="Not publicly disclosed", **{"Office No.": "(800) 321-0591"},
         **{"Office Email": "customerservice@edgepark.com"}, **{"Country Code": "+1"}, **{"Region Code": "800"},
         Summary="Independence Medical (RGH Enterprises) is a US wholesale distributor founded in 1990 in Twinsburg, OH, providing disposable medical supplies to DMEs, pharmacies, and homecare providers. It is an authorized Dexcom CGM distributor."),

    dict(Company="Premier Kids Care", Website="https://www.premierkidscare.com", Founded=1996,
         Headquarters="Monroe, GA, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (40+ states)"}, Ownership="Private", Employees="51-200",
         **{"Core Categories": "Pediatric Pharmacy; Home Medical Equipment; others"},
         **{"Specialty Focus": "Pediatric Diabetes Care; others"},
         **{"Key Brands Represented": "Dexcom; FreeStyle Libre; Tandem; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Joy McMillan"}, Role="President & CEO", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/premier-kids-care", **{"Office No.": "(770) 207-9001"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+1"}, **{"Region Code": "770"},
         Summary="Premier Kids Care is a US pediatric pharmacy and DME distributor founded in 1996, headquartered in Monroe, GA. It specializes in diabetes supplies and insulin delivery systems for children nationwide and is an authorized Dexcom CGM distributor."),

    dict(Company="Nazareno Enterprises & Services", Website="https://www.nazarenoservices.com", Founded="Not publicly disclosed",
         Headquarters="Bayamón, Puerto Rico, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "Puerto Rico & Caribbean"}, Ownership="Private", Employees="1-50",
         **{"Core Categories": "Diabetes Home Medical Equipment; Rehabilitation Equipment; others"},
         **{"Specialty Focus": "Diabetes Care, Rehab Services; others"},
         **{"Key Brands Represented": "Dexcom; Tandem; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="rx@nazarenollc.com",
         LinkedIn="https://www.linkedin.com/company/nazareno-services", **{"Office No.": "(787) 740-2934"},
         **{"Office Email": "rx@nazarenollc.com"}, **{"Country Code": "+1"}, **{"Region Code": "787"},
         Summary="Nazareno Enterprises & Services is a Puerto Rico-based medical distributor specializing in diabetes care. It distributes Dexcom CGMs and Tandem insulin pumps throughout Puerto Rico and the Caribbean."),

    dict(Company="DDP Medical Supply", Website="https://www.ddpmedical.com", Founded=1998,
         Headquarters="St. Petersburg, FL, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (nationwide)"}, Ownership="Private", Employees="51-200",
         **{"Core Categories": "Durable Medical Equipment; Diagnostic Wearables; others"},
         **{"Specialty Focus": "Diabetes Monitoring; others"},
         **{"Key Brands Represented": "Dexcom; FreeStyle Libre; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Joe DiFabio"}, Role="CEO", Email="cs@ddpmedical.com",
         LinkedIn="https://www.linkedin.com/company/ddp-medical", **{"Office No.": "(800) 437-4455"},
         **{"Office Email": "cs@ddpmedical.com"}, **{"Country Code": "+1"}, **{"Region Code": "727"},
         Summary="DDP Medical Supply is a US diabetes DME distributor founded in 1998, headquartered in St. Petersburg, FL. It specializes in CGM and insulin pump supplies and is an authorized distributor for both Abbott FreeStyle Libre and Dexcom."),

    dict(Company="GEMCO Medical", Website="https://www.gemcomedical.com", Founded=1984,
         Headquarters="Hudson, OH, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (nationwide)"}, Ownership="Private", Employees="51-200",
         **{"Core Categories": "Medical Devices; Diabetes Supplies; others"},
         **{"Specialty Focus": "Diabetes Care DME; others"},
         **{"Key Brands Represented": "Dexcom; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="info@gemco-med.com",
         LinkedIn="https://www.linkedin.com/company/gemcomedical", **{"Office No.": "(800) 733-7976"},
         **{"Office Email": "info@gemco-med.com"}, **{"Country Code": "+1"}, **{"Region Code": "330"},
         Summary="GEMCO Medical is an Ohio-based medical and diabetes supply distributor founded in 1984. It provides wholesale and 3PL services nationwide with a focus on diabetes products and is an authorized Dexcom CGM distributor."),

    dict(Company="Alpine Health", Website="https://www.alpinehealth.com", Founded=1994,
         Headquarters="Puerto Rico, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA / Puerto Rico"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Diabetes Supplies; Pharmacy; others"},
         **{"Specialty Focus": "Diabetes Management; others"},
         **{"Key Brands Represented": "FreeStyle Libre; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/alpinehealth-pr", **{"Office No.": "(800) 880-8805"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+1"}, **{"Region Code": "939"},
         Summary="Alpine Health is a Caribbean/US diabetes distributor founded in 1994, serving Puerto Rico. It is a retail authorized distributor of Abbott FreeStyle Libre CGMs and supplies various diabetes monitoring devices."),

    dict(Company="Dakota Drug, Inc.", Website="https://www.dakotadrug.com", Founded=1941,
         Headquarters="Rapid City, SD, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (Upper Midwest)"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Pharmaceuticals; Retail Pharmacy; others"},
         **{"Specialty Focus": "Community Pharmacy Distribution; others"},
         **{"Key Brands Represented": "FreeStyle Libre; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/dakota-drug-inc", **{"Office No.": "(605) 716-2200"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+1"}, **{"Region Code": "605"},
         Summary="Dakota Drug, Inc. is a regional pharmacy wholesaler founded in 1941, serving the Upper Midwest US. It is listed as an Abbott FreeStyle Libre CGM authorized distributor for retail pharmacies."),

    dict(Company="Morris & Dickson Co.", Website="https://www.morris-dickson.com", Founded=1966,
         Headquarters="Shreveport, LA, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (South)"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Pharmaceuticals; Retail Pharmacy; others"},
         **{"Specialty Focus": "Community Pharmacy Distribution; others"},
         **{"Key Brands Represented": "FreeStyle Libre; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/morris--dickson-co", **{"Office No.": "(800) 658-9989"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+1"}, **{"Region Code": "318"},
         Summary="Morris & Dickson Co. is a Southern US drug wholesaler founded in 1966, headquartered in Shreveport, LA. It is named on Abbott's US authorized retail distributor list, indicating CGM device distribution."),

    dict(Company="NC Mutual Drug", Website="https://www.ncmutualdrug.com", Founded=1901,
         Headquarters="Charlotte, NC, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (Mid-Atlantic)"}, Ownership="Cooperative", Employees="unspecified",
         **{"Core Categories": "Pharmaceuticals; Retail Pharmacy; others"},
         **{"Specialty Focus": "Community Pharmacy Cooperative Distribution; others"},
         **{"Key Brands Represented": "FreeStyle Libre; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/nc-mutual-drug-company", **{"Office No.": "(800) 638-3176"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+1"}, **{"Region Code": "704"},
         Summary="NC Mutual Drug is a pharmacy cooperative wholesaler founded in 1901, headquartered in Charlotte, NC. It is listed as an Abbott FreeStyle Libre CGM authorized distributor, supplying member pharmacies."),

    dict(Company="Value Drug", Website="https://www.valuedrugco.com", Founded="Not publicly disclosed",
         Headquarters="Muskogee, OK, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (OK, AR, KS)"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Retail Pharmacy; others"},
         **{"Specialty Focus": "Community Pharmacy Retail; others"},
         **{"Key Brands Represented": "FreeStyle Libre; others"}, **{"Online / Offline / Both": "Offline"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/value-drug-co", **{"Office No.": "Not publicly disclosed"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+1"}, **{"Region Code": "918"},
         Summary="Value Drug is a regional pharmacy retailer in Oklahoma, Arkansas, and Kansas. It is named as a US Abbott FreeStyle Libre CGM authorized retail distributor."),

    dict(Company="GoodHealth Medtech", Website="https://www.ghmt.in", Founded=2019,
         Headquarters="Mumbai, India", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "India"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Cardiac Monitoring; others"},
         **{"Key Brands Represented": "BioCalculus; iRhythm; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/goodhealth-medtech-pvt-ltd", **{"Office No.": "+91-22-26113084"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+91"}, **{"Region Code": "22"},
         Summary="GoodHealth Medtech is an India-based medical device distributor founded in 2019, headquartered in Mumbai. It is an authorized distributor of BioCalculus wireless wearable ECG monitors, specializing in cardiology monitoring."),

    dict(Company="Vitalom Healthcare", Website="https://www.vitalom.in", Founded=2018,
         Headquarters="Mumbai, India", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "India"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Medical Devices; Diagnostic Wearables; others"},
         **{"Specialty Focus": "Cardiac and Patient Monitoring; others"},
         **{"Key Brands Represented": "Bionet; ViewCare; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/vitalom-healthcare", **{"Office No.": "+91-22-40825870"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+91"}, **{"Region Code": "22"},
         Summary="Vitalom Healthcare is an India medical distributor founded in 2018, headquartered in Mumbai. It offers cardiac monitors, patient monitors, and ICU equipment as a channel partner for brands like Bionet."),

    dict(Company="Air Liquide Healthcare (UK)", Website="https://www.airliquidehealthcare.co.uk", Founded=1997,
         Headquarters="Farnborough, United Kingdom", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "UK & Ireland"}, Ownership="Private (Air Liquide)", Employees="unspecified",
         **{"Core Categories": "Diagnostic Wearables; Respiratory Equipment; others"},
         **{"Specialty Focus": "Diabetes Monitoring; others"},
         **{"Key Brands Represented": "Sibionics; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/air-liquide-healthcare", **{"Office No.": "+44 1252 556888"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+44"}, **{"Region Code": "1252"},
         Summary="Air Liquide Healthcare (UK) is the UK division of Air Liquide, founded in 1997, headquartered in Farnborough. It is the exclusive UK distributor for Sibionics CGM systems and other diabetes/respiratory devices."),

    dict(Company="Urathon", Website="https://www.urathon.com", Founded=2011,
         Headquarters="Cambridgeshire, United Kingdom", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "UK, Ireland, Middle East"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Diagnostic Wearables; Wound Care; others"},
         **{"Specialty Focus": "Diabetes Monitoring; others"},
         **{"Key Brands Represented": "Yuwell; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/urathon", **{"Office No.": "+44 1932 823483"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+44"}, **{"Region Code": "1932"},
         Summary="Urathon (formerly AGS UK) is a UK medical distributor founded in 2011, headquartered in Cambridgeshire. It is the exclusive UK partner for Yuwell's Anytime CGM diabetes monitoring system and also sells wound care and pharmacy products."),

    dict(Company="i.SigNal", Website="https://www.signalinc.co.kr", Founded=2016,
         Headquarters="Seoul, South Korea", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "Asia-Pacific (23+ countries)"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Cardiac Monitoring; others"},
         **{"Key Brands Represented": "Cardiac Insight; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/i-signal-corp", **{"Office No.": "+82-2-539-1117"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+82"}, **{"Region Code": "2"},
         Summary="i.SigNal (by Dreamtech) is an APAC medical distributor founded in 2016, headquartered in Seoul. It is the master distributor for Cardiac Insight's Cardea SOLO wearable ECG patch across 23 Asia-Pacific countries."),

    dict(Company="Dreamtech", Website="https://www.dreamtech.co.kr", Founded=2001,
         Headquarters="Seoul, South Korea", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "Asia-Pacific"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Medical Devices; Diagnostic Wearables; others"},
         **{"Specialty Focus": "Cardiac and Diabetes Monitoring; others"},
         **{"Key Brands Represented": "Cardiac Insight; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="Not publicly disclosed",
         LinkedIn="https://www.linkedin.com/company/dreamtechinc", **{"Office No.": "+82-2-2108-0412"},
         **{"Office Email": "Not publicly disclosed"}, **{"Country Code": "+82"}, **{"Region Code": "2"},
         Summary="Dreamtech is a Korean medical technology distributor founded in 2001, headquartered in Seoul (i.SigNal operates as its unit). It holds distribution rights for the Cardea SOLO wearable ECG and other monitoring devices across Asia."),

    dict(Company="CCS Medical", Website="https://www.ccsmed.com", Founded=2005,
         Headquarters="St. Petersburg, FL, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (nationwide)"}, Ownership="Private", Employees="~600",
         **{"Core Categories": "Diagnostic Wearables; Home Medical Equipment; others"},
         **{"Specialty Focus": "Chronic Disease Management; others"},
         **{"Key Brands Represented": "Dexcom; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="service@ccsmed.com",
         LinkedIn="Not publicly disclosed", **{"Office No.": "(888) 224-7782"},
         **{"Office Email": "service@ccsmed.com"}, **{"Country Code": "+1"}, **{"Region Code": "727"},
         Summary="CCS Medical is a Florida-based medical supply company founded in 2005, specializing in diabetes care supplies. It is a major DME provider for Medicare patients and an authorized Dexcom CGM distributor."),

    dict(Company="Edwards Health Care Services", Website="https://www.myehcs.com", Founded=1992,
         Headquarters="Hudson, OH, United States", **{"Continent / Geography": "North America"},
         **{"Operational Presence": "USA (nationwide)"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Diagnostic Wearables; Home Medical Equipment; others"},
         **{"Specialty Focus": "Diabetes Care; others"},
         **{"Key Brands Represented": "Dexcom; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="info@myehcs.com",
         LinkedIn="Not publicly disclosed", **{"Office No.": "(888) 344-3434"},
         **{"Office Email": "info@myehcs.com"}, **{"Country Code": "+1"}, **{"Region Code": "330"},
         Summary="Edwards Health Care Services is a family-owned DME distributor founded in 1992, based in Hudson, OH. It delivers diabetes devices and supplies to patients' homes and is an authorized Dexcom CGM distributor."),

    dict(Company="VitalAire Diabetes", Website="https://www.vitalairediabetes.co.za", Founded="Not publicly disclosed",
         Headquarters="Midrand, South Africa", **{"Continent / Geography": "Africa"},
         **{"Operational Presence": "South Africa"}, Ownership="Private (Air Liquide)", Employees="unspecified",
         **{"Core Categories": "Diagnostic Wearables; Medical Gases; others"},
         **{"Specialty Focus": "Diabetes Care, Home Oxygen; others"},
         **{"Key Brands Represented": "Dexcom; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="contact@vitalairediabetes.co.za",
         LinkedIn="Not publicly disclosed", **{"Office No.": "(0861) 339-266"},
         **{"Office Email": "contact@vitalairediabetes.co.za"}, **{"Country Code": "+27"}, **{"Region Code": "861"},
         Summary="VitalAire Diabetes is a South African medical gas and diabetes-care distributor headquartered in Midrand, part of Air Liquide. It is the exclusive distributor of Dexcom CGM systems in South Africa."),

    dict(Company="GluX", Website="https://www.gluxshop.hu", Founded="Not publicly disclosed",
         Headquarters="Budapest, Hungary", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "Hungary, Greece, Slovakia"}, Ownership="Private", Employees="unspecified",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Diabetes Monitoring; others"},
         **{"Key Brands Represented": "MicroTech LinX; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="info@gluxshop.hu",
         LinkedIn="Not publicly disclosed", **{"Office No.": "+36 1 446-0023"},
         **{"Office Email": "info@gluxshop.hu"}, **{"Country Code": "+36"}, **{"Region Code": "1"},
         Summary="GluX is a Hungarian distributor specializing in diabetes devices. It is the official distributor of MicroTech's LinX CGM systems in Hungary, and a partner for distribution in Greece and Slovakia."),

    dict(Company="Evercare Medical", Website="https://www.evercare-medical.com", Founded="Not publicly disclosed",
         Headquarters="United Kingdom", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "Europe-wide"}, Ownership="Private (Asker)", Employees="33",
         **{"Core Categories": "Medical Supplies; others"},
         **{"Specialty Focus": "Medical Consumer Products; others"},
         **{"Key Brands Represented": "Not publicly disclosed"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": "Not publicly disclosed"}, Role="Not publicly disclosed", Email="info@evercare-medical.com",
         LinkedIn="Not publicly disclosed", **{"Office No.": "+44 116 251 0650"},
         **{"Office Email": "info@evercare-medical.com"}, **{"Country Code": "+44"}, **{"Region Code": "116"},
         Summary="Evercare Medical is a European healthcare distributor (part of the Asker group) headquartered in the UK, offering a range of medical supplies including aesthetics and general medical products across Europe."),
]


def main() -> None:
    wb = openpyxl.load_workbook(XLSX)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(header)}
    comp_col = idx["Company"]

    existing_names = {ws.cell(row=r, column=comp_col).value for r in range(2, ws.max_row + 1)}

    start = ws.max_row + 1
    added = 0
    for rec in NEW_ROWS:
        if rec["Company"] in existing_names:
            print("SKIP (already present):", rec["Company"])
            continue
        r = ws.max_row + 1
        for col_name in COLUMNS:
            ws.cell(row=r, column=idx[col_name]).value = rec.get(col_name, "Not publicly disclosed")
        added += 1

    wb.save(XLSX)
    print(f"Added {added} new companies. Row count now: {ws.max_row}")


if __name__ == "__main__":
    main()
