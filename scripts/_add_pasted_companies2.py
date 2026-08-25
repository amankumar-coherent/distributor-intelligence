#!/usr/bin/env python3
"""Second batch of pasted companies to add.

Per user decision, the ~17 "[No citation]" entries are added as-is.
Excluded regardless (separate issues from the citation question):
  - DKSH Group, Globalmed Ltd.: exact duplicates of existing rows (DKSH, Globalmed Malaysia)
  - Medhold Group: same real company as "Medhold" (South Africa), already
    researched and removed earlier this session for no confirmed wearable link
  - KM Medidistributors: website given as "kmindia.com (Example)" -- an
    explicit fabricated placeholder domain, not a real site
  - Nihon Kohden Asia, Greenmeditech Ltd.: manufacturers of their own
    proprietary devices, not third-party channel partners/distributors
    (contradicts this sheet's distributor-only scope)
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

NPD = "Not publicly disclosed"

NEW_ROWS = [
    dict(Company="Byto Med India", Website="https://www.bytomedindia.com", Founded=NPD,
         Headquarters="Bareilly, Uttar Pradesh, India", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "India"}, Ownership="Private", Employees=NPD,
         **{"Core Categories": "Medical Devices; Patient Monitors; others"},
         **{"Specialty Focus": "Operation Theatre, ICU; others"},
         **{"Key Brands Represented": NPD}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email="info@bytomedindia.com",
         LinkedIn="https://www.linkedin.com/company/bytomedindia", **{"Office No.": "+91 97619 28525"},
         **{"Office Email": "info@bytomedindia.com"}, **{"Country Code": "+91"}, **{"Region Code": "243"},
         Summary="Byto Med India is a distributor of high-grade hospital equipment (ECG monitors, ventilators, OT lights, neonatal care), serving hospitals across India with a focus on operation-theatre and critical-care devices."),

    dict(Company="CIS Imaging Co., Ltd.", Website="https://www.cis-imaging.com", Founded=2023,
         Headquarters="Bangkok, Thailand", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "Thailand"}, Ownership="Public", Employees="2-10",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Remote Patient Monitoring; others"},
         **{"Key Brands Represented": "Corsano; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn="https://www.linkedin.com/company/cishealthcare", **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+66"}, **{"Region Code": "10170"},
         Summary="CIS Imaging (Thailand) is a Bangkok-based distributor specializing in healthcare IT and wearable monitoring solutions, partnering with Corsano Health for CardioWatch wearable ECG/PPG devices."),

    dict(Company="Connected Healthcare Systems", Website="https://www.chs.co.nz", Founded=NPD,
         Headquarters="Auckland, New Zealand", **{"Continent / Geography": "Oceania"},
         **{"Operational Presence": "New Zealand"}, Ownership="Private", Employees=NPD,
         **{"Core Categories": "Medical Devices; Patient Monitors; others"},
         **{"Specialty Focus": "ICU/Ward Monitoring; others"},
         **{"Key Brands Represented": "Mindray; Vocera; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email="info@chs.co.nz",
         LinkedIn="https://www.linkedin.com/company/connected-healthcare-systems", **{"Office No.": "+64 800 424 797"},
         **{"Office Email": "info@chs.co.nz"}, **{"Country Code": "+64"}, **{"Region Code": "9"},
         Summary="Connected Healthcare Systems (NZ) is a New Zealand-owned distributor for Mindray patient monitoring and anesthesia systems, also representing Vocera communication devices."),

    dict(Company="Inventure Medical", Website=NPD, Founded=NPD,
         Headquarters="Cairo, Egypt", **{"Continent / Geography": "Africa"},
         **{"Operational Presence": "Middle East & Africa"}, Ownership="Private (Inventure Capital Group)", Employees=NPD,
         **{"Core Categories": "Medical Devices; Respiratory Equipment; others"},
         **{"Specialty Focus": "Hospital Ventilator Systems; others"},
         **{"Key Brands Represented": "Philips; Siemens; Abbott; others"}, **{"Online / Offline / Both": "Offline"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+20"}, **{"Region Code": NPD},
         Summary="Inventure Medical (Egypt), part of a private equity group, became a leading distributor of advanced hospital equipment (ventilators, diagnostics, vaccines) during COVID, supplying equipment across Africa and the Gulf via affiliates."),

    dict(Company="Ibnsina Medical", Website="https://www.ibnsina-pharma.com", Founded=2001,
         Headquarters="Cairo, Egypt", **{"Continent / Geography": "Africa"},
         **{"Operational Presence": "Egypt"}, Ownership="Private (Ibnsina Pharma Group)", Employees=NPD,
         **{"Core Categories": "Medical Devices; Pharmaceuticals; others"},
         **{"Specialty Focus": "General Device Distribution, Cold-Chain Logistics; others"},
         **{"Key Brands Represented": NPD}, **{"Online / Offline / Both": "Offline"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+20"}, **{"Region Code": NPD},
         Summary="Ibnsina Medical (Egypt), a subsidiary of Ibnsina Pharma, is one of Egypt's largest medical distributors with over 50 warehouses, importing and distributing hospital equipment and pharmaceuticals nationwide."),

    dict(Company="Trivitron Africa", Website="https://www.trivitronafrica.com", Founded=2018,
         Headquarters="Dubai, United Arab Emirates", **{"Continent / Geography": "Middle East / Africa"},
         **{"Operational Presence": "15+ African countries"}, Ownership="Joint Venture (Trivitron & IFHA)", Employees=NPD,
         **{"Core Categories": "Medical Devices; Diagnostic Imaging; others"},
         **{"Specialty Focus": "Hospital Projects, Lab Diagnostics; others"},
         **{"Key Brands Represented": "Trivitron; Shimadzu; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+971"}, **{"Region Code": NPD},
         Summary="Trivitron Africa (Dubai) is a joint venture between India's Trivitron and Investment Funds for Health in Africa (IFHA), delivering lab diagnostics, imaging, and ICU equipment across roughly 15 African countries."),

    dict(Company="Vays Medical", Website="https://www.vays-medical.com", Founded=NPD,
         Headquarters="Miami, FL, United States", **{"Continent / Geography": "Americas"},
         **{"Operational Presence": "Latin America"}, Ownership="Private", Employees="~10",
         **{"Core Categories": "Medical Devices; Patient Monitors; others"},
         **{"Specialty Focus": "LATAM Market Access, Critical Care; others"},
         **{"Key Brands Represented": NPD}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn="https://www.linkedin.com/company/vaysmedical", **{"Office No.": "+1 (786) 784-2091"},
         **{"Office Email": NPD}, **{"Country Code": "+1"}, **{"Region Code": "305"},
         Summary="Vays Medical is a US-based MedTech market-access partner for Latin America, connecting US/EU OEMs to LATAM distributors and hospitals with a focus on monitoring and critical-care products."),

    dict(Company="ACS Medical GmbH", Website="https://www.acs-medical.de", Founded=1970,
         Headquarters="Bergisch Gladbach, Germany", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "Germany, Austria, Switzerland"}, Ownership="Private", Employees="150+",
         **{"Core Categories": "Diagnostic Wearables; Patient Monitors; others"},
         **{"Specialty Focus": "Cardiology, Telemetry; others"},
         **{"Key Brands Represented": "Philips; GE; Schiller; others"}, **{"Online / Offline / Both": "Offline"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+49"}, **{"Region Code": NPD},
         Summary="ACS Medical GmbH distributes diagnostic cardiology and patient monitoring systems across the DACH region, with long-term partnerships with Philips and GE."),

    dict(Company="Genexo", Website="https://www.genexo.eu", Founded=2019,
         Headquarters="Warsaw, Poland", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "Poland"}, Ownership="Private", Employees="~20",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Diabetes Care; others"},
         **{"Key Brands Represented": "Dexcom; FreeStyle Libre; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+48"}, **{"Region Code": NPD},
         Summary="Genexo is Dexcom's exclusive Polish distributor for FreeStyle Libre and Dexcom CGM systems, operating clinics and retail presence in major Polish cities."),

    dict(Company="Shelltek Enterprise", Website="https://www.shelltekenterprises.com", Founded=NPD,
         Headquarters="Mumbai, India", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "India"}, Ownership="Private", Employees="50-100",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Diabetes Care; others"},
         **{"Key Brands Represented": "Roche; Medtronic; Tandem; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+91"}, **{"Region Code": "22"},
         Summary="Shelltek Enterprise (India) distributes insulin pumps and diabetes devices such as Roche Accu-Chek, serving hospitals and patients across India via online and retail channels."),

    dict(Company="GlucoRx UK Ltd", Website="https://www.glucorx.co.uk", Founded=1999,
         Headquarters="Essex, United Kingdom", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "United Kingdom"}, Ownership="Private", Employees="50+",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Diabetes Technology; others"},
         **{"Key Brands Represented": "Dexcom; Accu-Chek; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+44"}, **{"Region Code": NPD},
         Summary="GlucoRx UK offers blood glucose monitors and CGM systems, acting as an authorized distributor for Dexcom and supplying diabetes products to pharmacies and clinics in the UK."),

    dict(Company="Medimed Ltd.", Website="https://www.medimed.hu", Founded=1998,
         Headquarters="Budapest, Hungary", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "Hungary, Slovakia, Romania"}, Ownership="Private", Employees="20-50",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Cardiology Monitoring; others"},
         **{"Key Brands Represented": "GE Healthcare; BTL; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+36"}, **{"Region Code": NPD},
         Summary="Medimed Ltd. (Hungary) distributes cardiac and monitoring devices including Holter and telemetry systems across Central Europe."),

    dict(Company="C-Care Medical", Website=NPD, Founded=NPD,
         Headquarters="Singapore", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "Southeast Asia"}, Ownership="Private", Employees=NPD,
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Emergency Medicine; others"},
         **{"Key Brands Represented": "Mindray; Fukuda Denshi; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+65"}, **{"Region Code": NPD},
         Summary="C-Care Medical (Singapore) supplies vital sign monitors and ECG devices to Southeast Asian hospitals."),

    dict(Company="Medi-Serve", Website=NPD, Founded=NPD,
         Headquarters="Hong Kong SAR, China", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "Hong Kong, China"}, Ownership="Private", Employees=NPD,
         **{"Core Categories": "Respiratory Equipment; others"},
         **{"Specialty Focus": "Homecare, Rehab; others"},
         **{"Key Brands Represented": "ResMed; Philips; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+852"}, **{"Region Code": NPD},
         Summary="Medi-Serve (Hong Kong) distributes respiratory and sleep therapy devices, including CPAP/BiPAP units and oximeters, across Hong Kong and China."),

    dict(Company="Maxcare Scientific", Website="https://www.maxcarescientific.com", Founded=1999,
         Headquarters="Sydney, Australia", **{"Continent / Geography": "Oceania"},
         **{"Operational Presence": "Australia"}, Ownership="Private", Employees="50-100",
         **{"Core Categories": "Diagnostic Wearables; others"},
         **{"Specialty Focus": "Diabetes Devices; others"},
         **{"Key Brands Represented": "FreeStyle Libre; Medtronic; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+61"}, **{"Region Code": NPD},
         Summary="Maxcare Scientific (Australia) specializes in diabetes care, distributing CGM systems such as FreeStyle Libre and insulin pump consumables nationwide."),

    dict(Company="Servimedical", Website="https://www.servimedical.com.mx", Founded=1990,
         Headquarters="Mexico City, Mexico", **{"Continent / Geography": "Americas"},
         **{"Operational Presence": "Mexico"}, Ownership="Private", Employees="100+",
         **{"Core Categories": "Diagnostic Wearables; Patient Monitors; others"},
         **{"Specialty Focus": "Cardiology Devices; others"},
         **{"Key Brands Represented": "GE; Mindray; Capnostream; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+52"}, **{"Region Code": NPD},
         Summary="Servimedical (Mexico) provides diagnostic cardiology (ECG, Holter) and patient monitoring systems, partnering with GE and Mindray."),

    dict(Company="AAC Medical", Website="https://www.aacmedical.ae", Founded=NPD,
         Headquarters="Dubai, United Arab Emirates", **{"Continent / Geography": "Middle East"},
         **{"Operational Presence": "UAE, GCC"}, Ownership="Private", Employees="~20",
         **{"Core Categories": "Respiratory Equipment; Mobility Aids; others"},
         **{"Specialty Focus": "Home Oxygen, Rehab; others"},
         **{"Key Brands Represented": "ResMed; Philips Respironics; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+971"}, **{"Region Code": NPD},
         Summary="AAC Medical (UAE) distributes home and hospital medical devices including oxygen concentrators, CPAP machines, and mobility aids across the Gulf."),

    dict(Company="MediPro Group", Website="https://www.mediprogroup.co.uk", Founded=NPD,
         Headquarters="Nottingham, United Kingdom", **{"Continent / Geography": "Europe"},
         **{"Operational Presence": "UK, Ireland"}, Ownership="Private", Employees="100-200",
         **{"Core Categories": "Diagnostic Wearables; Patient Monitors; others"},
         **{"Specialty Focus": "Intensive Care; others"},
         **{"Key Brands Represented": "Nihon Kohden; Schiller; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+44"}, **{"Region Code": NPD},
         Summary="MediPro Group (UK) supplies cardiac diagnostic and monitoring systems, including ECG and telemetry, to hospitals across the UK and Ireland."),

    dict(Company="Vaidya Medical Systems", Website="https://www.vaidyamedical.com", Founded=1989,
         Headquarters="Mumbai, India", **{"Continent / Geography": "Asia"},
         **{"Operational Presence": "India"}, Ownership="Private", Employees="50-100",
         **{"Core Categories": "Medical Devices; Patient Monitors; others"},
         **{"Specialty Focus": "Operating Room, ICU; others"},
         **{"Key Brands Represented": "Dräger; Maquet; GE; others"}, **{"Online / Offline / Both": "Both"},
         **{"Contact Person": NPD}, Role=NPD, Email=NPD,
         LinkedIn=NPD, **{"Office No.": NPD},
         **{"Office Email": NPD}, **{"Country Code": "+91"}, **{"Region Code": "22"},
         Summary="Vaidya Medical Systems is a Mumbai-based distributor of surgical, ICU, and monitoring equipment, including Dräger ventilators and GE patient monitors."),
]


def main() -> None:
    wb = openpyxl.load_workbook(XLSX)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(header)}
    comp_col = idx["Company"]

    existing_names = {ws.cell(row=r, column=comp_col).value for r in range(2, ws.max_row + 1)}

    added = 0
    for rec in NEW_ROWS:
        if rec["Company"] in existing_names:
            print("SKIP (already present):", rec["Company"])
            continue
        r = ws.max_row + 1
        for col_name in COLUMNS:
            ws.cell(row=r, column=idx[col_name]).value = rec.get(col_name, NPD)
        added += 1

    wb.save(XLSX)
    print(f"Added {added} new companies. Row count now: {ws.max_row}")


if __name__ == "__main__":
    main()
