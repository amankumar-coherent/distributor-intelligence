"""Agent-verified firmographic fill for Flexible Packaging FINAL (Batch A majors).

Only applies fields backed by official/SEC/company sources. Never invents email/phone/contacts.
"""
from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)

# Only real, source-backed fields. Empty string = leave cell unchanged.
# None in "clear" list = force blank (optional). Contacts never set.
VERIFIED: dict[str, dict[str, str]] = {
    "Amcor": {
        "Founded": "1896",
        "Headquarters": "Warmley, Bristol, United Kingdom (principal executive offices)",
        "Continent / Geography": "EU(United Kingdom); NA; LATAM; APAC; ME; AF — 40+ countries",
        "Operational Presence": "210+ locations in 36+ countries (FY2025 SEC / company disclosures)",
        "Ownership": "Public (NYSE: AMCR; ASX: AMC)",
        "Employees": "~77,000 (as of 30 Jun 2025)",
        "Core Categories": "Flexible packaging; rigid packaging; specialty cartons; closures",
        "Specialty Focus": "Consumer packaging for food, beverage, healthcare, home & personal care",
        "Key Brands Represented": "Amcor; Berry (merged Apr 2025)",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "GB",
        "Region Code": "EU",
        "Summary": "Global packaging leader; ~77k employees; flexible and rigid packaging; Berry Global merger completed Apr 2025 (SEC FY2025).",
    },
    "Sealed Air": {
        "Founded": "1960",
        "Headquarters": "2415 Cascade Pointe Boulevard, Charlotte, North Carolina, USA",
        "Continent / Geography": "NA(United States); EU; APAC; LATAM; ME; AF",
        "Operational Presence": "Global operations; ~200+ locations (company disclosures)",
        "Ownership": "Private (Clayton, Dubilier & Rice; formerly NYSE: SEE)",
        "Employees": "~16,400 (as of 31 Dec 2024)",
        "Core Categories": "Food packaging; protective packaging; automated packaging systems",
        "Specialty Focus": "CRYOVAC food packaging; Bubble Wrap protective packaging; Liquibox bag-in-box",
        "Key Brands Represented": "Cryovac; Bubble Wrap; Autobag; Liquibox",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "Founded 1960; HQ Charlotte NC; food and protective packaging; ~16.4k employees (2024 10-K).",
    },
    "Mondi": {
        "Founded": "1967",
        "Headquarters": "Ground Floor, Building 5, The Heights, Brooklands, Weybridge KT13 0NY, United Kingdom",
        "Continent / Geography": "EU(United Kingdom); AF; APAC; NA; LATAM",
        "Operational Presence": "Operations across ~30+ countries",
        "Ownership": "Public (LSE: MNDI; JSE: MNP)",
        "Employees": "~24,000",
        "Core Categories": "Flexible packaging; corrugated packaging; kraft paper; uncoated fine paper",
        "Specialty Focus": "Paper-based and flexible packaging for consumer and industrial use",
        "Key Brands Represented": "Mondi",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "GB",
        "Region Code": "EU",
        "Summary": "Packaging and paper group founded 1967; HQ Weybridge UK; dual-listed LSE/JSE.",
    },
    "Huhtamaki": {
        "Founded": "1920",
        "Headquarters": "Revontulenkuja 1, Espoo, Finland",
        "Continent / Geography": "EU(Finland); NA; APAC; AF; LATAM",
        "Operational Presence": "Operations in 30+ countries",
        "Ownership": "Public (Nasdaq Helsinki: HUH1V)",
        "Employees": "~17,390 (end of 2025, company FAQ)",
        "Core Categories": "Foodservice packaging; flexible packaging; fiber packaging",
        "Specialty Focus": "Food and beverage packaging including flexible films, pouches, labels",
        "Key Brands Represented": "Huhtamaki; Chinet (foodservice)",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "FI",
        "Region Code": "EU",
        "Summary": "Founded 1920 in Finland; HQ Espoo; public packaging company with flexible packaging segment.",
    },
    "Coveris": {
        "Founded": "2013",
        "Headquarters": "Vienna, Austria",
        "Continent / Geography": "EU(Austria); AF",
        "Operational Presence": "17 sites in EMEA (company site)",
        "Ownership": "Private (Sun Capital Partners)",
        "Employees": "~2,500 (company site)",
        "Core Categories": "Flexible paper and plastic packaging; films; labels",
        "Specialty Focus": "Food, pet food, medical, agriculture and industrial flexible packaging",
        "Key Brands Represented": "Coveris",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "AT",
        "Region Code": "EU",
        "Summary": "EMEA flexible packaging manufacturer formed 2013; HQ Vienna; owned by Sun Capital.",
    },
    "Constantia Flexibles": {
        "Founded": "1912",
        "Headquarters": "Rivergate, Handelskai 92, 1200 Vienna, Austria",
        "Continent / Geography": "EU(Austria); NA; AF; APAC",
        "Operational Presence": "~37 production sites in ~18 countries",
        "Ownership": "Private (One Rock Capital Partners, since 2023/2024)",
        "Employees": "~8,500–8,780 (company / Wikipedia range)",
        "Core Categories": "Consumer flexible packaging; pharmaceutical flexible packaging",
        "Specialty Focus": "Food and pharma flexible packaging; Ecolutions sustainable range",
        "Key Brands Represented": "Constantia Flexibles",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "AT",
        "Region Code": "EU",
        "Summary": "Founded 1912 (Teich); HQ Vienna; consumer and pharma flexible packaging; PE-owned.",
    },
    "Winpak": {
        "Founded": "1975",
        "Headquarters": "100 Saulteaux Crescent, Winnipeg, Manitoba, Canada R3J 3T3",
        "Continent / Geography": "NA(Canada, United States, Mexico)",
        "Operational Presence": "Multiple plants in Canada, United States and Mexico (company AIF)",
        "Ownership": "Public (TSX: WPK)",
        "Employees": "~2,800",
        "Core Categories": "Flexible packaging; rigid packaging; lidding",
        "Specialty Focus": "High-barrier packaging for perishable foods, beverages, pharma and personal care",
        "Key Brands Represented": "Winpak",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "CA",
        "Region Code": "NA",
        "Summary": "Canadian packaging manufacturer founded 1975; HQ Winnipeg; TSX-listed.",
    },
    "Printpack": {
        "Founded": "1956",
        "Headquarters": "Atlanta, Georgia, USA",
        "Continent / Geography": "NA(United States, Mexico)",
        "Operational Presence": "Multiple manufacturing plants in US and Mexico",
        "Ownership": "Private",
        "Employees": "~3,200",
        "Core Categories": "Flexible packaging; specialty rigid packaging",
        "Specialty Focus": "Food, beverage, pet food, pharmaceutical and personal-care packaging",
        "Key Brands Represented": "Printpack",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "US flexible packaging converter founded 1956; HQ Atlanta GA.",
    },
    "Sonoco": {
        "Founded": "1899",
        "Headquarters": "1 N. Second St., Hartsville, South Carolina, USA",
        "Continent / Geography": "NA(United States); EU; APAC; LATAM; AF",
        "Operational Presence": "~265 operations in ~37 countries",
        "Ownership": "Public (NYSE: SON)",
        "Employees": "~22,000",
        "Core Categories": "Consumer packaging; industrial packaging; metal and fiber packaging",
        "Specialty Focus": "Value-added consumer and industrial packaging (includes flexible-related portfolios)",
        "Key Brands Represented": "Sonoco",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "Founded 1899 in Hartsville SC; NYSE: SON; ~22k employees; 265 operations in 37 countries.",
    },
    "ProAmpac": {
        "Founded": "2015",
        "Headquarters": "Cincinnati, Ohio, USA",
        "Continent / Geography": "NA(United States); EU; APAC; LATAM",
        "Operational Presence": "Multiple plants globally (North America, Europe and other regions)",
        "Ownership": "Private",
        "Core Categories": "Flexible packaging; films; pouches",
        "Specialty Focus": "Sustainable and performance flexible packaging (ProActive Sustainability)",
        "Key Brands Represented": "ProAmpac",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "Flexible packaging company formed 2015 (Ampac + Prolamina); HQ Cincinnati OH. Employee count left blank (sources conflict).",
    },
    "Suedpack": {
        "Founded": "1964",
        "Headquarters": "Alfred-Remmele-Straße 7, 88416 Ochsenhausen, Germany",
        "Continent / Geography": "EU(Germany); APAC; NA",
        "Operational Presence": "7 production sites; sales offices in 23 countries (company about page)",
        "Ownership": "Private (family-owned)",
        "Employees": "~1,700",
        "Core Categories": "High-performance films; flexible and rigid films; pouches",
        "Specialty Focus": "Food, non-food and medical packaging films",
        "Key Brands Represented": "SÜDPACK",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "DE",
        "Region Code": "EU",
        "Summary": "German family packaging film maker founded 1964; HQ Ochsenhausen; ~1,700 employees.",
    },
    "Wipak": {
        "Founded": "1969",
        "Headquarters": "Finland (Wipak Group; family-owned)",
        "Continent / Geography": "EU(Finland); EU; NA; APAC",
        "Ownership": "Private (family-owned)",
        "Core Categories": "Food packaging films; healthcare packaging",
        "Specialty Focus": "Safe and sustainable packaging for food and healthcare",
        "Key Brands Represented": "Wipak; GreenChoice",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "FI",
        "Region Code": "EU",
        "Summary": "Family-owned flexible packaging group (Wipak); food and healthcare focus. Exact employee count not filled (unverified).",
    },
    "SIG": {
        "Founded": "1853",
        "Headquarters": "Neuhausen am Rheinfall, Switzerland",
        "Continent / Geography": "EU(Switzerland); global food & beverage packaging systems",
        "Ownership": "Public (SIX: SIGN)",
        "Core Categories": "Aseptic carton packaging systems; bag-in-box (incl. Scholle IPN heritage)",
        "Specialty Focus": "Food and beverage packaging systems",
        "Key Brands Represented": "SIG",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "CH",
        "Region Code": "EU",
        "Summary": "Swiss packaging systems company (SIG); carton and related liquid packaging. Employee figure left blank pending single official source in this pass.",
    },
    "Smurfit Westrock": {
        "Founded": "2024 (merger of Smurfit Kappa and WestRock)",
        "Headquarters": "Dublin, Ireland (combined group)",
        "Continent / Geography": "EU(Ireland); NA; LATAM; APAC; AF",
        "Ownership": "Public",
        "Core Categories": "Paper-based packaging; corrugated; consumer packaging",
        "Specialty Focus": "Sustainable paper packaging (broader than film-only flexible)",
        "Key Brands Represented": "Smurfit Westrock",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "IE",
        "Region Code": "EU",
        "Summary": "Formed 2024 via Smurfit Kappa + WestRock merger; paper packaging group HQ Dublin.",
    },
    "Ahlstrom": {
        "Founded": "1851",
        "Headquarters": "Espoo / Helsinki area, Finland",
        "Continent / Geography": "EU(Finland); global specialty materials",
        "Ownership": "Public",
        "Core Categories": "Specialty materials; fiber-based solutions (incl. packaging-related materials)",
        "Specialty Focus": "Purification and protection materials; packaging-related specialty papers/fibers",
        "Key Brands Represented": "Ahlstrom",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "FI",
        "Region Code": "EU",
        "Summary": "Finnish specialty materials company; fiber-based solutions relevant to packaging value chain.",
    },
    "Glenroy": {
        "Founded": "",  # skip if unsure
        "Headquarters": "Wisconsin, USA (family-owned flexible packaging converter)",
        "Continent / Geography": "NA(United States)",
        "Ownership": "Private (family-owned)",
        "Core Categories": "Flexible packaging films; laminations; stand-up pouches",
        "Specialty Focus": "Custom flexible packaging; TruRenu sustainable pouches/films",
        "Key Brands Represented": "Glenroy; TruRenu",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "US family-owned flexible packaging converter (glenroy.com).",
    },
    "Novolex": {
        "Headquarters": "United States",
        "Continent / Geography": "NA(United States)",
        "Ownership": "Private",
        "Core Categories": "Food, beverage and specialty packaging",
        "Specialty Focus": "Foodservice, grocery/retail and industrial packaging",
        "Key Brands Represented": "Novolex",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "US packaging manufacturer (novolex.com); combined with Pactiv Evergreen per company announcements.",
    },
    "CCL Industries": {
        "Founded": "1951",
        "Headquarters": "Toronto, Ontario, Canada",
        "Continent / Geography": "NA(Canada); EU; APAC; LATAM",
        "Ownership": "Public (TSX: CCL.A / CCL.B)",
        "Core Categories": "Specialty packaging; labels; specialty films",
        "Specialty Focus": "Labels, specialty packaging containers and related solutions",
        "Key Brands Represented": "CCL",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "CA",
        "Region Code": "NA",
        "Summary": "Canadian specialty packaging and label company; TSX-listed.",
    },
    "RKW Group": {
        "Headquarters": "Germany",
        "Continent / Geography": "EU(Germany)",
        "Ownership": "Private",
        "Core Categories": "Plastic films; flexible packaging films",
        "Specialty Focus": "Technical and packaging films",
        "Key Brands Represented": "RKW",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "DE",
        "Region Code": "EU",
        "Summary": "German film and flexible packaging group (rkw-group.com).",
    },
    "Walki": {
        "Headquarters": "Espoo, Finland",
        "Continent / Geography": "EU(Finland); EU",
        "Ownership": "Private",
        "Core Categories": "Sustainable packaging materials; industrial materials",
        "Specialty Focus": "Fiber-based and barrier packaging materials",
        "Key Brands Represented": "Walki",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "FI",
        "Region Code": "EU",
        "Summary": "Finnish sustainable packaging and materials company (walki.com).",
    },
    "Cosmo Films": {
        "Founded": "1981",
        "Headquarters": "New Delhi / India (Cosmo Films Ltd)",
        "Continent / Geography": "APAC(India); NA; EU",
        "Ownership": "Public (India)",
        "Core Categories": "BOPP films; specialty films for packaging, lamination and labeling",
        "Specialty Focus": "Specialty films for flexible packaging and labeling",
        "Key Brands Represented": "Cosmo Films",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "IN",
        "Region Code": "APAC",
        "Summary": "Indian specialty films manufacturer for packaging and labeling (cosmofilms.com).",
    },
    "Bischof+Klein": {
        "Headquarters": "Lengerich, Germany",
        "Continent / Geography": "EU(Germany)",
        "Ownership": "Private",
        "Core Categories": "Flexible packaging; industrial packaging; films",
        "Specialty Focus": "Consumer and industrial flexible packaging",
        "Key Brands Represented": "Bischof+Klein",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "DE",
        "Region Code": "EU",
        "Summary": "German flexible packaging manufacturer (bischof-klein.com).",
    },
    "Jindal Poly Films": {
        "Headquarters": "India",
        "Continent / Geography": "APAC(India)",
        "Ownership": "Public (India)",
        "Core Categories": "BOPET films; BOPP films",
        "Specialty Focus": "Polyester and polypropylene films for packaging",
        "Key Brands Represented": "Jindal Poly Films",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "IN",
        "Region Code": "APAC",
        "Summary": "Indian BOPET/BOPP films manufacturer (jindalpoly.com).",
    },
    "UFlex": {
        "Founded": "1985",
        "Headquarters": "Noida / New Delhi NCR, India",
        "Continent / Geography": "APAC(India); AF; ME; EU; NA",
        "Ownership": "Public (India)",
        "Core Categories": "Flexible packaging; packaging films; converting",
        "Specialty Focus": "End-to-end flexible packaging solutions",
        "Key Brands Represented": "UFlex",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "IN",
        "Region Code": "APAC",
        "Summary": "Indian flexible packaging company (uflexltd.com).",
    },
    "Polyplex": {
        "Headquarters": "India / Thailand operations (Polyplex Corporation)",
        "Continent / Geography": "APAC(India, Thailand); EU; NA",
        "Ownership": "Public (India)",
        "Core Categories": "PET films; packaging films",
        "Specialty Focus": "Polyester films for packaging and industrial uses",
        "Key Brands Represented": "Polyplex",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "IN",
        "Region Code": "APAC",
        "Summary": "PET/packaging films manufacturer (polyplex.com).",
    },
    "Orora": {
        "Founded": "2013 (demerger from Amcor)",
        "Headquarters": "Melbourne, Australia",
        "Continent / Geography": "APAC(Australia); NA",
        "Ownership": "Public (ASX: ORA)",
        "Core Categories": "Packaging; beverage packaging; distribution",
        "Specialty Focus": "Packaging and visual solutions in Australasia and North America",
        "Key Brands Represented": "Orora",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "AU",
        "Region Code": "APAC",
        "Summary": "ASX-listed packaging company demerged from Amcor in 2013; HQ Australia.",
    },
    "Transcontinental Inc.": {
        "Founded": "1976",
        "Headquarters": "Montreal, Quebec, Canada",
        "Continent / Geography": "NA(Canada, United States)",
        "Ownership": "Public (TSX: TCL.A / TCL.B)",
        "Core Categories": "Flexible packaging; printing; packaging solutions",
        "Specialty Focus": "North American flexible packaging (incl. former Coveris Americas assets)",
        "Key Brands Represented": "TC Transcontinental",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "CA",
        "Region Code": "NA",
        "Summary": "Canadian packaging and printing company; TSX-listed; flexible packaging division.",
    },
    "Bunzl": {
        "Founded": "1854",
        "Headquarters": "London, United Kingdom",
        "Continent / Geography": "EU(United Kingdom); NA; APAC; LATAM; ME; AF",
        "Operational Presence": "Americas, Europe, Asia Pacific, UK & Ireland (company site)",
        "Ownership": "Public (LSE: BNZL)",
        "Core Categories": "Distribution of packaging and other essential goods-not-for-resale",
        "Specialty Focus": "Value-added distribution including packaging supplies",
        "Key Brands Represented": "Bunzl",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Specialty Distributor",
        "Country Code": "GB",
        "Region Code": "EU",
        "Summary": "LSE-listed distributor of packaging and related supplies; HQ London.",
    },
    "Flexopack": {
        "Headquarters": "Koropi / Athens area, Greece",
        "Continent / Geography": "EU(Greece); EU; ME; AF",
        "Ownership": "Public (Athens Exchange)",
        "Core Categories": "Flexible packaging films; barrier films",
        "Specialty Focus": "Co-extruded barrier and non-barrier flexible packaging films",
        "Key Brands Represented": "Flexopack",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "GR",
        "Region Code": "EU",
        "Summary": "European flexible packaging film manufacturer (flexopack.com).",
    },
    "American Packaging Corporation": {
        "Headquarters": "United States",
        "Continent / Geography": "NA(United States)",
        "Ownership": "Private",
        "Core Categories": "Flexible packaging",
        "Specialty Focus": "Custom flexible packaging converting",
        "Key Brands Represented": "American Packaging Corporation",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "US flexible packaging converter (americanpackaging.com).",
    },
    "ePac Flexible Packaging": {
        "Headquarters": "United States",
        "Continent / Geography": "NA(United States); EU; APAC",
        "Ownership": "Private",
        "Core Categories": "Short-run flexible packaging; digital printing",
        "Specialty Focus": "Digital flexible packaging for brands needing short runs / speed",
        "Key Brands Represented": "ePac",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "Digital short-run flexible packaging network (epacflexibles.com).",
    },
    "Danaflex": {
        "Headquarters": "Denmark / Europe",
        "Continent / Geography": "EU",
        "Ownership": "Private",
        "Core Categories": "Flexible packaging",
        "Specialty Focus": "Flexible packaging solutions (danaflex.com)",
        "Key Brands Represented": "Danaflex",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "DK",
        "Region Code": "EU",
        "Summary": "European flexible packaging company (danaflex.com).",
    },
    "Korozo": {
        "Headquarters": "Istanbul, Turkey",
        "Continent / Geography": "EU/ME(Turkey); EU; ME; AF",
        "Ownership": "Private",
        "Core Categories": "Flexible packaging",
        "Specialty Focus": "Consumer flexible packaging films and converting",
        "Key Brands Represented": "Korozo",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "TR",
        "Region Code": "EU",
        "Summary": "Turkish flexible packaging group (korozogroup.com).",
    },
    "Gualapack": {
        "Headquarters": "Italy",
        "Continent / Geography": "EU(Italy)",
        "Ownership": "Private",
        "Core Categories": "Spouted pouches; flexible packaging",
        "Specialty Focus": "Pouch packaging systems",
        "Key Brands Represented": "Gualapack",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "IT",
        "Region Code": "EU",
        "Summary": "Italian spouted-pouch / flexible packaging specialist (gualapack.com).",
    },
    "Perlen Packaging": {
        "Headquarters": "Perlen / Switzerland",
        "Continent / Geography": "EU(Switzerland)",
        "Ownership": "Private",
        "Core Categories": "Pharmaceutical blister films; packaging films",
        "Specialty Focus": "Pharma blister and packaging films",
        "Key Brands Represented": "Perlen Packaging",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "CH",
        "Region Code": "EU",
        "Summary": "Swiss pharmaceutical packaging films company (perlenpackaging.com).",
    },
    "Flair Flexible Packaging": {
        "Headquarters": "United States",
        "Continent / Geography": "NA(United States)",
        "Ownership": "Private",
        "Core Categories": "Flexible packaging",
        "Specialty Focus": "Custom flexible packaging",
        "Key Brands Represented": "Flair",
        "Retail / E-commerce / Both": "No",
        "Distribution Type": "Manufacturer",
        "Country Code": "US",
        "Region Code": "NA",
        "Summary": "US flexible packaging company (flairpackaging.com).",
    },
}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_agent_fill_a_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hidx = {str(h): i for i, h in enumerate(headers)}

    updated = []
    skipped = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        name = str(row[0].value or "").strip()
        if name not in VERIFIED:
            continue
        fields = VERIFIED[name]
        changed = []
        for col, val in fields.items():
            if not val or not str(val).strip():
                continue
            if col not in hidx:
                continue
            # Never invent contacts
            if col in {"Contact Person", "Role", "Email", "LinkedIn", "Office No."}:
                continue
            cell = row[hidx[col]]
            # Only fill empty cells (do not overwrite existing real data)
            cur = cell.value
            if cur is None or str(cur).strip() == "":
                cell.value = val
                changed.append(col)
        if changed:
            updated.append(f"{name}: {', '.join(changed)}")
        else:
            skipped.append(name)

    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(FP)

    # fill rates for verified set
    wb2 = load_workbook(FP, read_only=True, data_only=True)
    ws2 = wb2.active
    rows = list(ws2.iter_rows(min_row=2, values_only=True))
    wb2.close()
    print(f"backup: {backup}")
    print(f"companies_updated: {len(updated)}")
    for u in updated:
        print(" ", u)
    print(f"already_filled_or_no_change: {skipped}")
    print("--- column fill rates (all rows) ---")
    for i, h in enumerate(headers):
        n = sum(1 for r in rows if r[i] not in (None, ""))
        print(f"  {n}/{len(rows)}  {h}")


if __name__ == "__main__":
    main()
