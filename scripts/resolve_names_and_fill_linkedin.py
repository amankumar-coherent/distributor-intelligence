"""Resolve messy company names + apply LinkedIn/web-found Founded/HQ/Cities.

Only updates rows where name resolution is confident (website/LI match).
"""
from __future__ import annotations

from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

# old Excel name -> updates
RESOLVE: dict[str, dict[str, str]] = {
    # LinkedIn + web: Rockland Industries / ROC-LON muslin since 1832
    "Roc Lonfabrics": {
        "resolved_name": "ROC-LON Fabrics (Rockland Industries)",
        "founded": "1832",
        "hq": "Baltimore, Maryland, USA (heritage HQ); manufacturing Bamberg, South Carolina, USA",
        "cities": "Baltimore, MD; Bamberg, SC; USA; export 90+ countries historically",
        "note": "Name resolved Roc Lonfabrics→ROC-LON/Rockland Industries; LI /company/roc-lon-fabrics + web About since 1832",
    },
    # LinkedIn: Vision West Inc. (optometric buying group) — matches channel intent
    "Optometric / ENT buying groups channel": {
        "resolved_name": "Vision West Inc. (VWI)",
        "founded": "1988",
        "hq": "15970 Bernardo Center Drive, San Diego, CA 92127, USA",
        "cities": "San Diego, California, USA; nationwide optometric membership",
        "note": "Name resolved VisionWest channel→Vision West Inc.; LI /company/vwi Founded 1988",
    },
    # LinkedIn: Central India Fabrics Indore (related textile listing)
    "Centralindiafabrics": {
        "resolved_name": "Central India Fabrics",
        "hq": "Sector-B, Sawer Road Industrial Area, near Shiv Nagar, Shade No.4, Indore, Madhya Pradesh 452015, India",
        "cities": "Indore, Madhya Pradesh, India",
        "note": "Name resolved Centralindiafabrics→Central India Fabrics; LI /company/central-india-fabrics",
    },
    # LinkedIn auto page
    "Bharuchaassociates": {
        "resolved_name": "Bharucha Associates",
        "hq": "India (LinkedIn: Bharucha Associates - India)",
        "cities": "India",
        "note": "Name resolved Bharuchaassociates→Bharucha Associates; LI /company/bharucha-associates---india",
    },
    # Web+LI slogan since 1832 already used; keep SyFabrics etc. already done
    "Casafashions": {
        "resolved_name": "Casa Fashions (Karur)",
        "hq": "1/81-1, K.E. Rengasamy Nagar, Andankoil East, Athur Post, Karur 639008, Tamil Nadu, India",
        "cities": "Karur, Tamil Nadu, India",
        "note": "Name resolved Casafashions→Casa Fashions Karur (website address); no LI page found",
    },
    "Anand Agencies": {
        "resolved_name": "Anand Agencies (Chennai)",
        "hq": "No. S2/61, Karpagam Nagar 2nd St, Selaiyur, Chennai, Tamil Nadu 600073, India",
        "cities": "Chennai, Tamil Nadu, India",
        "note": "Name kept Anand Agencies Chennai (website); LI had multiple Anand Agencies — Chennai beauty match not found",
    },
    "Laojiangtaigroup": {
        "resolved_name": "Lao Jiangtai Liquor Industry (Group) Co., Ltd.",
        "hq": "Guizhou, China",
        "cities": "Guizhou, China; China distribution",
        "note": "Name resolved Laojiangtaigroup→Lao Jiangtai Liquor Group; no LI company page",
    },
    "Shaoxing Xiaonizi Textile": {
        "resolved_name": "Shaoxing Xiaonizi Textile Co., Ltd.",
        "hq": "Shaoxing, Zhejiang, China",
        "cities": "Shaoxing, China; global export",
        "note": "Name cleaned; website xnztex.com; no LI page",
    },
    "Guangdong Jiujiu Liquor Co.,": {
        "resolved_name": "Guangdong Jiujiu Liquor Co., Ltd.",
        "hq": "Guangdong, China",
        "cities": "Guangdong, China",
        "note": "Name cleaned; no LI page",
    },
    "DrinkUpNY / specialty spirits e-retail": {
        "resolved_name": "DrinkUpNY",
        "hq": "New York, USA",
        "cities": "New York, USA; USA shipping states",
        "note": "Name resolved DrinkUpNY; no LI company page found",
    },
    "BaijiuShop / specialty online importers": {
        "resolved_name": "BaijiuShop",
        "hq": "International online",
        "cities": "Select export / online markets",
        "note": "Name resolved BaijiuShop; no LI company page",
    },
    "Elderberryfarmsestate": {
        "resolved_name": "Elderberry Farms Estate",
        "hq": "USA",
        "cities": "USA",
        "note": "Name resolved Elderberryfarmsestate→Elderberry Farms Estate; no LI page",
    },
    "Elderfarms": {
        "resolved_name": "Elder Farms",
        "hq": "USA",
        "cities": "USA",
        "note": "Name resolved Elderfarms→Elder Farms; no LI page",
    },
    "Heartlandelderberryfarms": {
        "resolved_name": "Heartland Elderberry Farms",
        "hq": "USA Midwest",
        "cities": "USA Midwest",
        "note": "Name resolved Heartlandelderberryfarms→Heartland Elderberry Farms; no LI page",
    },
    "Hanyinails": {
        "resolved_name": "Hanyi Nails",
        "hq": "China (OEM/private-label nail polish)",
        "cities": "China; global OEM/wholesale",
        "note": "Name resolved Hanyinails→Hanyi Nails; no LI page",
    },
    "Dmdistributors": {
        "resolved_name": "DM Distributors",
        "hq": "USA",
        "cities": "USA",
        "note": "Name resolved Dmdistributors→DM Distributors; no LI page",
    },
    "Doraykumas": {
        "resolved_name": "Doray Kumas",
        "hq": "Turkey",
        "cities": "Turkey; fabric wholesale",
        "note": "Name resolved Doraykumas→Doray Kumas; no LI page",
    },
    "Countrytextile": {
        "resolved_name": "Country Textile",
        "hq": "Online textile store",
        "cities": "Online international",
        "note": "Name resolved Countrytextile→Country Textile; no LI page",
    },
    "Fabricfits": {
        "resolved_name": "Fabric Fits",
        "hq": "Online",
        "cities": "Online",
        "note": "Name resolved Fabricfits→Fabric Fits; no LI page",
    },
    "Dfsbikefork": {
        "resolved_name": "DFS Bike Fork",
        "hq": "China / Asia MTB fork brand",
        "cities": "China; global online dealers",
        "note": "Name resolved Dfsbikefork→DFS Bike Fork; no LI page",
    },
    "Fullsuspension": {
        "resolved_name": "Full Suspension",
        "hq": "33391 Rhinestone Rd, Wausau, Wisconsin, USA",
        "cities": "Wausau, Wisconsin, USA",
        "note": "Name resolved Fullsuspension→Full Suspension; website address; no clear LI match",
    },
    "Electric Hc": {
        "resolved_name": "Electric HC (Yueqing)",
        "hq": "Liushi Town, Yueqing City, Zhejiang, China; Sales center Suzhou, Jiangsu",
        "cities": "Yueqing / Suzhou, China",
        "note": "Name cleaned Electric HC; website address; no LI page",
    },
    "Prosemicsr": {
        "resolved_name": "Prosemic",
        "hq": "#901, 9th Floor, No.2 Building, No.8 Dayan Road, Suzhou FTZ, Jiangsu, China",
        "cities": "Suzhou, Jiangsu, China",
        "note": "Name resolved Prosemicsr→Prosemic; website HQ; no LI page",
    },
    "Taikicosmetics": {
        "resolved_name": "Taiki Cosmetics",
        "hq": "Online cosmetics brand",
        "cities": "Online international",
        "note": "Name resolved Taikicosmetics→Taiki Cosmetics; no LI page",
    },
    "Doctoraltheaglobal": {
        "resolved_name": "Doctor Althea Global",
        "hq": "Online (doctoraltheaglobal.com)",
        "cities": "Online international",
        "note": "Name resolved Doctoraltheaglobal→Doctor Althea Global; no LI page",
    },
    "Spiritspring": {
        "resolved_name": "Spirit Spring",
        "hq": "Online spirits brand",
        "cities": "Online",
        "note": "Name resolved Spiritspring→Spirit Spring; no LI page",
    },
    "Zhonghuabaijiu": {
        "resolved_name": "Zhonghua Baijiu",
        "hq": "China",
        "cities": "China",
        "note": "Name resolved Zhonghuabaijiu→Zhonghua Baijiu; no LI page",
    },
    "Zwbaijiu": {
        "resolved_name": "ZW Baijiu",
        "hq": "No.999 North Yuanjia Village, Longsansi Town, Shanghe County, Jinan, Shandong, China",
        "cities": "Jinan / Shanghe, Shandong, China",
        "note": "Name resolved Zwbaijiu→ZW Baijiu; website address",
    },
    "BeautyManufacturingSolutions / OEM nail fillers": {
        "resolved_name": "Beauty Manufacturing Solutions (BMS Nail)",
        "hq": "USA / Asia OEM nail channel (bmsnail.com)",
        "cities": "Global private-label OEM",
        "note": "Name resolved→BMS Nail / Beauty Manufacturing Solutions; no LI page",
    },
    "Nailish / EU gel distributors": {
        "resolved_name": "Nailish",
        "hq": "Europe",
        "cities": "Europe",
        "note": "Name cleaned Nailish; no LI page",
    },
    "Korean Beauty Supply": {
        "resolved_name": "Korean Beauty Supply",
        "hq": "USA K-beauty import/retail channel",
        "cities": "USA; Korea sourcing",
        "note": "Name kept; LI search returned other K-beauty wholesalers, no exact page",
    },
    "Mountain Medical Equipment distributors": {
        "resolved_name": "Mountain Medical Equipment",
        "hq": "USA (mountainmedical.com)",
        "cities": "USA",
        "note": "Name cleaned; LI has similarly named Idaho DME (mmedme.com) — not confirmed same entity",
    },
    "Sewing Machine Warehouse": {
        "resolved_name": "Sewing Machine Warehouse (Cypress, TX)",
        "hq": "13233 N Eldridge Pkwy, Cypress, Texas, USA",
        "cities": "Cypress, Texas, USA",
        "note": "Name kept Cypress TX entity; LI 'Sewing Machine Warehouse' North Hills CA is different (kneedle.com)",
    },
    "Sewing Machine Plus": {
        "resolved_name": "Sewing Machine Plus",
        "hq": "USA",
        "cities": "USA online",
        "note": "Name kept; no exact LI match (Colette's Sewing Machines Plus is different)",
    },
    "Wholesale Nail Supply": {
        "resolved_name": "Wholesale Nail Supply",
        "hq": "USA (wholesalenailsupply.com)",
        "cities": "USA wholesale",
        "note": "Name kept; LI has US Maxim / Princess Nail Supply as nearby wholesalers — not same website",
    },
    "A & N Fabrics": {
        "resolved_name": "A & N Fabrics",
        "hq": "New York, NY, USA (apparel fabric district)",
        "cities": "New York, NY, USA",
        "note": "Name kept; no exact LI match",
    },
    "Fsykmed": {
        "resolved_name": "FSYK Med",
        "hq": "No. 23-6, Yanfeng Industrial Zone, Shishan Town, Nanhai District, Foshan, Guangdong, China",
        "cities": "Foshan, Guangdong, China",
        "note": "Name resolved Fsykmed→FSYK Med; website address; no LI page",
    },
    "Imcwc": {
        "resolved_name": "IMCWC",
        "hq": "2470 Patterson Rd #8, Grand Junction, CO 81505, USA",
        "cities": "Grand Junction, Colorado, USA",
        "note": "Name cleaned IMCWC; website address; no LI match",
    },
    "Borneoseller": {
        "resolved_name": "Borneo Seller",
        "hq": "Worldwide B2B seller network (India-linked listings)",
        "cities": "Global B2B; India sellers",
        "note": "Name resolved Borneoseller→Borneo Seller; marketplace; no LI page",
    },
    "Hethstia": {
        "resolved_name": "Hethstia",
        "hq": "Online instant-food / noodle brand",
        "cities": "Online international",
        "note": "Name kept; no LI page",
    },
    "Sinomie Noodles": {
        "resolved_name": "Sinomie Noodles",
        "hq": "China",
        "cities": "China; export",
        "note": "Name kept; no LI page",
    },
    "Guimpex": {
        "resolved_name": "Guimpex",
        "hq": "International trade (guimpex.net)",
        "cities": "International",
        "note": "Name kept; LI 'gu impex' Hyderabad not confirmed same",
    },
    "Bnatur": {
        "resolved_name": "BNatur",
        "hq": "Armenia (bnatur.am)",
        "cities": "Armenia; online",
        "note": "Name cleaned BNatur; no LI page",
    },
    "Chinese Baijiu": {
        "resolved_name": "Chinese Baijiu (Norvee / Guiyang export listing)",
        "hq": "Guiyang, China (office listed on chinese-baijiu.com)",
        "cities": "Guiyang, China",
        "note": "Name contextualized; site is brand/export listing not single corp; no LI page",
    },
    # Government / marketplace — resolve label clearly, do not invent corp founded
    "Gesund": {
        "resolved_name": "gesund.bund.de (German Federal Ministry of Health portal)",
        "founded": "N/A — government portal (not a distributor)",
        "hq": "Berlin, Germany — Bundesministerium für Gesundheit",
        "cities": "Germany",
        "note": "Name resolved: not a company; federal health portal",
    },
    "Pdunippd": {
        "resolved_name": "PDUNIPPD (Pt. Deendayal Upadhyaya National Institute for Persons with Physical Disabilities)",
        "founded": "N/A — government institute (not a distributor)",
        "hq": "Pune, Maharashtra, India",
        "cities": "Pune, India",
        "note": "Name resolved: government institute portal pdunippd.nic.in",
    },
    "Goldsupplier": {
        "resolved_name": "GoldSupplier.com (B2B supplier directory)",
        "founded": "N/A — marketplace directory (not a single distributor)",
        "hq": "China / Hong Kong B2B directory",
        "cities": "China; global B2B buyers",
        "note": "Name resolved: marketplace platform, not one company",
    },
    # Generic nail polish placeholder sites — resolve as online retailers
    "Nail Polish Boutique": {
        "resolved_name": "Nail Polish Boutique (online)",
        "hq": "Online retailer (nailpolishboutique.com)",
        "cities": "Online",
        "note": "Name labeled online retailer; no LI company page",
    },
    "Nail Polish Depot": {
        "resolved_name": "Nail Polish Depot (online)",
        "hq": "Online retailer (nailpolishdepot.com)",
        "cities": "Online",
        "note": "Name labeled online retailer; no LI company page",
    },
    "Nail Polish Direct": {
        "resolved_name": "Nail Polish Direct (online)",
        "hq": "Online retailer (nailpolishdirect.com)",
        "cities": "Online",
        "note": "Name labeled online retailer; no LI company page",
    },
    "Nail Polish Network": {
        "resolved_name": "Nail Polish Network (online)",
        "hq": "Online retailer (nailpolishnetwork.com)",
        "cities": "Online",
        "note": "Name labeled online retailer; no LI company page",
    },
    "Nail Polish Shop": {
        "resolved_name": "Nail Polish Shop (online)",
        "hq": "Online retailer (nailpolishshop.com)",
        "cities": "Online",
        "note": "Name labeled online retailer; no LI company page",
    },
    "Nail Supply House": {
        "resolved_name": "Nail Supply House",
        "hq": "USA professional nail supply",
        "cities": "USA",
        "note": "Name kept; no LI page",
    },
}


def main() -> None:
    n = 0
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(path)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=False))
        hdr_i = next(
            i
            for i, r in enumerate(rows[:12])
            if r and r[0].value and str(r[0].value).strip() == "Company"
        )
        header = [str(c.value or "").strip() for c in rows[hdr_i]]
        ci = header.index("Company")
        fi = header.index("Founded")
        hi = header.index("Headquarters")
        ri = header.index("Cities / Regions")
        si = header.index("Data Sources") if "Data Sources" in header else None
        changed = False
        for r in rows[hdr_i + 1 :]:
            company = str(r[ci].value or "").strip()
            if company not in RESOLVE:
                continue
            info = RESOLVE[company]
            r[ci].value = info["resolved_name"]
            if "founded" in info:
                r[fi].value = info["founded"]
            if "hq" in info:
                r[hi].value = info["hq"]
            if "cities" in info:
                r[ri].value = info["cities"]
            if si is not None:
                prev = str(r[si].value or "").strip()
                tag = "Name resolved"
                note = info.get("note", "")
                if tag not in prev:
                    r[si].value = (prev + "; " if prev else "") + f"{tag}: {note}"
            n += 1
            changed = True
            print(f"{path.parent.name}: {company} → {info['resolved_name']}")
        if changed:
            wb.save(path)
        wb.close()
    print(f"\nResolved rows: {n}")


if __name__ == "__main__":
    main()
