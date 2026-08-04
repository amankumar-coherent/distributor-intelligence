"""Apply extra employee/turnover bands for mid/small companies."""
from pathlib import Path

import openpyxl

EXTRA = {
    "AliMed": ("~200–500", "Private medical distributor"),
    "ScripHessco": ("~200–500", "Private chiropractic/DME distributor"),
    "Allegro Medical": ("~50–200", "Private online medical supply"),
    "Health Products For You": ("~50–200", "Private online DME retailer"),
    "Activeforever": ("~50–200", "Private online mobility retailer"),
    "School Health Corporation": ("~200–500", "Private school health distributor"),
    "School Health balance/screening supplies": ("~200–500", "Private school health distributor"),
    "Concordance Healthcare Solutions": ("~2,000+", "Private healthcare distributor"),
    "Direct Healthcare Group / UK DME": ("~500–1,000", "Private UK DME manufacturer/distributor"),
    "Aidapt / UK mobility wholesale": ("~50–200", "Private UK mobility wholesale"),
    "Breakthru Beverage Group": ("~5,000+", "Private wine/spirits wholesale"),
    "Young's Market Company": ("~2,000+", "Private West Coast spirits wholesale"),
    "Winebow Group": ("~1,000+", "Private fine wine wholesale"),
    "CNS Imports": ("~50–200", "Private spirits importer/distributor"),
    "Cnsimports": ("~50–200", "Private spirits importer/distributor"),
    "Opici Family Distributing": ("~500–1,000", "Private family spirits distributor"),
    "DrinkUpNY": ("~10–50", "Private NY liquor retailer"),
    "Shiwandistillery": ("~1,000+", "Guangdong Shiwan Wine Group"),
    "Duty Free Americas / travel retail spirits": ("~1,000+", "Private travel retail"),
    "DFS Group (travel retail)": ("~9,000+", "LVMH travel retail"),
    "China Duty Free Group": ("~10,000+", "State-owned / public travel retail"),
    "BevMo!": ("~3,000+", "GoPuff-owned beverage retail"),
    "Astor Wines & Spirits": ("~50–200", "Private NY wine retail"),
    "Summergate Fine Wines & Spirits": ("~200–500", "Private China importer"),
    "ASC Fine Wines (China on-trade/off-trade)": ("~500–1,000", "Private China wine importer"),
    "Jazz Trading (London) Ltd": ("~10–50", "Private UK importer"),
    "MHW Ltd.": ("~100–500", "Private spirits importer"),
    "e3 Diagnostics": ("~50–200", "Private audiology distributor"),
    "Allied Medical": ("~100–500", "Private AU medical distributor"),
    "Cedar Gate Technologies": ("~500–1,000", "Private healthcare technology"),
    "Natural Partners": ("~100–500", "Private practitioner distributor"),
    "Emerson Ecologics / Fullscript": ("~500–1,000", "Private practitioner platform"),
    "Wellevate / Prevention Wellness Dist.": ("~100–500", "Fullscript / practitioner channel"),
    "Elder Farms": ("~10–50", "Private family farm brand"),
    "Heartland Elderberry Farms": ("~10–50", "Private family farm brand"),
    "Garden of Life": ("~500–1,000", "Nestlé Health Science brand"),
    "Herb Pharm": ("~100–500", "Private herbal manufacturer"),
    "Sundown Naturals": ("~500–1,000 (brand)", "Nestlé Health Science / Nature's Bounty"),
    "Sundown Naturals Distribution": ("~500–1,000 (brand)", "Nestlé Health Science / Nature's Bounty"),
    "Vitacost": ("~1,000+ (Kroger e-comm)", "The Kroger Co."),
    "Vitacost (Kroger)": ("~1,000+ (Kroger e-comm)", "The Kroger Co."),
    "iHerb Wholesale / B2B channel": ("~2,000+", "Private online supplements"),
    "The Vitamin Shoppe Distribution": ("~5,000+", "Private specialty vitamin retail"),
    "Hethstia": ("~500–1,000", "Private konjac noodle manufacturer"),
    "Bicycle Warehouse": ("~50–200", "Private bike retail"),
    "Universal Cycles": ("~50–200", "Private bike parts wholesale"),
    "Worldwide Cyclery": ("~50–200", "Private online MTB retail"),
    "DVO Suspension distributors": ("~50–200", "Private MTB suspension brand"),
    "KHS Distribution / KHS Bicycles": ("~200–500", "Private bike brand/distributor"),
    "99 Bikes (AU retail chain)": ("~500–1,000", "Private AU bike retail"),
    "BikesOnline / DTC AU-NZ": ("~50–200", "Private DTC bike retail"),
    "BikeDiscount / Mantel": ("~200–500", "Private EU online bike retail"),
    "Bilstein": ("~1,000+", "thyssenkrupp suspension brand"),
    "Koni": ("~500–1,000", "Private shock absorber brand"),
    "Pon.Bike distribution (Cannondale, Santa Cruz parent ops)": (
        "~5,000+ (Pon.Bike)",
        "Pon Holdings bike group",
    ),
    "Mood Fabrics": ("~100–500", "Private fabric wholesale/retail"),
    "Sewing Machine Plus": ("~50–200", "Private family sewing retail"),
    "SyFabrics": ("~10–50", "Private LA fabric wholesale"),
    "Casa Fashions (Karur)": ("~50–200", "Private India textile manufacturer"),
    "Central India Fabrics": ("~50–200", "Private India fabric manufacturer"),
    "ROC-LON Fabrics (Rockland Industries)": ("~100–500", "Private textile manufacturer"),
    "Whaleys (Bradford) Ltd": ("~50–200", "Private UK theatrical fabrics"),
    "Liberty Fabrics (Wholesale)": ("~500–1,000 (Liberty)", "Liberty London fabrics"),
    "Rose Brand": ("~100–500", "Private theatrical textiles"),
    "OnlineFabricStore": ("~50–200", "Private online fabric retail"),
    "DM Distributors": ("~10–50", "Private upholstery supply"),
    "Doray Kumas": ("~10–50", "Private Turkey textile wholesale"),
    "Fabriclore": ("~50–200", "Private India B2B fabric"),
    "Spotlight (Australia/NZ)": ("~5,000+", "Private craft/fabric retail AU/NZ"),
    "Beauty Joint": ("~50–200", "Private nail beauty distributor"),
    "Dearsundays": ("~50–200", "Private nail wellness brand"),
    "The Nailhub / pro e-wholesalers": ("~10–50", "Private pro nail e-commerce"),
    "Kaganails": ("~10–50", "Private UK gel nail distributor"),
    "Burmax Beauty Distributors": ("~200–500", "Private salon beauty wholesale"),
    "Marianna Industries distribution": ("~200–500", "Private beauty wholesale"),
    "Beauty Manufacturing Solutions (BMS Nail)": ("~100–500", "Private cosmetic manufacturer"),
    "Wholesale Nail Supply": ("~10–50", "Private nail wholesale"),
    "CND (Creative Nail Design) distribution": ("~500–1,000 (brand)", "Revlon / CND professional nail"),
    "Gelish / Hand & Nail Harmony distributors": ("~200–500", "Private professional nail brand"),
    "Kiara Sky distribution": ("~100–500", "Private professional nail brand"),
    "Young Nails distribution": ("~50–200", "Private professional nail brand"),
    "Nail Systems International (NSI) distribution": ("~50–200", "Private professional nail brand"),
    "Light Elegance distribution": ("~50–200", "Private professional nail brand"),
    "Jessicacosmetics": ("~50–200", "Private nail brand"),
    "Chemistwarehouse": ("~5,000+", "Private AU pharmacy retail"),
    "Notino": ("~1,000+", "Private EU online beauty"),
    "Cult Beauty": ("~200–500", "THG online beauty"),
    "Lookfantastic / THG Beauty": ("~1,000+", "THG plc beauty"),
    "Beauty Bay": ("~200–500", "Online beauty retailer"),
    "Beauty Bay sheet masks": ("~200–500", "Online beauty retailer"),
    "Doctor Althea Global": ("~50–200", "Private K-beauty brand"),
    "Peach & Lily": ("~50–200", "Private K-beauty retailer/brand"),
    "Peach & Lily distribution": ("~50–200", "Private K-beauty retailer/brand"),
    "Soko Glam": ("~50–200", "Private K-beauty retailer"),
    "Sokoglam": ("~50–200", "Private K-beauty retailer"),
    "Soko Glam wholesale partners": ("~50–200", "Private K-beauty retailer"),
    "YesStyle": ("~500–1,000", "Private HK online fashion/beauty"),
    "YesStyle Wholesale inquiries / B2B": ("~500–1,000", "Private HK online fashion/beauty"),
    "Beautylish": ("~50–200", "Private online beauty"),
    "Beauty Bridge": ("~10–50", "Private online beauty"),
    "SkinStore": ("~100–500", "THG online skincare"),
    "Anand Agencies (Chennai)": ("~10–50", "Private India distributor"),
    "Jumia": ("~5,000+", "Public Africa e-commerce"),
    "Prosemic": ("~200–500", "Private China shunt resistor maker"),
    "Master Electronics": ("~200–500", "Private electronics distributor"),
    "Onlinecomponents.com": ("~200–500", "Master Electronics online"),
    "Distrelec": ("~500–1,000", "Private EU electronics distributor"),
    "Chip One Stop": ("~200–500", "Private Asia electronics distributor"),
    "Quest Components": ("~50–200", "Private components distributor"),
    "Classic Components": ("~50–200", "Private components distributor"),
    "Sourceability / Sourcengine": ("~200–500", "Private component sourcing"),
    "Findchips": ("~50–200", "Supply Frame parts aggregator"),
    "JLCPCB Parts / JiaLiChuang": ("~1,000+", "Private PCB / parts China"),
}


def main() -> None:
    n = 0
    for p in sorted(Path(r"D:\vendor-intel\output").rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(p)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=False))
        hdr_i = next(
            i
            for i, r in enumerate(rows[:12])
            if r and r[0].value and str(r[0].value).strip() == "Company"
        )
        header = [str(c.value or "").strip() for c in rows[hdr_i]]
        ci = header.index("Company")
        ei = header.index("Employees")
        ti = header.index("Turnover / Scale")
        for r in rows[hdr_i + 1 :]:
            name = str(r[ci].value or "").strip()
            if name not in EXTRA:
                continue
            emp, turn = EXTRA[name]
            cur_e = str(r[ei].value or "").lower()
            cur_t = str(r[ti].value or "").lower()
            if ("not " in cur_e) or ("sme" in cur_e) or (not cur_e.strip()):
                r[ei].value = emp
                n += 1
            if ("not " in cur_t) or ("sme" in cur_t) or (not cur_t.strip()):
                r[ti].value = turn
                n += 1
        wb.save(p)
        wb.close()
    print(f"extra scale cells updated: {n}")


if __name__ == "__main__":
    main()
