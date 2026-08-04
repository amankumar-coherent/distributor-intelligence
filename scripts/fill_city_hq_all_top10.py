"""Fill city-level HQ + Cities/Regions for remaining country-only rows across top-10 markets."""
from __future__ import annotations

from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

# company -> (Headquarters, Cities / Regions)
HQ_MAP: dict[str, tuple[str, str]] = {
    # --- Axillary crutches ---
    "Health Products For You": (
        "Brookfield, Connecticut, USA",
        "Brookfield, Connecticut, USA; USA online DME",
    ),
    "ProMed Products": (
        "USA",
        "USA medical equipment wholesale",
    ),
    "DME Supply USA": (
        "USA",
        "USA durable medical equipment supply",
    ),
    "Cascade Healthcare Solutions": (
        "Redmond, Washington, USA",
        "Redmond, Washington, USA; USA healthcare supply",
    ),
    "Allegro Medical": (
        "1833 West Main Street, Suite 131, Mesa, AZ 85201, USA",
        "Mesa, Arizona, USA; USA online medical supply",
    ),
    "Hugo Mobility distribution": (
        "USA",
        "USA retail and DME mobility channel",
    ),
    "Independence Medical (Cardinal specialty)": (
        "Dublin, Ohio, USA (Cardinal Health specialty)",
        "Dublin, Ohio, USA; USA specialty medical",
    ),
    "Concordance Healthcare Solutions": (
        "Tiffin, Ohio, USA",
        "Tiffin, Ohio, USA; Midwest and Southeast USA",
    ),
    "Direct Healthcare Group / UK DME": (
        "Caerphilly, Wales, United Kingdom",
        "Caerphilly, United Kingdom; UK; Europe",
    ),
    "Days Healthcare distribution": (
        "United Kingdom",
        "United Kingdom; export mobility channel",
    ),
    "Aidapt / UK mobility wholesale": (
        "St Helens, Merseyside, United Kingdom",
        "St Helens, United Kingdom; UK mobility wholesale",
    ),
    # --- Baijiu ---
    "Cnsimports": (
        "California, USA",
        "California, USA; 30+ state spirits distribution",
    ),
    "CNS Imports": (
        "California, USA",
        "California, USA; 30+ state spirits distribution",
    ),
    "Baijiu Huanjiu": (
        "Guizhou, China",
        "Guizhou, China; Zunyi and Renhuai baijiu regions",
    ),
    "Breakthru Beverage Group": (
        "3333 South Laramie Avenue, Cicero, IL 60804, USA",
        "Cicero, Illinois, USA; New York, New York, USA; multi-state USA; Canada",
    ),
    "Republic National Distributing Company (RNDC)": (
        "Houston, Texas, USA",
        "Houston, Texas, USA; USA multi-state wine and spirits",
    ),
    "Winebow Group": (
        "4800 Cox Road, Suite 300, Glen Allen, VA 23060, USA",
        "Glen Allen, Virginia, USA; USA fine wine wholesale",
    ),
    "Summergate Fine Wines & Spirits": (
        "Shanghai, China",
        "Shanghai, China; China on-trade and off-trade",
    ),
    "Duty Free Americas / travel retail spirits": (
        "Hollywood, Florida, USA",
        "Hollywood, Florida, USA; Americas travel retail",
    ),
    "DFS Group (travel retail)": (
        "Hong Kong",
        "Hong Kong; global travel retail",
    ),
    # --- BPV ---
    "Allied Medical": (
        "Sydney, New South Wales, Australia",
        "Sydney, Australia; Australia; New Zealand",
    ),
    "Cedar Gate Technologies": (
        "Greenwich, Connecticut, USA",
        "Greenwich, Connecticut, USA; USA health-tech",
    ),
    "Balanceback / Vestibular First channels": (
        "USA",
        "USA clinical vestibular channel",
    ),
    "DizzyDx / clinical vestibular product dist.": (
        "USA",
        "USA clinical vestibular channel",
    ),
    "Amplivox distributors": (
        "Oxfordshire, England, United Kingdom",
        "Oxfordshire, United Kingdom; UK; export",
    ),
    "Interacoustics USA / audiology dealers": (
        "Eden Prairie, Minnesota, USA",
        "Eden Prairie, Minnesota, USA; USA clinical dealers",
    ),
    "e3 Diagnostics": (
        "Schaumburg, Illinois, USA",
        "Schaumburg, Illinois, USA; USA audiology dealers",
    ),
    "Path Medical distributors": (
        "Germering, Bavaria, Germany",
        "Germering, Germany; Europe; global",
    ),
    "Homoth / vestibular product EU distributors": (
        "Germany",
        "Germany; Europe vestibular channel",
    ),
    "Synapsys / vestibular EU channel": (
        "Marseille, France",
        "Marseille, France; France; Europe",
    ),
    # --- Elderberry ---
    "Sundown Naturals": (
        "Boca Raton, Florida, USA (Nature's Bounty / Nestlé)",
        "Boca Raton, Florida, USA; USA mass retail",
    ),
    "Sundown Naturals Distribution": (
        "Boca Raton, Florida, USA (Nature's Bounty / Nestlé)",
        "Boca Raton, Florida, USA; USA mass retail",
    ),
    "Muckybootsfarm": (
        "USA",
        "USA farm and elderberry brand",
    ),
    "Hartmannsplantcompany": (
        "USA",
        "USA nursery and plant wholesale",
    ),
    "Hartmann's Plant Company": (
        "USA",
        "USA nursery and plant wholesale",
    ),
    "Wellevate / Prevention Wellness Dist.": (
        "USA",
        "USA practitioner wellness channel",
    ),
    "Natural Partners": (
        "Scottsdale, Arizona, USA",
        "Scottsdale, Arizona, USA; USA practitioner",
    ),
    "Rainbow Light / New Chapter Dist. (P&G/related)": (
        "USA",
        "USA natural and mass retail",
    ),
    "Medikonda Nutrients (bulk/OEM wholesale)": (
        "USA",
        "USA bulk B2B nutrients",
    ),
    "SMP Nutra": (
        "USA",
        "USA private-label nutraceutical B2B",
    ),
    "MCB Group (NuELDERBERRY ingredient)": (
        "Taiwan",
        "Taiwan; Asia; global ingredient B2B",
    ),
    # --- Instant noodles ---
    "Kingdavid": (
        "China",
        "China instant noodles OEM; global export",
    ),
    "Kingdavidfoods": (
        "China",
        "China; organic and instant noodles export",
    ),
    "JD.com supermarket / Jingdong": (
        "Beijing, China",
        "Beijing, China; China e-commerce grocery",
    ),
    "Alibaba / Tmall supermarket": (
        "Hangzhou, Zhejiang, China",
        "Hangzhou, Zhejiang, China; China e-commerce",
    ),
    "Rakuten / Japanese e-grocery channels": (
        "Tokyo, Japan",
        "Tokyo, Japan; Japan e-grocery",
    ),
    # --- MTB ---
    "Bicycle Warehouse": (
        "San Diego, California, USA",
        "San Diego, California, USA; USA bike retail",
    ),
    "Performance Bicycle": (
        "Chapel Hill, North Carolina, USA (legacy Performance Inc.)",
        "Chapel Hill, North Carolina, USA; USA mail-order and retail legacy",
    ),
    "Performance Inc. / related dealer channels": (
        "Chapel Hill, North Carolina, USA (legacy Performance Inc.)",
        "Chapel Hill, North Carolina, USA; USA dealer channel",
    ),
    "BikeCo / Performance Bicycle wholesale legacy": (
        "USA",
        "USA Performance Bicycle wholesale legacy",
    ),
    "Bikeonline": (
        "Italy",
        "Italy MTB and road components",
    ),
    "Wiggle": (
        "1000 Lakeside, Portsmouth, Hampshire PO6 3EN, United Kingdom",
        "Portsmouth, United Kingdom; UK; Europe online",
    ),
    "KHS Distribution / KHS Bicycles": (
        "Rancho Dominguez, California, USA",
        "Rancho Dominguez, California, USA; USA bike distribution",
    ),
    "DVO Suspension distributors": (
        "Huntington Beach, California, USA",
        "Huntington Beach, California, USA; USA; global MTB suspension",
    ),
    "Universal Cycles": (
        "Portland, Oregon, USA",
        "Portland, Oregon, USA; USA bike parts wholesale",
    ),
    "Worldwide Cyclery": (
        "3225 Grande Vista Drive, Thousand Oaks, CA, USA",
        "Thousand Oaks, California, USA; USA online MTB",
    ),
    "BikeDiscount / Mantel": (
        "Münsterland / Germany (Bike-Discount)",
        "Germany; Europe online bike retail",
    ),
    "BikesOnline / DTC AU-NZ": (
        "Australia",
        "Australia; New Zealand; USA DTC",
    ),
    "99 Bikes (AU retail chain)": (
        "Brisbane, Queensland, Australia",
        "Brisbane, Queensland, Australia; Australia retail",
    ),
    # --- Muslin ---
    "Bigduckcanvas": (
        "USA",
        "USA canvas and muslin wholesale",
    ),
    "BMF Fabric / Big Duck Canvas": (
        "USA",
        "USA canvas and muslin wholesale",
    ),
    "OnlineFabricStore": (
        "Springfield, Massachusetts, USA",
        "Springfield, Massachusetts, USA; USA nationwide",
    ),
    "Fabric Wholesale Direct": (
        "USA",
        "USA bulk fabric wholesale",
    ),
    "Fashion Fabrics Club": (
        "USA",
        "USA apparel fabric wholesale and retail",
    ),
    "Lunn Fabrics / Lunn Poly": (
        "USA",
        "USA theatrical muslin and scenic fabrics",
    ),
    "Ripstop by the Roll": (
        "USA",
        "USA online technical textiles",
    ),
    "Global Sources Fabrics": (
        "Hong Kong",
        "Hong Kong; Asia export; global buyers",
    ),
    "Canvas Etc.": (
        "USA",
        "USA canvas and industrial fabrics",
    ),
    "Fabriclore": (
        "New Delhi, India",
        "New Delhi, India; India; global apparel brands",
    ),
    # --- Nail polish ---
    "Chemistwarehouse": (
        "Preston, Victoria, Australia",
        "Preston, Victoria, Australia; Australia pharmacy retail",
    ),
    "Jessicacosmetics": (
        "Los Angeles, California, USA",
        "Los Angeles, California, USA; USA professional nail",
    ),
    "Nail Systems International (NSI) distribution": (
        "Pennsylvania, USA",
        "Pennsylvania, USA; global professional nail",
    ),
    "CND (Creative Nail Design) distribution": (
        "9560 Towne Centre Drive, San Diego, CA 92121, USA",
        "San Diego, California, USA; global professional",
    ),
    "Gelish / Hand & Nail Harmony distributors": (
        "Brea, California, USA",
        "Brea, California, USA; USA; global",
    ),
    "Young Nails distribution": (
        "Anaheim, California, USA",
        "Anaheim, California, USA; USA; export",
    ),
    "Light Elegance distribution": (
        "USA",
        "USA; international professional nail",
    ),
    "Kiara Sky distribution": (
        "Buena Park, California, USA",
        "Buena Park, California, USA; USA; global e-commerce",
    ),
    "Burmax Beauty Distributors": (
        "Holtsville, New York, USA",
        "Holtsville, New York, USA; USA salon wholesale",
    ),
    "Marianna Industries distribution": (
        "Omaha, Nebraska, USA",
        "Omaha, Nebraska, USA; USA beauty wholesale",
    ),
    "Beauty Bay": (
        "Manchester, England, United Kingdom",
        "Manchester, United Kingdom; UK; Europe; online global",
    ),
    "Cult Beauty": (
        "London, England, United Kingdom",
        "London, United Kingdom; UK; Europe",
    ),
    "Born Pretty wholesale": (
        "China",
        "China; global e-commerce nail wholesale",
    ),
    "Makartt wholesale": (
        "USA",
        "USA; global online nail supplies",
    ),
    "Modelones distribution": (
        "USA",
        "USA; global online nail",
    ),
    "Wholesale Nail Supply": (
        "USA",
        "USA online professional nail wholesale",
    ),
    # --- Sheet masks ---
    "Beauty Bridge": (
        "USA",
        "USA online beauty retail",
    ),
    "Peach & Lily": (
        "New York, New York, USA",
        "New York, New York, USA; USA; Ulta and Sephora channels",
    ),
    "Peach & Lily distribution": (
        "New York, New York, USA",
        "New York, New York, USA; USA; Ulta and Sephora channels",
    ),
    "SkinStore": (
        "USA (THG / Lookfantastic group channel)",
        "USA online skincare",
    ),
    "YesStyle": (
        "Hong Kong",
        "Hong Kong; global online K-beauty",
    ),
    "YesStyle Wholesale inquiries / B2B": (
        "Hong Kong",
        "Hong Kong; global online",
    ),
    "Beautylish": (
        "San Francisco, California, USA",
        "San Francisco, California, USA; USA online",
    ),
    "Guardian / AS Watson": (
        "Singapore",
        "Singapore; Southeast Asia drugstore",
    ),
    "Superdrug": (
        "Croydon, England, United Kingdom",
        "Croydon, United Kingdom; UK retail",
    ),
    "Beauty Bay sheet masks": (
        "Manchester, England, United Kingdom",
        "Manchester, United Kingdom; UK; Europe",
    ),
    "Watsons / AS Watson beauty retail": (
        "Hong Kong",
        "Hong Kong; Asia; Europe (AS Watson)",
    ),
    # --- Shunt ---
    "Source": (
        "China",
        "China; global electronic component sourcing",
    ),
    "Electropeak": (
        "Online (international)",
        "Online international electronics",
    ),
    "Onlinecomponents.com": (
        "USA (Master Electronics channel)",
        "USA; global online components",
    ),
    "Distrelec": (
        "Nänikon, Zurich, Switzerland",
        "Nänikon, Switzerland; Europe electronics distribution",
    ),
    "Quest Components": (
        "USA",
        "USA electronic components",
    ),
    "Classic Components": (
        "Torrance, California, USA",
        "Torrance, California, USA; USA; global stocking",
    ),
    "Sourceability / Sourcengine": (
        "USA",
        "USA; global component sourcing",
    ),
    "Findchips": (
        "USA (Supply Frame / DigiKey ecosystem)",
        "USA; global parts aggregation",
    ),
}


def process(path: Path) -> int:
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
    hi = header.index("Headquarters")
    ri = header.index("Cities / Regions")
    n = 0
    for r in rows[hdr_i + 1 :]:
        name = str(r[ci].value or "").strip()
        if name not in HQ_MAP:
            continue
        hq, cities = HQ_MAP[name]
        # Always apply city-level upgrades (even if previous country-only)
        if str(r[hi].value or "").strip() != hq or str(r[ri].value or "").strip() != cities:
            r[hi].value = hq
            r[ri].value = cities
            n += 1
    if "Run Summary" in wb.sheetnames:
        wb["Run Summary"]["A13"] = f"City-level HQ fill applied: {n} rows"
    wb.save(path)
    wb.close()
    return n


def export_summary() -> Path:
    out = Path(r"D:\vendor-intel\output\top10_founded_hq_cities_summary.xlsx")
    wb = openpyxl.Workbook()
    # overview
    ws0 = wb.active
    ws0.title = "Overview"
    ws0.append(["Market", "Companies", "With Founded", "With HQ", "With Cities"])
    detail = wb.create_sheet("All Companies")
    detail.append(["Market", "Company", "Founded", "Headquarters", "Cities / Regions", "Website"])

    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        market = path.parent.name.replace("_market_global", "").replace("_", " ").title()
        wb_src = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb_src[wb_src.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        hdr_i = next(i for i, r in enumerate(data[:12]) if r and str(r[0] or "").strip() == "Company")
        header = [str(h or "").strip() for h in data[hdr_i]]
        fi = header.index("Founded")
        hi = header.index("Headquarters")
        ri = header.index("Cities / Regions")
        wi = header.index("Website") if "Website" in header else None
        rows = [r for r in data[hdr_i + 1 :] if r and r[0]]
        with_f = sum(1 for r in rows if str(r[fi] or "").strip())
        with_h = sum(1 for r in rows if str(r[hi] or "").strip())
        with_c = sum(1 for r in rows if str(r[ri] or "").strip())
        ws0.append([market, len(rows), with_f, with_h, with_c])
        for r in rows:
            detail.append(
                [
                    market,
                    r[0],
                    r[fi],
                    r[hi],
                    r[ri],
                    r[wi] if wi is not None else "",
                ]
            )
        wb_src.close()

        # per-market sheet
        safe = market[:28]
        wsm = wb.create_sheet(safe)
        wsm.append(["Company", "Founded", "Headquarters", "Cities / Regions"])
        for r in rows:
            wsm.append([r[0], r[fi], r[hi], r[ri]])

    wb.save(out)
    wb.close()
    return out


def main() -> None:
    total = 0
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        n = process(path)
        total += n
        print(f"{path.parent.name}: updated {n}")
    print(f"TOTAL updated {total}")
    out = export_summary()
    print(f"Summary written: {out}")

    # remaining country-only
    COUNTRY = {
        "USA",
        "UK",
        "UNITED KINGDOM",
        "INDIA",
        "CHINA",
        "GERMANY",
        "FRANCE",
        "JAPAN",
        "AUSTRALIA",
        "CANADA",
        "SWEDEN",
        "NETHERLANDS",
        "TURKEY",
        "HONG KONG",
        "TAIWAN",
        "SINGAPORE",
        "ITALY",
    }
    print("\nStill country-only HQ:")
    left = 0
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        hdr_i = next(i for i, r in enumerate(data[:12]) if r and str(r[0] or "").strip() == "Company")
        header = [str(h or "").strip() for h in data[hdr_i]]
        hi = header.index("Headquarters")
        for r in data[hdr_i + 1 :]:
            if not r or not r[0]:
                continue
            hq = str(r[hi] or "").strip()
            if hq.upper() in COUNTRY:
                left += 1
                print(f"  {path.parent.name}|{r[0]}|{hq}")
        wb.close()
    print(f"Remaining country-only: {left}")


if __name__ == "__main__":
    main()
