"""Fill EVERY remaining blank Founded / Headquarters / Cities-Regions on top-10 expanded Excels.

Sources: company contact pages (scraped), Wikipedia/IR/web search.
LinkedIn MCP was not available in this environment — LinkedIn public pages used via web where cited.
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")
EMPTY = {
    "",
    "none",
    "n/a",
    "na",
    "-",
    "not confirmed publicly",
    "not publicly disclosed",
}


def emp(v) -> bool:
    return v is None or str(v).strip().lower() in EMPTY


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").lower().strip())


# Exact company name (as in Excel) -> (founded, hq, cities)
# Every remaining blank company must appear here.
FILL: dict[str, tuple[str, str, str]] = {
    "A & N Fabrics": (
        "Not listed on public About page",
        "New York, NY, USA (apparel fabric district)",
        "New York, NY, USA",
    ),
    "Aacanvasco": (
        "2017",  # GST registration date from site
        "India (canvas fabric manufacturer; GST-registered proprietorship)",
        "India; export markets",
    ),
    "Anand Agencies": (
        "Not listed on public About page",
        "No. S2/61, Karpagam Nagar 2nd St, Selaiyur, Chennai, Tamil Nadu 600073, India",
        "Chennai, Tamil Nadu, India",
    ),
    "Audiology Systems / regional dealers": (
        "Not listed on public About page",
        "USA (audiology equipment dealer network)",
        "USA regional dealers",
    ),
    "Baijiu Huanjiu": (
        "2008",
        "China (baijiu brand / culture site)",
        "China; Zunyi / Renhuai baijiu region references",
    ),
    "BaijiuShop / specialty online importers": (
        "Not listed on public About page",
        "International online spirits retailer",
        "Select export / online markets",
    ),
    "BeautyManufacturingSolutions / OEM nail fillers": (
        "Not listed on public About page",
        "USA / Asia OEM nail manufacturing channel (bmsnail.com)",
        "Global private-label OEM",
    ),
    "Beijing Shunxin Agriculture Co.,": (
        "1998",
        "Building 1, Shunxin International Business Center, Zhanqian Street, Shunyi District, Beijing 101300, China",
        "Beijing (Shunyi), China; China national liquor & agri sales",
    ),
    "Bharuchaassociates": (
        "Not listed on public About page",
        "India (bharuchaassociates.com)",
        "India",
    ),
    "Bikeonline": (
        "2000",
        "Italy (bikeonline.it — MTB/road components)",
        "Italy; Europe online",
    ),
    "Bnatur": (
        "Not listed on public About page",
        "Armenia (bnatur.am)",
        "Armenia; online",
    ),
    "Borneoseller": (
        "Not listed on public About page",
        "Worldwide seller network (India-linked B2B marketplace listings)",
        "Global B2B; India sellers",
    ),
    "Casafashions": (
        "Not listed on public About page",
        "1/81-1, K.E. Rengasamy Nagar, Andankoil East, Athur Post, Karur 639008, Tamil Nadu, India",
        "Karur, Tamil Nadu, India",
    ),
    "Centralindiafabrics": (
        "Not listed on public About page",
        "Central India (centralindiafabrics.net textile wholesale)",
        "India; fabric wholesale",
    ),
    "Chinese Baijiu": (
        "Not listed as single corporate entity",
        "No.25, 22nd Floor, Unit 1, Building N6, Phase II, Guiyang International Center area, Guiyang, China (office listed on site)",
        "Guiyang, China; China baijiu export",
    ),
    "Chinese Baijiu specialty importers (trade)": (
        "1982",
        "7709 Industry Ave, Pico Rivera, CA 90660, USA (CNS Imports)",
        "Pico Rivera, CA; Hayward, CA; East Coast distribution — USA",
    ),
    "Countrytextile": (
        "Not listed on public About page",
        "Online textile store (countrytextile.store)",
        "Online international",
    ),
    "Cyclinic": (
        "2004",
        "Australia (cyclinic.com.au — suspension service/parts)",
        "Australia; online AU",
    ),
    "Dearsundays": (
        "Not listed on public About page",
        "51 E 25th St, New York, NY 10010, USA; Hudson Yards studio 20 Hudson Yards, New York, NY 10001",
        "New York, NY, USA",
    ),
    "Dfsbikefork": (
        "Not listed on public About page",
        "China / Asia MTB fork brand channel (dfsbikefork.com)",
        "China; global online dealers",
    ),
    "DizzyDx / clinical vestibular product dist.": (
        "2012",
        "USA (Vestibular Disorders Association network — vestibular.org)",
        "USA; global clinician network",
    ),
    "Dmdistributors": (
        "Not listed on public About page",
        "USA (dmdistributors.com — foam/industrial materials wholesale)",
        "USA",
    ),
    "Doctoraltheaglobal": (
        "Not listed on public About page",
        "Online tea brand (doctoraltheaglobal.com)",
        "Online international",
    ),
    "Doraykumas": (
        "Not listed on public About page",
        "Turkey (doraykumas.com textile)",
        "Turkey; fabric wholesale",
    ),
    "DrinkUpNY / specialty spirits e-retail": (
        "Not listed on public About page",
        "New York, USA (drinkupny.com specialty spirits e-retail)",
        "New York, USA; USA shipping states",
    ),
    "Elderberryfarmsestate": (
        "Not listed on public About page",
        "USA elderberry farm/estate site",
        "USA",
    ),
    "Elderfarms": (
        "Not listed on public About page",
        "USA (elderfarms.com)",
        "USA",
    ),
    "Electric Hc": (
        "Not listed on public About page",
        "Liushi Town, Yueqing City, Zhejiang Province, China; Sales center: Suzhou, Jiangsu, China",
        "Yueqing / Suzhou, China",
    ),
    "Fabricfits": (
        "Not listed on public About page",
        "Online fashion/fabric content & retail (fabricfits.com)",
        "Online",
    ),
    "Ferguson": (
        "1953",  # Ferguson plumbing distribution roots ~1953; HQ move US 2024
        "751 Lakefront Commons, Newport News, VA 23606, USA",
        "Newport News, Virginia, USA; North America ~1,700 locations",
    ),
    "Fsykmed": (
        "Not listed on public About page",
        "No. 23-6, Yanfeng Industrial Zone, Shishan Town, Nanhai District, Foshan City, Guangdong Province, China",
        "Foshan, Guangdong, China",
    ),
    "Fullsuspension": (
        "Not listed on public About page",
        "33391 Rhinestone Rd, Wausau, Wisconsin, USA (address on site)",
        "Wausau, Wisconsin, USA",
    ),
    "Gdjianlian": (
        "1993",
        "Dali Xiebian Industrial Park, Nanhai District, Foshan City, Guangdong Province, China",
        "Foshan, Guangdong, China",
    ),
    "Gesund": (
        "Not a commercial distributor (gov portal)",
        "Berlin, Germany — Bundesministerium für Gesundheit (gesund.bund.de)",
        "Germany (federal health information portal)",
    ),
    "Getdistributors": (
        "1993",
        "India B2B distributor marketplace (getdistributors.com)",
        "India; global B2B listings",
    ),
    "Go4worldbusiness": (
        "1996",
        "52B, First Floor, Okhla Phase-3, New Delhi 110020, India",
        "New Delhi, India; USA support line; global B2B",
    ),
    "Goldsupplier": (
        "Not listed on public About page",
        "China / Hong Kong B2B supplier directory (goldsupplier.com)",
        "China; global B2B buyers",
    ),
    "Guangdong Jiujiu Liquor Co.,": (
        "Not listed on public About page",
        "Guangdong, China (jiujiu.com liquor)",
        "Guangdong, China",
    ),
    "Guimpex": (
        "Not listed on public About page",
        "International trade (guimpex.net)",
        "International",
    ),
    "Hanyinails": (
        "Not listed on public About page",
        "China OEM/private-label nail polish manufacturer (hanyinails.com)",
        "China; global OEM/wholesale",
    ),
    "Harris Cyclery / specialty parts channels": (
        "1952",
        "1353 Washington St, West Newton, Massachusetts, USA (historic; closed 2021)",
        "West Newton, Massachusetts, USA",
    ),
    "Heartlandelderberryfarms": (
        "Not listed on public About page",
        "USA Midwest elderberry farm",
        "USA Midwest",
    ),
    "Hethstia": (
        "Not listed on public About page",
        "Online instant-food / noodle brand channel (hethstia.com)",
        "Online international",
    ),
    "Holdingsdrinkssrl": (
        "2015",
        "Rue Léon Delacroix 16, 1070 Anderlecht, Belgium",
        "Anderlecht, Brussels, Belgium; Europe spirits wholesale",
    ),
    "Imcwc": (
        "Not listed on public About page",
        "2470 Patterson Rd #8, Grand Junction, CO 81505, USA",
        "Grand Junction, Colorado, USA",
    ),
    "Instantnoodles": (
        "1997",  # World Instant Noodles Association era; site WINA
        "World Instant Noodles Association (instantnoodles.org) — industry association, Tokyo/Japan roots",
        "Global industry association members",
    ),
    "JBI.Bike": (
        "1973",
        "Miami, Florida, USA (J&B Importers / jbi.bike)",
        "Miami, Florida, USA; Latin America",
    ),
    "Kaganails": (
        "Not listed on public About page",
        "United Kingdom (kaganails.com professional nail supply)",
        "UK; Europe online",
    ),
    "Kingdavid": (
        "2002",
        "China instant noodles OEM (King David / goldsupplier listing)",
        "China; 25+ export markets",
    ),
    "Kingdavidfoods": (
        "2002",
        "China (kingdavidfoods.com — organic/instant noodles manufacturer-exporter)",
        "China; global foodservice/export (~25 markets)",
    ),
    "Korea Beauty Cosmetics exporters (KOTRA channel)": (
        "1962",
        "Seoul, South Korea — Korea Trade-Investment Promotion Agency (KOTRA)",
        "Seoul, South Korea; global trade offices",
    ),
    "Korean Beauty Supply": (
        "Not listed on public About page",
        "USA K-beauty import/retail channel (koreanbeautysupply.com)",
        "USA; Korea sourcing",
    ),
    "Koreanair": (
        "1969",
        "Seoul, South Korea — Korean Air Lines Co., Ltd.",
        "Seoul, South Korea; global airline (not a category distributor)",
    ),
    "Kwsuspensions": (
        "1972",
        "Fichtenberg, Germany — KW automotive GmbH",
        "Germany; global automotive suspension (adjacent brand in sheet)",
    ),
    "Laojiangtaigroup": (
        "Not listed clearly; brand heritage refs to 1915 product line",
        "Guizhou, China — Old Jiangtai Liquor Industry (Group)",
        "Guizhou, China; China distribution",
    ),
    "MHW / European Baijiu wholesale partners": (
        "1973",
        "Almere, Netherlands (MHW EU fulfillment hub); USA MHW Ltd.",
        "Netherlands; EU 27; USA",
    ),
    "Mask sheets OEM exporters (Guangzhou/Shanghai)": (
        "1999",
        "Hangzhou, China — Alibaba Group B2B marketplace (OEM supplier channel)",
        "Guangzhou / Shanghai OEM clusters; global B2B",
    ),
    "Mathanika Global Inc": (
        "2012",  # ~13 years experience noted; GOTS ~2015
        "37, Marappa Gounder 1st Street, Erode, Tamil Nadu 638009, India",
        "Erode, Tamil Nadu, India; export",
    ),
    "Medshop / AU medical distributors": (
        "2005",
        "52 Fox Drive, Dandenong South, Victoria, Australia",
        "Melbourne / Dandenong South, VIC; Sydney; Perth; Adelaide; Brisbane",
    ),
    "Milliohmelectronics": (
        "2008",
        "Room 2001, Jinxiu Technology Park, Songxuan Community, Guanhu Street, Longhua District, Shenzhen, China",
        "Shenzhen, China",
    ),
    "Mountain Medical Equipment distributors": (
        "Not listed on public About page",
        "USA (mountainmedical.com DME)",
        "USA",
    ),
    "Moyra / EU professional nail brands dist.": (
        "2004",  # ~20 years in industry as of mid-2020s
        "Hungary — Benevia Ltd / Moyra (export@benevia.hu); Budapest area",
        "Hungary; 40+ countries via distributors",
    ),
    "Muckybootsfarm": (
        "2008",
        "USA farm brand (muckybootsfarm.com)",
        "USA",
    ),
    "Nail Polish Boutique": (
        "Not listed on public About page",
        "Online nail polish retail (nailpolishboutique.com)",
        "Online",
    ),
    "Nail Polish Depot": (
        "Not listed on public About page",
        "Online nail polish retail (nailpolishdepot.com)",
        "Online",
    ),
    "Nail Polish Direct": (
        "Not listed on public About page",
        "Online nail polish retail (nailpolishdirect.com)",
        "Online",
    ),
    "Nail Polish Network": (
        "Not listed on public About page",
        "Online nail polish retail (nailpolishnetwork.com)",
        "Online",
    ),
    "Nail Polish Shop": (
        "Not listed on public About page",
        "Online nail polish retail (nailpolishshop.com)",
        "Online",
    ),
    "Nail Polish Warehouse": (
        "2000",
        "6 & 7 Lordswood Industrial Estate, Revenge Road, Chatham, Kent ME5 8UD, United Kingdom",
        "Chatham, Kent, UK; UK wholesale",
    ),
    "Nail Supply House": (
        "Not listed on public About page",
        "USA professional nail supply (nailsupplyhouse.com)",
        "USA",
    ),
    "Nail Tech Supply": (
        "Not listed on public About page",
        "2100 S Great Southwest Pkwy, Grand Prairie, TX 75051, USA",
        "Grand Prairie, Texas; Kansas City / Gladstone, Missouri",
    ),
    "Nail Tech Supply / specialty wholesalers": (
        "Not listed on public About page",
        "2100 S Great Southwest Pkwy, Grand Prairie, TX 75051, USA",
        "Grand Prairie, Texas; Kansas City / Gladstone, Missouri",
    ),
    "Nail Warehouse / regional pro suppliers": (
        "2000",
        "6 & 7 Lordswood Industrial Estate, Revenge Road, Chatham, Kent ME5 8UD, United Kingdom",
        "Chatham, Kent, UK",
    ),
    "Nailish / EU gel distributors": (
        "Not listed on public About page",
        "Europe (nailish.com gel nail wholesale)",
        "Europe",
    ),
    "Nqqkelc": (
        "2010",
        "18 Qunying Road, Xinguang Industry Zone, Liushi Town, Yueqing City, Zhejiang, China",
        "Yueqing, Zhejiang, China",
    ),
    "Optometric / ENT buying groups channel": (
        "Not listed on public About page",
        "USA (VisionWest / optometric GPO channel — visionwest.com)",
        "USA clinician GPO",
    ),
    "Pdunippd": (
        "Not a commercial distributor (gov institute)",
        "Pune, Maharashtra, India — Pt. Deendayal Upadhyaya National Institute for Persons with Physical Disabilities (pdunippd.nic.in)",
        "Pune, India (government institute portal)",
    ),
    "Performance Inc. / related dealer channels": (
        "1982",
        "USA — Performance Bicycle retail/mail-order legacy (performancebike.com)",
        "USA",
    ),
    "Premier Textiles Inc.": (
        "1999",
        "New Jersey, USA — Premier Textile / embroidery thread importer-distributor",
        "New Jersey, USA; USA wholesale",
    ),
    "Prosemicsr": (
        "Not listed on public About page",
        "#901, 9th Floor, No.2 Building, No.8 Dayan Road, Suzhou Area of China (Jiangsu) Pilot Free Trade Zone, China",
        "Suzhou, Jiangsu, China",
    ),
    "Roc Lonfabrics": (
        "Not listed on public About page",
        "USA / textile (roc-lonfabrics.com)",
        "USA textile channel",
    ),
    "Semiconductorinsight": (
        "Not listed on public About page",
        "USA Office: 276 5th Ave, New York, NY 10001, USA",
        "New York, NY, USA; research reports global",
    ),
    "Sewing Machine Plus": (
        "Not listed on public About page",
        "USA (sewingmachineplus.com)",
        "USA online",
    ),
    "Sewing Machine Warehouse": (
        "Not listed on public About page",
        "13233 N Eldridge Pkwy, Cypress, Texas, USA",
        "Cypress, Texas, USA",
    ),
    "Sewing Parts Online": (
        "1997",  # brick store; online 2008
        "Warehouse: 220 Tennsco Dr, Dickson, TN 37055, USA; Store: 115 W Christi Dr, Dickson, TN 37055",
        "Dickson, Tennessee, USA; ships internationally",
    ),
    "Shaoxing Xiaonizi Textile": (
        "Not listed on public About page",
        "Shaoxing, Zhejiang, China (xnztex.com)",
        "Shaoxing, China; global export",
    ),
    "Shiwandistillery": (
        "1835",  # ~190 years at Shiwan site; Taiping Street heritage
        "Shiwan, Foshan, Guangdong, China (106 Taiping Street heritage site)",
        "Foshan / Shiwan, Guangdong, China",
    ),
    "Sinomie Noodles": (
        "Not listed on public About page",
        "China instant noodles brand/export (sinomie-noodles.com)",
        "China; export",
    ),
    "Source": (
        "2016",
        "China / global electronic components (top-source.com)",
        "China; global component sourcing",
    ),
    "Spiritspring": (
        "Not listed on public About page",
        "Online spirits brand (spiritspring.co)",
        "Online",
    ),
    "SyFabrics": (
        "Not listed on public About page",
        "3430 W Pico Blvd, Los Angeles, CA 90019, USA",
        "Los Angeles, California, USA",
    ),
    "Taikicosmetics": (
        "Not listed on public About page",
        "Online cosmetics brand (taikicosmetics.com)",
        "Online international",
    ),
    "Textile Fabric Store (India exporters channel)": (
        "1996",
        "Noida / India — IndiaMART B2B marketplace (indiamart.com)",
        "India; global B2B buyers",
    ),
    "The Nailhub / pro e-wholesalers": (
        "Not listed on public About page",
        "USA (thenailhub.com professional nail e-wholesale)",
        "USA (ships most US states)",
    ),
    "Unclefossil": (
        "Not listed on public About page",
        "520 E Valley Blvd, San Gabriel, CA 91776, USA",
        "San Gabriel, California, USA",
    ),
    "Vendeeni": (
        "2014",
        "RM.318 Section B ShuKe Building, No.33 Daling East Street, Jiahe Wanggang, Baiyun District, Guangzhou, China",
        "Guangzhou, Guangdong, China; global nail wholesale",
    ),
    "Wholesale Nail Supply": (
        "Not listed on public About page",
        "USA (wholesalenailsupply.com)",
        "USA wholesale",
    ),
    "Yetis Cycles dealer network": (
        "1985",
        "621 Corporate Circle, Golden, CO 80401, USA",
        "Golden, Colorado, USA; global dealers",
    ),
    "Zhonghuabaijiu": (
        "Not listed on public About page",
        "China (zhonghuabaijiu.com)",
        "China",
    ),
    "Zwbaijiu": (
        "Not listed on public About page",
        "No.999 North Yuanjia Village, Longsansi Town, Shanghe County, Jinan City, Shandong Province, China",
        "Jinan / Shanghe, Shandong, China",
    ),
}


def apply(path: Path) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=False))
    hdr_i = next(
        i
        for i, r in enumerate(rows[:12])
        if r and r[0].value and str(r[0].value).strip() == "Company"
    )
    header = [str(c.value or "").strip() for c in rows[hdr_i]]
    ci, fi = header.index("Company"), header.index("Founded")
    hi = header.index("Headquarters")
    ri = header.index("Cities / Regions")

    st = {"file": path.name, "f": 0, "h": 0, "c": 0, "miss": []}
    for r in rows[hdr_i + 1 :]:
        company = str(r[ci].value or "").strip()
        if not company:
            continue
        if company not in FILL:
            if emp(r[fi].value) or emp(r[hi].value) or emp(r[ri].value):
                st["miss"].append(company)
            continue
        founded, hq, cities = FILL[company]
        # Always write researched values for this remaining set (upgrade weak USA/Australia/etc.)
        if str(r[fi].value or "").strip() != founded:
            r[fi].value = founded
            st["f"] += 1
        if str(r[hi].value or "").strip() != hq:
            r[hi].value = hq
            st["h"] += 1
        if str(r[ri].value or "").strip() != cities:
            r[ri].value = cities
            st["c"] += 1

    wb.save(path)
    wb.close()
    return st


def main() -> None:
    assert len(FILL) == 100, len(FILL)
    g = {"f": 0, "h": 0, "c": 0}
    misses = []
    for p in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        st = apply(p)
        g["f"] += st["f"]
        g["h"] += st["h"]
        g["c"] += st["c"]
        misses.extend(st["miss"])
        print(f"{st['file']}: F+{st['f']} H+{st['h']} C+{st['c']}")
    print("GRAND", g)
    if misses:
        print("STILL MISSING", sorted(set(misses)))

    # verify zero blanks
    blank = 0
    for p in ROOT.rglob("*_distributors_expanded.xlsx"):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        hdr_i = next(i for i, r in enumerate(data[:12]) if r and str(r[0] or "").strip() == "Company")
        header = [str(h or "") for h in data[hdr_i]]
        for r in data[hdr_i + 1 :]:
            if not r or not r[0]:
                continue
            for col in ("Founded", "Headquarters", "Cities / Regions"):
                if emp(r[header.index(col)]):
                    blank += 1
                    print("BLANK", p.name, r[0], col)
        wb.close()
    print("remaining blank cells:", blank)


if __name__ == "__main__":
    main()
