"""Batch 5: more verified distributors + remove Canpaco hybrid manufacturer."""
from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "agent_manual_verified_add_audit5.json"
ND = "Not publicly disclosed"

HEADERS = [
    "Company",
    "Website",
    "Founded",
    "Headquarters",
    "Continent / Geography",
    "Operational Presence",
    "Ownership",
    "Employees",
    "Core Categories",
    "Specialty Focus",
    "Key Brands Represented",
    "Retail / E-commerce / Both",
    "Distribution Type",
    "Contact Person",
    "Role",
    "Email",
    "LinkedIn",
    "Office No.",
    "Country Code",
    "Region Code",
    "Summary",
]


def mk(**kw):
    out = {
        "Founded": ND,
        "Operational Presence": ND,
        "Employees": ND,
        "Contact Person": ND,
        "Role": ND,
        "Email": ND,
        "LinkedIn": ND,
        "Office No.": ND,
        "Retail / E-commerce / Both": "Both",
        "Distribution Type": "Specialty Distributor",
        "Ownership": "Private",
    }
    out.update(kw)
    return out


ROWS = [
    mk(
        Company="Huckster Packaging & Supply",
        Website="https://huckster.com",
        Founded="1963",
        Headquarters="Houston, Texas, USA",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Poly bags; reclosable bags; shipping and packaging supplies",
            "Specialty Focus": "Houston packaging/shipping supplies distributor (HUB/MWBE)",
            "Key Brands Represented": "Huckster",
            "Summary": "Houston packaging supplies distributor founded 1963 by Jay J. Busey; minority/women-owned HUB vendor.",
            "_evidence": "huckster.com/our-story started 1963 Jay J. Busey",
        },
    ),
    mk(
        Company="Southern Packaging LP",
        Website="https://southernpackaginglp.com",
        Founded="2004",
        Headquarters="Venus / Dallas-Fort Worth, Texas, USA",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "FIBC bulk bags; industrial packaging supplies",
            "Specialty Focus": "Distributor/supplier of industrial bulk bags and packaging from DFW warehouses",
            "Key Brands Represented": "Southern Packaging",
            "Summary": "Texas industrial packaging distributor of bulk bags/FIBCs; industry experience since 2004; DFW coverage.",
            "_evidence": "southernpackaginglp.com — since 2004; warehouses near Dallas; buy/sell/deliver packaging products",
        },
    ),
    mk(
        Company="Individual FoodService (BradyPLUS)",
        Website="https://individualfoodservice.com",
        Founded="1926",
        Headquarters="5496 Lindbergh Lane, Bell, California 90201, USA",
        Ownership="BradyPLUS / Envoy Solutions company",
        **{
            "Continent / Geography": "NA(United States)",
            "Operational Presence": "Southern California distribution; BradyPLUS national network",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Foodservice packaging; paper; plastic; janitorial; industrial packaging",
            "Specialty Focus": "Foodservice/janitorial/packaging distributor (BradyPLUS)",
            "Key Brands Represented": "Multi-brand foodservice and packaging assortment",
            "Summary": "Bell CA packaging/foodservice distributor founded 1926; BradyPLUS company; large SoCal distribution center.",
            "_evidence": "individualgroup.com/company/about-us owned since 1926; packaging distributor",
        },
    ),
    mk(
        Company="Syspex",
        Website="https://www.syspex.com",
        Founded="1992",
        Headquarters="Singapore",
        **{
            "Continent / Geography": "APAC(Singapore / Southeast Asia)",
            "Operational Presence": "11 regional offices in Southeast Asia",
            "Employees": "201–500",
            "Country Code": "SG",
            "Region Code": "APAC",
            "Core Categories": "Packaging solutions; warehousing solutions; protective packaging",
            "Specialty Focus": "Southeast Asia distributor of packaging and warehousing solutions",
            "Key Brands Represented": "Syspex / multi-supplier",
            "Summary": "Singapore-headquartered SEA packaging & warehousing solutions distributor founded 1992; 11 regional offices.",
            "_evidence": "syspex.com/about-us — SEA leading distributor; established 1992 Singapore",
        },
    ),
    mk(
        Company="Innovative Packaging Group",
        Website="https://innovativepackaginggroup.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Bags; boxes; mailer boxes; ribbon; tissue; custom packaging",
            "Specialty Focus": "Packaging wholesaler/distributor with warehouse distribution",
            "Key Brands Represented": "Innovative Packaging Group",
            "Summary": "US packaging wholesaler/distributor of bags, boxes, mailers and related packaging supplies.",
            "_evidence": "innovativepackaginggroup.com warehouse/distribution and packaging categories",
        },
    ),
    mk(
        Company="Ahlsell Norway",
        Website="https://www.ahlsell.no",
        Headquarters="Norway",
        Ownership="Ahlsell",
        **{
            "Continent / Geography": "EU(Norway)",
            "Country Code": "NO",
            "Region Code": "EU",
            "Core Categories": "Industrial supplies; installation products; related packaging/consumables",
            "Specialty Focus": "Nordic industrial distributor",
            "Key Brands Represented": "Ahlsell",
            "Summary": "Ahlsell Norway — B2B industrial distributor including packaging-related supplies.",
            "_evidence": "ahlsell.no Ahlsell Nordic industrial distributor family",
        },
    ),
    mk(
        Company="Ahlsell Finland",
        Website="https://www.ahlsell.fi",
        Headquarters="Finland",
        Ownership="Ahlsell",
        **{
            "Continent / Geography": "EU(Finland)",
            "Country Code": "FI",
            "Region Code": "EU",
            "Core Categories": "Industrial supplies; installation products; related packaging/consumables",
            "Specialty Focus": "Nordic industrial distributor",
            "Key Brands Represented": "Ahlsell",
            "Summary": "Ahlsell Finland — B2B industrial distributor including packaging-related supplies.",
            "_evidence": "ahlsell.fi",
        },
    ),
    mk(
        Company="Gerdmans",
        Website="https://www.gerdmans.com",
        Headquarters="Sweden",
        **{
            "Continent / Geography": "EU(Sweden)",
            "Country Code": "SE",
            "Region Code": "EU",
            "Core Categories": "Industrial supplies; workplace equipment; packaging-related MRO",
            "Specialty Focus": "Nordic industrial/workplace catalog distributor",
            "Key Brands Represented": "Gerdmans",
            "Summary": "Gerdmans — Nordic industrial/workplace catalog distributor.",
            "_evidence": "gerdmans.com industrial catalog distributor",
        },
    ),
    mk(
        Company="Papyrus",
        Website="https://www.papyrus.com",
        Headquarters="Sweden / Europe",
        Ownership="Paper merchant group",
        **{
            "Continent / Geography": "EU(Europe)",
            "Country Code": "SE",
            "Region Code": "EU",
            "Core Categories": "Paper; packaging materials",
            "Specialty Focus": "European paper and packaging merchant",
            "Key Brands Represented": "Papyrus",
            "Summary": "Papyrus — European paper/packaging merchant distributor.",
            "_evidence": "papyrus.com paper packaging merchant",
        },
    ),
    mk(
        Company="Inapa",
        Website="https://www.inapa.com",
        Headquarters="Portugal / Europe",
        Ownership="Inapa",
        **{
            "Continent / Geography": "EU(Europe)",
            "Country Code": "PT",
            "Region Code": "EU",
            "Core Categories": "Paper; packaging",
            "Specialty Focus": "European paper and packaging merchant",
            "Key Brands Represented": "Inapa",
            "Summary": "Inapa — European paper/packaging merchant distributor.",
            "_evidence": "inapa.com paper packaging merchant",
        },
    ),
    mk(
        Company="Emballages Service France",
        Website="https://www.emballages-service.fr",
        Headquarters="France",
        **{
            "Continent / Geography": "EU(France)",
            "Country Code": "FR",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies; shipping materials",
            "Specialty Focus": "French packaging supplies merchant",
            "Key Brands Represented": "Emballages Service",
            "Summary": "Emballages Service France — packaging supplies merchant/distributor.",
            "_evidence": "emballages-service.fr packaging merchant",
        },
    ),
    mk(
        Company="JPG Belgium",
        Website="https://www.jpg.be",
        Headquarters="Belgium",
        Ownership="JPG / Manutan family catalog",
        **{
            "Continent / Geography": "EU(Belgium)",
            "Country Code": "BE",
            "Region Code": "EU",
            "Core Categories": "Industrial supplies; packaging; workplace consumables",
            "Specialty Focus": "Belgian B2B catalog distributor",
            "Key Brands Represented": "JPG",
            "Summary": "JPG Belgium — B2B catalog distributor including packaging supplies.",
            "_evidence": "jpg.be catalog distributor (JPG/Manutan family)",
        },
    ),
    mk(
        Company="Premier Packaging NZ",
        Website="https://www.premierpackaging.co.nz",
        Headquarters="New Zealand",
        **{
            "Continent / Geography": "APAC(New Zealand)",
            "Country Code": "NZ",
            "Region Code": "APAC",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "New Zealand packaging supplies distributor",
            "Key Brands Represented": "Premier Packaging",
            "Summary": "Premier Packaging NZ — packaging supplies distributor.",
            "_evidence": "premierpackaging.co.nz",
        },
    ),
    mk(
        Company="Bidvest Paperplus",
        Website="https://www.bidvestpaperplus.co.za",
        Headquarters="South Africa",
        Ownership="Bidvest",
        **{
            "Continent / Geography": "MEA(South Africa)",
            "Country Code": "ZA",
            "Region Code": "MEA",
            "Core Categories": "Paper; packaging; office and packaging supplies",
            "Specialty Focus": "South African paper and packaging distributor (Bidvest)",
            "Key Brands Represented": "Bidvest Paperplus",
            "Summary": "Bidvest Paperplus — South African paper/packaging distributor.",
            "_evidence": "bidvestpaperplus.co.za Bidvest packaging/paper distribution",
        },
    ),
    mk(
        Company="Packaging Distributors South Africa",
        Website="https://www.packagingdistributors.co.za",
        Headquarters="South Africa",
        **{
            "Continent / Geography": "MEA(South Africa)",
            "Country Code": "ZA",
            "Region Code": "MEA",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "South African packaging distributor",
            "Key Brands Represented": "Packaging Distributors",
            "Summary": "Packaging Distributors South Africa — packaging distributor.",
            "_evidence": "packagingdistributors.co.za",
        },
    ),
    mk(
        Company="Seow Khim Polythelene (SKP)",
        Website="https://www.skp.com.sg",
        Headquarters="Singapore",
        **{
            "Continent / Geography": "APAC(Singapore)",
            "Country Code": "SG",
            "Region Code": "APAC",
            "Core Categories": "Polyethylene packaging; plastic packaging supplies",
            "Specialty Focus": "Singapore packaging materials supplier/distributor",
            "Key Brands Represented": "SKP",
            "Summary": "Seow Khim Polythelene (SKP) Singapore — packaging materials supplier/distributor.",
            "_evidence": "skp.com.sg Singapore packaging supply store/distributor",
        },
    ),
    mk(
        Company="Tantraco Enterprise",
        Website="https://www.tantraco.com.sg",
        Headquarters="Singapore",
        **{
            "Continent / Geography": "APAC(Singapore)",
            "Country Code": "SG",
            "Region Code": "APAC",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "Singapore packaging supplies distributor",
            "Key Brands Represented": "Tantraco",
            "Summary": "Tantraco Enterprise Singapore — packaging supplies distributor.",
            "_evidence": "tantraco.com.sg",
        },
    ),
    mk(
        Company="Bizongo",
        Website="https://www.bizongo.com",
        Headquarters="India",
        **{
            "Continent / Geography": "APAC(India)",
            "Country Code": "IN",
            "Region Code": "APAC",
            "Core Categories": "Packaging procurement; packaging supplies platform",
            "Specialty Focus": "India B2B packaging sourcing/distribution platform",
            "Key Brands Represented": "Bizongo",
            "Summary": "Bizongo — India B2B packaging platform/distributor connecting buyers to packaging supply.",
            "_evidence": "bizongo.com packaging B2B platform",
        },
    ),
    mk(
        Company="UPack India",
        Website="https://www.upack.in",
        Headquarters="India",
        **{
            "Continent / Geography": "APAC(India)",
            "Country Code": "IN",
            "Region Code": "APAC",
            "Core Categories": "Packaging supplies; shipping materials",
            "Specialty Focus": "India packaging supplies marketplace/distributor",
            "Key Brands Represented": "UPack",
            "Summary": "UPack India — packaging supplies distributor/marketplace.",
            "_evidence": "upack.in packaging supplies",
        },
    ),
    mk(
        Company="India Pack",
        Website="https://www.indiapack.co.in",
        Founded="2016",
        Headquarters="India",
        **{
            "Continent / Geography": "APAC(India)",
            "Country Code": "IN",
            "Region Code": "APAC",
            "Core Categories": "Packaging procurement; packaging supplies",
            "Specialty Focus": "India packaging sourcing/channel partner",
            "Key Brands Represented": "India Pack",
            "Summary": "India Pack — packaging procurement/sourcing channel partner (founded 2016 per prior verified pass).",
            "_evidence": "indiapack.co.in; prior DeepSeek pass accepted founded 2016",
        },
    ),
    mk(
        Company="Shippers Supply",
        Website="https://www.shipperssupply.com",
        Founded="1975",
        Headquarters="Edmonton, Alberta, Canada (historic; now under Spicers brand network)",
        Ownership="Spicers Canada / CNG (acquired)",
        **{
            "Continent / Geography": "NA(Canada)",
            "Operational Presence": "Western Canada packaging distribution heritage",
            "Country Code": "CA",
            "Region Code": "NA",
            "Core Categories": "Packaging; shipping; warehouse equipment",
            "Specialty Focus": "Canadian packaging distributor (Shippers Supply heritage brand)",
            "Key Brands Represented": "Shippers Supply / Spicers",
            "Summary": "Shippers Supply — Western Canada packaging distributor founded 1975; later acquired into Spicers Canada.",
            "_evidence": "Growjo/PR — founded 1975 Edmonton; Spicers acquisition 2017",
        },
    ),
    mk(
        Company="Unisource Canada",
        Website="https://www.unisource.ca",
        Headquarters="Canada",
        **{
            "Continent / Geography": "NA(Canada)",
            "Country Code": "CA",
            "Region Code": "NA",
            "Core Categories": "Paper; packaging; facility supplies",
            "Specialty Focus": "Canadian paper/packaging distributor brand",
            "Key Brands Represented": "Unisource",
            "Summary": "Unisource Canada — paper and packaging facility supplies distributor brand.",
            "_evidence": "unisource.ca",
        },
    ),
    mk(
        Company="Ideal Packaging Canada",
        Website="https://www.idealpackaging.ca",
        Headquarters="Canada",
        **{
            "Continent / Geography": "NA(Canada)",
            "Country Code": "CA",
            "Region Code": "NA",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "Canadian packaging distributor",
            "Key Brands Represented": "Ideal Packaging",
            "Summary": "Ideal Packaging Canada — packaging distributor.",
            "_evidence": "idealpackaging.ca",
        },
    ),
    mk(
        Company="Landmark Paper & Packaging",
        Website="https://www.landmarkpaper.com",
        Headquarters="Canada",
        **{
            "Continent / Geography": "NA(Canada)",
            "Country Code": "CA",
            "Region Code": "NA",
            "Core Categories": "Paper; packaging",
            "Specialty Focus": "Canadian paper and packaging distributor",
            "Key Brands Represented": "Landmark",
            "Summary": "Landmark Paper & Packaging — Canadian paper/packaging distributor.",
            "_evidence": "landmarkpaper.com",
        },
    ),
    mk(
        Company="Atlantic Packaging Products",
        Website="https://www.atlantic.ca",
        Headquarters="Canada",
        **{
            "Continent / Geography": "NA(Canada)",
            "Country Code": "CA",
            "Region Code": "NA",
            "Core Categories": "Packaging products; corrugated and packaging supplies",
            "Specialty Focus": "Canadian packaging products company/distributor network",
            "Key Brands Represented": "Atlantic Packaging",
            "Summary": "Atlantic Packaging Products Canada — packaging products supplier/distributor.",
            "_evidence": "atlantic.ca",
        },
    ),
    mk(
        Company="Pride Container",
        Website="https://www.pridecontainer.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Containers; packaging containers",
            "Specialty Focus": "US packaging containers distributor",
            "Key Brands Represented": "Pride Container",
            "Summary": "Pride Container — US packaging containers distributor.",
            "_evidence": "pridecontainer.com",
        },
    ),
    mk(
        Company="Hisco",
        Website="https://www.hisco.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Industrial supplies; packaging-related MRO; adhesives; tapes",
            "Specialty Focus": "US industrial distributor including packaging materials",
            "Key Brands Represented": "Hisco / multi-brand",
            "Summary": "Hisco — US industrial distributor including packaging-related supplies.",
            "_evidence": "hisco.com industrial distributor",
        },
    ),
    mk(
        Company="The Bottle Crew",
        Website="https://www.thebottlecrew.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Bottles; containers; packaging",
            "Specialty Focus": "US bottles/containers packaging distributor",
            "Key Brands Represented": "The Bottle Crew",
            "Summary": "The Bottle Crew — US bottles and packaging containers distributor.",
            "_evidence": "thebottlecrew.com",
        },
    ),
    mk(
        Company="Box Brothers",
        Website="https://www.boxbrothers.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Boxes; shipping packaging; packaging supplies",
            "Specialty Focus": "US packaging/shipping supplies retailer-distributor",
            "Key Brands Represented": "Box Brothers",
            "Summary": "Box Brothers — US packaging and shipping supplies distributor/retailer.",
            "_evidence": "boxbrothers.com",
        },
    ),
    mk(
        Company="Uboxes",
        Website="https://www.uboxes.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Moving boxes; shipping packaging",
            "Specialty Focus": "US boxes/packaging supplies merchant",
            "Key Brands Represented": "Uboxes",
            "Summary": "Uboxes — US boxes and packaging supplies merchant.",
            "_evidence": "uboxes.com",
        },
    ),
    mk(
        Company="MrBoxOnline",
        Website="https://www.mrboxonline.com",
        Headquarters="Sunrise, Florida, USA",
        Ownership="SupplyOne (acquired A&M Tape & Packaging / MrBoxOnline)",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Shipping boxes; packaging materials; warehouse supplies",
            "Specialty Focus": "Online packaging/shipping supplies distributor",
            "Key Brands Represented": "MrBoxOnline",
            "Summary": "MrBoxOnline — online US packaging and shipping supplies distributor (SupplyOne family).",
            "_evidence": "mrboxonline.com packaging shipping supplies distributor",
        },
    ),
    mk(
        Company="Package Design & Supply",
        Website="https://packagedesignsupply.com",
        Headquarters="Buffalo, New York, USA",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Packaging supplies; custom packaging programs",
            "Specialty Focus": "Buffalo packaging distributor/fabricator channel",
            "Key Brands Represented": "Package Design & Supply",
            "Summary": "Package Design & Supply — Buffalo NY packaging distributor/channel supplier.",
            "_evidence": "packagedesignsupply.com",
        },
    ),
    mk(
        Company="Acme Paper & Supply",
        Website="https://www.acmepaper.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Paper; packaging supplies",
            "Specialty Focus": "US paper and packaging distributor",
            "Key Brands Represented": "Acme Paper",
            "Summary": "Acme Paper & Supply — US paper/packaging distributor.",
            "_evidence": "acmepaper.com",
        },
    ),
    mk(
        Company="Carolina Retail Packaging",
        Website="https://www.carolinaretail.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Retail packaging; packaging supplies",
            "Specialty Focus": "US retail packaging distributor",
            "Key Brands Represented": "Carolina Retail Packaging",
            "Summary": "Carolina Retail Packaging — US retail packaging distributor.",
            "_evidence": "carolinaretail.com",
        },
    ),
    mk(
        Company="WB Mason",
        Website="https://www.wbmason.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Office supplies; packaging; shipping materials",
            "Specialty Focus": "US office/packaging supplies distributor",
            "Key Brands Represented": "WB Mason",
            "Summary": "WB Mason — US office and packaging/shipping supplies distributor.",
            "_evidence": "wbmason.com",
        },
    ),
    mk(
        Company="Reliable Corporation",
        Website="https://www.reliable.com",
        Headquarters="United States",
        **{
            "Continent / Geography": "NA(United States)",
            "Country Code": "US",
            "Region Code": "NA",
            "Core Categories": "Office supplies; packaging",
            "Specialty Focus": "US office/packaging supplies catalog distributor",
            "Key Brands Represented": "Reliable",
            "Summary": "Reliable — US office and packaging supplies catalog distributor.",
            "_evidence": "reliable.com",
        },
    ),
    mk(
        Company="All Size Packaging Canada",
        Website="https://www.allsizepackaging.com",
        Headquarters="Canada",
        **{
            "Continent / Geography": "NA(Canada)",
            "Country Code": "CA",
            "Region Code": "NA",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "Canadian packaging distributor",
            "Key Brands Represented": "All Size Packaging",
            "Summary": "All Size Packaging — Canadian packaging distributor.",
            "_evidence": "allsizepackaging.com",
        },
    ),
    mk(
        Company="Empaques del Norte",
        Website="https://www.empaquesdelnorte.com",
        Headquarters="Mexico",
        **{
            "Continent / Geography": "LATAM(Mexico)",
            "Country Code": "MX",
            "Region Code": "LATAM",
            "Core Categories": "Industrial packaging; empaques",
            "Specialty Focus": "Mexico industrial packaging distributor",
            "Key Brands Represented": "Empaques del Norte",
            "Summary": "Empaques del Norte — Mexico industrial packaging distributor.",
            "_evidence": "empaquesdelnorte.com",
        },
    ),
    mk(
        Company="Clearpack",
        Website="https://www.clearpack.com",
        Headquarters="Singapore / Asia",
        **{
            "Continent / Geography": "APAC(Singapore / Asia)",
            "Country Code": "SG",
            "Region Code": "APAC",
            "Core Categories": "Packaging machinery; packaging solutions distribution",
            "Specialty Focus": "Asia packaging solutions distributor",
            "Key Brands Represented": "Clearpack",
            "Summary": "Clearpack — Asia packaging solutions distributor.",
            "_evidence": "clearpack.com",
        },
    ),
    mk(
        Company="Lesters Packaging",
        Website="https://www.lesterspackaging.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "UK packaging supplier/distributor",
            "Key Brands Represented": "Lesters Packaging",
            "Summary": "Lesters Packaging — UK packaging supplier/distributor.",
            "_evidence": "lesterspackaging.co.uk",
        },
    ),
    mk(
        Company="Direct Packaging UK",
        Website="https://www.directpackaging.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "UK packaging supplier/distributor",
            "Key Brands Represented": "Direct Packaging",
            "Summary": "Direct Packaging UK — packaging supplier/distributor.",
            "_evidence": "directpackaging.co.uk",
        },
    ),
    mk(
        Company="Aropack",
        Website="https://www.aropack.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "UK packaging merchant",
            "Key Brands Represented": "Aropack",
            "Summary": "Aropack — UK packaging merchant/supplier.",
            "_evidence": "aropack.co.uk",
        },
    ),
    mk(
        Company="Abco Kovex UK",
        Website="https://www.abcokovex.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies; protective packaging",
            "Specialty Focus": "UK packaging supplies distributor",
            "Key Brands Represented": "Abco Kovex",
            "Summary": "Abco Kovex UK — packaging supplies distributor.",
            "_evidence": "abcokovex.co.uk",
        },
    ),
    mk(
        Company="IMG Packing",
        Website="https://www.imgpacking.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packing boxes; packaging supplies",
            "Specialty Focus": "UK packaging supply store/distributor",
            "Key Brands Represented": "IMG Packing",
            "Summary": "IMG Packing — UK packaging supplies distributor.",
            "_evidence": "imgpacking.co.uk",
        },
    ),
    mk(
        Company="BM Packaging",
        Website="https://www.bmpackaging.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "UK packaging supplier",
            "Key Brands Represented": "BM Packaging",
            "Summary": "BM Packaging — UK packaging supplier/distributor.",
            "_evidence": "bmpackaging.co.uk",
        },
    ),
    mk(
        Company="Agenda Boxes",
        Website="https://www.agendaboxes.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Boxes; packaging",
            "Specialty Focus": "UK boxes/packaging supplier",
            "Key Brands Represented": "Agenda Boxes",
            "Summary": "Agenda Boxes — UK boxes and packaging supplier.",
            "_evidence": "agendaboxes.co.uk",
        },
    ),
    mk(
        Company="Jaymar Packaging",
        Website="https://www.jaymarpackaging.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "UK packaging supplier",
            "Key Brands Represented": "Jaymar",
            "Summary": "Jaymar Packaging — UK packaging supplier/distributor.",
            "_evidence": "jaymarpackaging.co.uk",
        },
    ),
    mk(
        Company="UK Packaging Supplies London",
        Website="https://www.ukpackagingsupplies.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "UK packaging supplies merchant",
            "Key Brands Represented": "UK Packaging Supplies",
            "Summary": "UK Packaging Supplies — packaging supplies merchant.",
            "_evidence": "ukpackagingsupplies.co.uk",
        },
    ),
    mk(
        Company="Online Packaging Supplies UK",
        Website="https://www.onlinepackagingsupplies.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging supplies",
            "Specialty Focus": "UK online packaging supplies merchant",
            "Key Brands Represented": "Online Packaging Supplies",
            "Summary": "Online Packaging Supplies UK — packaging supplies merchant.",
            "_evidence": "onlinepackagingsupplies.co.uk",
        },
    ),
    mk(
        Company="Cheap Packaging Materials",
        Website="https://www.cheappackagingmaterials.co.uk",
        Headquarters="United Kingdom",
        **{
            "Continent / Geography": "EU(United Kingdom)",
            "Country Code": "GB",
            "Region Code": "EU",
            "Core Categories": "Packaging materials",
            "Specialty Focus": "UK packaging materials merchant",
            "Key Brands Represented": "Cheap Packaging Materials",
            "Summary": "Cheap Packaging Materials UK — packaging materials merchant.",
            "_evidence": "cheappackagingmaterials.co.uk",
        },
    ),
    # Ahlsell Denmark already added in batch3 - skip if exists
]


def domain_of(url: str) -> str:
    try:
        return urlparse(url if "://" in url else "https://" + url).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def main():
    backup = FP.with_name(FP.stem + f"_pre_agent_manual5_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(FP, backup)
    wb = load_workbook(FP)
    ws = wb.active

    # Remove Canpaco hybrid manufacturer/converter
    removed = []
    keep = [list(c.value for c in ws[1])]
    for r in ws.iter_rows(min_row=2, values_only=True):
        name = (r[0] or "").strip()
        website = (r[1] or "").strip()
        d = domain_of(website)
        if d == "canpaco.com" or name.lower().startswith("canpaco"):
            removed.append({"company": name, "website": website, "reason": "hybrid converter/manufacturer"})
            continue
        keep.append(list(r))
    ws.delete_rows(1, ws.max_row)
    for row in keep:
        ws.append(row)

    names, domains = set(), set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        names.add((r[0] or "").strip().lower())
        domains.add(domain_of(str(r[1] or "")))

    added, skipped = [], []
    for rowd in ROWS:
        d = domain_of(rowd["Website"])
        if rowd["Company"].lower() in names or d in domains:
            skipped.append(rowd["Company"])
            continue
        ws.append([rowd.get(h) for h in HEADERS])
        names.add(rowd["Company"].lower())
        domains.add(d)
        rec = {k: rowd.get(k) for k in HEADERS}
        rec["_evidence"] = rowd.get("_evidence")
        added.append(rec)

    for col in ws.columns:
        for cell in col:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col[0].column_letter].width = 22
    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(FP)
    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "removed": removed,
        "added_count": len(added),
        "skipped": skipped,
        "final_rows": ws.max_row - 1,
        "added": added,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print("removed", len(removed), "added", len(added), "skipped", len(skipped), "final", audit["final_rows"])


if __name__ == "__main__":
    main()
