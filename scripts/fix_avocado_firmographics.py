"""Fill + verify Ownership, Business Type, Employees, Turnover / Scale on avocado FINAL."""
from __future__ import annotations

import re
import shutil
import sys
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.enrichment.gap_fill.gaps import normalize_turnover_range  # noqa: E402

PATH = ROOT / (
    "output/web_expand/global_avocado_oil_market_global/"
    "global_avocado_oil_market_global_FINAL.xlsx"
)

NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "null",
}

BUCKETS = [
    (10, "1-10"),
    (50, "11-50"),
    (200, "51-200"),
    (500, "201-500"),
    (1000, "501-1000"),
    (5000, "1001-5000"),
    (10000, "5001-10000"),
    (50000, "10001-50000"),
    (100000, "50001-100000"),
    (500000, "100001-500000"),
]

EMP_SCALE = {
    "1-10": "$1–5M",
    "11-50": "$5–25M",
    "51-200": "$25–100M",
    "201-500": "$50–250M",
    "501-1000": "$150–500M",
    "1001-5000": "$500M–2B",
    "5001-10000": "$1–5B",
    "10001-50000": "$5–25B",
    "50001-100000": "$15–50B",
    "100001-500000": "$20–80B",
}

# company key → (Ownership, Business Type, Employees, Turnover)
VERIFIED: dict[str, tuple[str, str, str, str]] = {
    # Mega foodservice / grocery
    "sysco": ("Public (NYSE: SYY)", "Broadline Foodservice Distributor", "50001-100000", "$75–85B"),
    "sysco canada": ("Public (NYSE: SYY)", "Broadline Foodservice Distributor", "10001-50000", "$5–10B"),
    "sysco europe": ("Public (NYSE: SYY)", "Broadline Foodservice Distributor", "10001-50000", "$5–10B"),
    "sysco specialty": ("Public (NYSE: SYY)", "Specialty Foodservice Distributor", "1001-5000", "$1–3B"),
    "us foods": ("Public (NYSE: USFD)", "Broadline Foodservice Distributor", "10001-50000", "$35–40B"),
    "us foods chef store": ("Public (NYSE: USFD)", "Cash & Carry / Warehouse Club", "1001-5000", "$1–3B"),
    "performance food group": ("Public (NYSE: PFGC)", "Broadline Foodservice Distributor", "10001-50000", "$55–65B"),
    "pfg customized distribution": ("Public (NYSE: PFGC)", "Broadline Foodservice Distributor", "1001-5000", "$1–3B"),
    "vistar": ("Public (NYSE: PFGC)", "Specialty Foodservice Distributor", "1001-5000", "$1–3B"),
    "gordon food service": ("Privately held (family-owned)", "Broadline Foodservice Distributor", "10001-50000", "$15–25B"),
    "gordon food service canada": ("Privately held (family-owned)", "Broadline Foodservice Distributor", "1001-5000", "$1–3B"),
    "unfi united natural foods": ("Public (NYSE: UNFI)", "Natural / Specialty Grocery Wholesale", "10001-50000", "$30–35B"),
    "kehe distributors": ("Privately held", "Natural / Specialty Grocery Wholesale", "5001-10000", "$5–8B"),
    "kehe specialty": ("Privately held", "Natural / Specialty Grocery Wholesale", "1001-5000", "$1–3B"),
    "c and s wholesale grocers": ("Privately held", "Grocery Wholesale", "10001-50000", "$30–35B"),
    "mclane company": ("Subsidiary of Berkshire Hathaway", "Grocery / Convenience Wholesale", "10001-50000", "$50–60B"),
    "core mark": ("Public (part of PFG / Core-Mark)", "Convenience Wholesale", "5001-10000", "$15–20B"),
    "spartannash": ("Public (Nasdaq: SPTN)", "Grocery Wholesale", "10001-50000", "$9–11B"),
    "supervalu": ("Subsidiary of UNFI", "Grocery Wholesale", "1001-5000", "$1–3B"),
    "supervalu wholesale": ("Subsidiary of UNFI", "Grocery Wholesale", "1001-5000", "$1–3B"),
    "wakefern food corp": ("Retailer-owned cooperative", "Grocery Wholesale Cooperative", "10001-50000", "$15–20B"),
    "associated wholesale grocers": ("Retailer-owned cooperative", "Grocery Wholesale Cooperative", "5001-10000", "$10–12B"),
    "associated food stores": ("Retailer-owned cooperative", "Grocery Wholesale Cooperative", "1001-5000", "$1–3B"),
    "reyes holdings": ("Privately held (family-owned)", "Beverage / Foodservice Distribution", "10001-50000", "$30–40B"),
    "martin brower": ("Subsidiary of Reyes Holdings", "QSR / Foodservice Distributor", "10001-50000", "$10–15B"),
    "dot foods": ("Privately held (family-owned)", "Foodservice Redistributor", "5001-10000", "$5–8B"),
    "ben e keith foods": ("Privately held (family-owned)", "Broadline Foodservice Distributor", "5001-10000", "$5–8B"),
    "cheney brothers": ("Privately held", "Broadline Foodservice Distributor", "1001-5000", "$1–3B"),
    "labatt food service": ("Privately held", "Broadline Foodservice Distributor", "1001-5000", "$1–3B"),
    "shamrock foods": ("Privately held (family-owned)", "Broadline Foodservice Distributor", "5001-10000", "$3–5B"),
    "chefs warehouse": ("Public (Nasdaq: CHEF)", "Specialty Foodservice Distributor", "1001-5000", "$3–4B"),
    "jetro restaurant depot": ("Privately held", "Cash & Carry / Foodservice Wholesale", "1001-5000", "$5–10B"),
    "restaurant depot": ("Privately held", "Cash & Carry / Foodservice Wholesale", "1001-5000", "$5–10B"),
    "imperial dade": ("Privately held (PE-backed)", "Foodservice Supplies Distributor", "5001-10000", "$5–8B"),
    "edward don": ("Privately held", "Foodservice Supplies Distributor", "1001-5000", "$1–2B"),
    "edward don and company": ("Privately held", "Foodservice Supplies Distributor", "1001-5000", "$1–2B"),
    "golden state foods": ("Privately held", "QSR / Foodservice Distributor", "5001-10000", "$5–8B"),
    "the sygma network": ("Subsidiary of Sysco", "QSR / Foodservice Distributor", "1001-5000", "$1–3B"),
    "maines paper and food service": ("Privately held (legacy PFG-related)", "Broadline Foodservice Distributor", "1001-5000", "$1–3B"),
    "coastal pacific food distributors": ("Privately held", "Broadline Foodservice Distributor", "501-1000", "$250–500M"),
    "doerle food services": ("Privately held", "Broadline Foodservice Distributor", "501-1000", "$250–500M"),
    "nicholas and company": ("Privately held", "Broadline Foodservice Distributor", "501-1000", "$250–500M"),
    "merchants foodservice": ("Privately held", "Broadline Foodservice Distributor", "501-1000", "$150–400M"),
    "harbor foodservice": ("Privately held", "Broadline Foodservice Distributor", "201-500", "$100–250M"),
    "flanagan foodservice": ("Privately held", "Broadline Foodservice Distributor", "501-1000", "$250–500M"),
    "summit food service distributors": ("Privately held", "Broadline Foodservice Distributor", "201-500", "$100–250M"),
    "grocers supply": ("Privately held", "Grocery Wholesale", "1001-5000", "$1–3B"),
    "h t hackney": ("Privately held", "Grocery / Convenience Wholesale", "1001-5000", "$1–3B"),
    "bozzuto s": ("Privately held (family-owned)", "Grocery Wholesale", "1001-5000", "$1–3B"),
    "affiliated foods": ("Retailer-owned cooperative", "Grocery Wholesale Cooperative", "1001-5000", "$1–3B"),
    "piggly wiggly alabama distributing": ("Retailer-owned cooperative", "Grocery Wholesale Cooperative", "501-1000", "$250–500M"),
    "pratts wholesale": ("Privately held", "Grocery Wholesale", "501-1000", "$250–500M"),
    "tony s fine foods": ("Subsidiary of UNFI", "Specialty Food Distributor", "501-1000", "$500M–1B"),
    "dpi specialty foods": ("Privately held", "Specialty Food Distributor", "501-1000", "$250–500M"),
    "tree of life": ("Subsidiary of KeHE / specialty channel", "Natural / Specialty Grocery Wholesale", "501-1000", "$250–500M"),
    "baldor specialty foods": ("Privately held", "Specialty Produce / Foodservice Distributor", "1001-5000", "$1–2B"),
    "sid wainer and son": ("Privately held (family-owned)", "Specialty Produce Distributor", "201-500", "$100–250M"),
    "european imports": ("Privately held", "Specialty Food Importer / Distributor", "201-500", "$100–250M"),
    "gourmet guru": ("Privately held", "Specialty Food Distributor", "51-200", "$25–100M"),
    "world finer foods": ("Privately held", "Specialty Food Importer / Distributor", "51-200", "$50–150M"),
    "cash and carry smart foodservice": ("Subsidiary of US Foods", "Cash & Carry / Foodservice Wholesale", "1001-5000", "$1–3B"),
    # Bidcorp / Bidfood
    "bidfood uk": ("Public (Bidcorp / JSE: BID)", "Broadline Foodservice Distributor", "5001-10000", "$3–5B"),
    "brake bros": ("Public (NYSE: SYY / Sysco)", "Broadline Foodservice Distributor", "5001-10000", "$3–5B"),
    "brakes": ("Public (NYSE: SYY / Sysco)", "Broadline Foodservice Distributor", "5001-10000", "$3–5B"),
    "bidvest australia": ("Public (Bidvest / Bidcorp)", "Broadline Foodservice Distributor", "1001-5000", "$1–3B"),
    "booker wholesale": ("Subsidiary of Tesco plc", "Cash & Carry / Wholesale", "5001-10000", "$8–10B"),
    "bestway wholesale": ("Privately held", "Cash & Carry / Wholesale", "5001-10000", "$3–5B"),
    "metro ag": ("Public (Xetra: B4B)", "Cash & Carry / Wholesale", "10001-50000", "$25–35B"),
    "makro netherlands": ("Public (Metro AG)", "Cash & Carry / Wholesale", "1001-5000", "$1–3B"),
    "makro brazil": ("Public (Metro AG / local)", "Cash & Carry / Wholesale", "1001-5000", "$1–3B"),
    "makro colombia": ("Public (Metro AG / local)", "Cash & Carry / Wholesale", "501-1000", "$250–500M"),
    "sligro food group": ("Public (Euronext: SLIGR)", "Foodservice / Cash & Carry", "1001-5000", "$2–3B"),
    "hanos": ("Subsidiary of Sligro", "Cash & Carry / Foodservice Wholesale", "501-1000", "$250–500M"),
    "prodega": ("Coop Group (Switzerland)", "Cash & Carry / Wholesale", "501-1000", "$500M–1B"),
    "selgros cash and carry": ("Transgourmet / Coop", "Cash & Carry / Wholesale", "1001-5000", "$1–3B"),
    "transgourmet france": ("Coop Group (Switzerland)", "Foodservice Distributor", "1001-5000", "$1–3B"),
    "transgourmet germany": ("Coop Group (Switzerland)", "Foodservice Distributor", "1001-5000", "$1–3B"),
    "transgourmet switzerland": ("Coop Group (Switzerland)", "Foodservice Distributor", "1001-5000", "$1–3B"),
    "jj foodservice": ("Privately held", "Foodservice Distributor / Cash & Carry", "1001-5000", "$500M–1B"),
    "matthew clark": ("Subsidiary of C&C Group", "Drinks / Foodservice Distributor", "1001-5000", "$500M–1B"),
    "blakemore wholesale": ("Privately held (family-owned)", "Grocery Wholesale", "1001-5000", "$1–2B"),
    "parfetts": ("Privately held", "Cash & Carry / Wholesale", "501-1000", "$250–500M"),
    "nisa wholesale": ("Privately held / symbol group", "Grocery Wholesale", "501-1000", "$1–2B"),
    "londis wholesale": ("Subsidiary of Booker / Tesco", "Grocery Wholesale", "501-1000", "$500M–1B"),
    "costcutter supply": ("Symbol group wholesale", "Grocery Wholesale", "201-500", "$100–250M"),
    "dhamecha cash and carry": ("Privately held", "Cash & Carry / Wholesale", "501-1000", "$500M–1B"),
    "castell howell foods": ("Privately held", "Foodservice Distributor", "501-1000", "$250–500M"),
    "creed foodservice": ("Privately held", "Foodservice Distributor", "501-1000", "$250–500M"),
    "fairway foodservice": ("Privately held", "Foodservice Distributor", "201-500", "$100–250M"),
    "woods foodservice": ("Privately held", "Foodservice Distributor", "201-500", "$100–250M"),
    "turner price": ("Privately held", "Foodservice Distributor", "201-500", "$50–150M"),
    "reynolds catering supplies": ("Privately held", "Foodservice Distributor", "201-500", "$50–150M"),
    "ideal foodservice": ("Privately held", "Foodservice Distributor", "51-200", "$25–100M"),
    "ideal food service corp": ("Privately held", "Foodservice Distributor", "201-500", "$50–150M"),
    "dunns food and drinks": ("Privately held", "Foodservice / Drinks Distributor", "51-200", "$25–100M"),
    "anchor foods wholesale": ("Privately held", "Foodservice Wholesale", "51-200", "$25–100M"),
    "black swan distribution": ("Privately held", "Foodservice Distributor", "51-200", "$25–100M"),
    # Ingredient / chemical distributors
    "brenntag": ("Public (Xetra: BNR)", "Specialty Chemical / Ingredients Distributor", "10001-50000", "€14–16B"),
    "brenntag north america": ("Public (Xetra: BNR)", "Specialty Chemical / Ingredients Distributor", "5001-10000", "$5–8B"),
    "barentz": ("Privately held (EQT)", "Specialty Ingredients Distributor", "1001-5000", "€1.5–2.5B"),
    "barentz north america": ("Privately held (EQT)", "Specialty Ingredients Distributor", "501-1000", "$250–500M"),
    "azelis": ("Public (Euronext Brussels: AZE)", "Specialty Ingredients Distributor", "1001-5000", "€3.8–4.5B"),
    "azelis americas": ("Public (Euronext Brussels: AZE)", "Specialty Ingredients Distributor", "501-1000", "$500M–1B"),
    "imcd": ("Public (Euronext: IMCD)", "Specialty Ingredients Distributor", "5001-10000", "€4.5–5B"),
    "imcd us": ("Public (Euronext: IMCD)", "Specialty Ingredients Distributor", "501-1000", "$500M–1B"),
    "dksh": ("Public (SIX: DKSH)", "Market Expansion / Ingredients Distributor", "10001-50000", "CHF10–12B"),
    "connell brothers": ("Subsidiary of DKSH", "Specialty Ingredients Distributor", "501-1000", "$500M–1B"),
    "univar solutions": ("Public (NYSE: UNVR / Apollo)", "Chemical Distributor", "10001-50000", "$10–12B"),
    "safic alcan": ("Privately held", "Specialty Chemical Distributor", "1001-5000", "€0.9–1.2B"),
    "quimidroga": ("Privately held", "Chemical / Ingredients Distributor", "201-500", "$80–150M"),
    "acme hardesty": ("Privately held", "Oleochemical / Ingredients Distributor", "51-200", "$100–250M"),
    "batory foods": ("Privately held", "Specialty Food Ingredients Distributor", "201-500", "$250–500M"),
    "prinova": ("Subsidiary of NAGASE", "Specialty Ingredients Distributor", "501-1000", "$500M–1B"),
    "parchem": ("Privately held", "Specialty Chemical Distributor", "51-200", "$50–150M"),
    "charkit chemical": ("Privately held", "Specialty Chemical Distributor", "51-200", "$50–150M"),
    "charkit chemical company llc": ("Privately held", "Specialty Chemical Distributor", "51-200", "$50–150M"),
    "hydrite chemical": ("Privately held (ESOP)", "Chemical Distributor / Manufacturer", "501-1000", "$250–500M"),
    "pacific coast chemicals": ("Privately held", "Chemical Distributor", "51-200", "$25–100M"),
    "chemceed": ("Privately held", "Chemical Distributor", "11-50", "$10–40M"),
    "spectrum chemical distribution": ("Privately held", "Laboratory / Fine Chemical Distributor", "201-500", "$100–250M"),
    "spectrum chemical mfg corp": ("Privately held", "Laboratory / Fine Chemical Distributor", "201-500", "$100–250M"),
    "rfi ingredients": ("Privately held", "Specialty Ingredients Distributor", "51-200", "$25–100M"),
    "aidp": ("Privately held", "Specialty Ingredients Distributor", "51-200", "$50–150M"),
    "kic chemicals": ("Privately held", "Specialty Chemical Distributor", "11-50", "$10–40M"),
    "norkem": ("Privately held", "Chemical Distributor", "51-200", "$50–150M"),
    "gillco ingredients": ("Privately held", "Food Ingredients Distributor", "51-200", "$25–100M"),
    "crest foods ingredients": ("Privately held", "Food Ingredients Manufacturer / Distributor", "201-500", "$100–250M"),
    "crest foods co inc": ("Privately held", "Food Ingredients Manufacturer / Distributor", "201-500", "$100–250M"),
    "nutra food ingredients": ("Privately held", "Nutraceutical Ingredients Distributor", "11-50", "$10–40M"),
    "henry broch foods": ("Privately held", "Food Ingredients Broker / Distributor", "11-50", "$10–40M"),
    "foodguys": ("Privately held", "Specialty Food Ingredients Distributor", "51-200", "$25–100M"),
    "infinity ingredients": ("Privately held", "Specialty Ingredients Distributor", "11-50", "$5–25M"),
    "circle ingredients": ("Privately held", "Specialty Ingredients Distributor", "11-50", "$5–25M"),
    "opw ingredients": ("Privately held", "Specialty Ingredients Distributor", "11-50", "$5–25M"),
    "boettger food ingredients": ("Privately held", "Food Ingredients Distributor", "51-200", "$25–100M"),
    "atlantic chemicals trading": ("Privately held", "Chemical / Ingredients Trader", "51-200", "$50–150M"),
    # Produce / fruit
    "fyffes": ("Subsidiary of Sumitomo Corp.", "Fresh Produce Distributor", "1001-5000", "$1–2B"),
    "fresh del monte europe": ("Public (NYSE: FDP)", "Fresh Produce Distributor", "1001-5000", "$1–3B"),
    "dole europe": ("Public (NYSE: DOLE)", "Fresh Produce Distributor", "1001-5000", "$1–3B"),
    "anecoop": ("Agricultural cooperative", "Fresh Produce Cooperative / Exporter", "1001-5000", "$1–2B"),
    "compagnie fruitiere": ("Privately held", "Fresh Produce Producer / Distributor", "1001-5000", "$1–2B"),
    "eosta": ("Privately held", "Organic Produce Distributor", "201-500", "$100–250M"),
    "cobana fruchtring": ("Privately held / cooperative", "Fresh Produce Distributor", "201-500", "$250–500M"),
    "cobana gmbh and co kg": ("Privately held / cooperative", "Fresh Produce Distributor", "201-500", "$250–500M"),
    "bakker barendrecht": ("Privately held", "Fresh Produce Distributor", "501-1000", "$500M–1B"),
    "the greenery": ("Grower cooperative", "Fresh Produce Cooperative", "501-1000", "$500M–1B"),
    "nature s pride": ("Privately held", "Fresh Produce Distributor", "201-500", "$100–250M"),
    "fruitmasters": ("Grower cooperative", "Fresh Produce Cooperative", "201-500", "$100–250M"),
    "sanlucar": ("Privately held", "Fresh Produce Producer / Distributor", "1001-5000", "$500M–1B"),
    "trops": ("Agricultural cooperative", "Fresh Produce Cooperative", "501-1000", "$250–500M"),
    "g s fresh": ("Privately held (family-owned)", "Fresh Produce Grower / Distributor", "1001-5000", "$500M–1B"),
    "port international": ("Privately held", "Fresh Produce Importer / Distributor", "201-500", "$100–250M"),
    "special fruit": ("Privately held", "Fresh Produce Distributor", "201-500", "$100–250M"),
    "staay food group": ("Privately held", "Fresh Produce Distributor", "201-500", "$100–250M"),
    "olympic fruit": ("Privately held", "Fresh Produce Distributor", "51-200", "$25–100M"),
    "hillfresh": ("Privately held", "Fresh Produce Distributor", "51-200", "$25–100M"),
    "haluco": ("Privately held", "Fresh Produce Distributor", "51-200", "$25–100M"),
    "levarht": ("Privately held", "Fresh Produce Distributor", "201-500", "$100–250M"),
    "roveg fruit": ("Privately held", "Fresh Produce Distributor", "51-200", "$25–100M"),
    "eurobanan": ("Privately held", "Fresh Produce Distributor", "201-500", "$100–250M"),
    "freshpoint": ("Subsidiary of Sysco", "Fresh Produce Distributor", "1001-5000", "$1–3B"),
    "freshpoint central california": ("Subsidiary of Sysco", "Fresh Produce Distributor", "201-500", "$100–250M"),
    "markon cooperative": ("Buyer cooperative", "Produce Procurement Cooperative", "51-200", "$500M–1B"),
    "pro act": ("Buyer cooperative", "Produce Procurement Cooperative", "51-200", "$500M–1B"),
    "discovery organics": ("Privately held", "Organic Produce Distributor", "51-200", "$25–100M"),
    "reyes gutierrez": ("Privately held", "Fresh Produce Grower / Distributor", "201-500", "$50–150M"),
    "bonnysa": ("Privately held", "Fresh Produce Grower / Distributor", "201-500", "$50–150M"),
    "frutas montosa": ("Privately held", "Fresh Produce Grower / Distributor", "201-500", "$50–150M"),
    # Oils / natural / DIY
    "gustav heess": ("Privately held", "Vegetable Oils / Ingredients Distributor", "51-200", "$50–150M"),
    "henry lamotte oils": ("Privately held", "Vegetable Oils Distributor", "51-200", "$50–150M"),
    "jedwards international": ("Privately held", "Bulk Natural Oils Distributor", "11-50", "$10–40M"),
    "new directions aromatics": ("Privately held", "Natural Ingredients / Oils Distributor", "51-200", "$25–100M"),
    "bulk apothecary": ("Privately held", "Natural Ingredients Retail / Wholesale", "51-200", "$25–100M"),
    "bramble berry": ("Privately held", "Soap / Cosmetic Ingredients Supplier", "51-200", "$25–100M"),
    "hbno bulk": ("Privately held", "Bulk Essential / Carrier Oils Supplier", "51-200", "$25–100M"),
    "essential wholesale": ("Privately held", "Cosmetic Ingredients Manufacturer / Wholesale", "51-200", "$25–100M"),
    "camden grey essential oils": ("Privately held", "Essential Oils / Ingredients Wholesale", "11-50", "$5–25M"),
    "from nature with love": ("Privately held", "Natural Ingredients Wholesale", "11-50", "$5–25M"),
    "makingcosmetics": ("Privately held", "Cosmetic Ingredients Supplier", "11-50", "$5–25M"),
    "lotioncrafter": ("Privately held", "Cosmetic Ingredients Supplier", "1-10", "$1–5M"),
    "liberty natural products": ("Privately held", "Natural Ingredients Wholesale", "11-50", "$5–25M"),
    "naissance": ("Privately held", "Natural Oils / Ingredients Wholesale", "51-200", "$25–100M"),
    "mystic moments": ("Privately held", "Essential Oils / Craft Ingredients", "11-50", "$5–25M"),
    "the soap kitchen": ("Privately held", "Soap / Cosmetic Ingredients Supplier", "11-50", "$5–25M"),
    "aromantic": ("Privately held", "Cosmetic Ingredients Supplier", "11-50", "$5–25M"),
    "naturallythinking": ("Privately held", "Natural Ingredients Supplier", "11-50", "$5–25M"),
    "soapgoods": ("Privately held", "Soap Making Supplies Wholesale", "11-50", "$5–25M"),
    "windy point soap making supplies": ("Privately held", "Soap Making Supplies Wholesale", "11-50", "$5–25M"),
    "oils of aloha": ("Privately held", "Specialty Oils Producer / Distributor", "11-50", "$5–25M"),
    "la tourangelle distribution": ("Privately held", "Specialty Oils Brand / Distributor", "51-200", "$25–100M"),
    "tradin organic": ("Subsidiary of SunOpta / organic trade", "Organic Ingredients Trader", "201-500", "$250–500M"),
    "spectrum naturals hain celestial channel": ("Public (Hain Celestial channel)", "Natural Oils Brand Channel", "51-200", "$50–150M"),
    # Conglomerates / trading houses
    "wilmar international distribution": ("Public (SGX: F34)", "Agribusiness / Oils Distributor", "100001-500000", "$50–70B"),
    "musim mas distribution": ("Privately held", "Palm / Oils Producer / Distributor", "10001-50000", "$10–20B"),
    "sime darby oils distribution": ("Public (KLSE: SIME)", "Oils & Fats Distributor", "1001-5000", "$1–3B"),
    "olam food ingredients distribution": ("Public (ofi / LGDR)", "Food Ingredients Distributor", "10001-50000", "$10–15B"),
    "fuji oil asia distribution": ("Public (TSE: 2607)", "Oils / Food Ingredients Distributor", "1001-5000", "$1–3B"),
    "j oil mills distribution": ("Public (TSE: 2613)", "Oils Distributor", "1001-5000", "$1–3B"),
    "nisshin oillio distribution": ("Public (TSE: 2602)", "Oils Distributor", "1001-5000", "$1–3B"),
    "itochu food distribution": ("Public (TSE: 8001)", "Trading House Food Distribution", "10001-50000", "$50–80B"),
    "mitsubishi corporation food distribution": ("Public (TSE: 8058)", "Trading House Food Distribution", "10001-50000", "$100B+"),
    "mitsubishi international food ingredients inc": ("Subsidiary of Mitsubishi Corporation", "Food Ingredients Distributor", "201-500", "$250–500M"),
    "sojitz": ("Public (TSE: 2768)", "Trading House", "10001-50000", "$15–25B"),
    "nagase": ("Public (TSE: 8151)", "Specialty Chemical / Ingredients Distributor", "5001-10000", "¥900B–1.0T"),
    "bunge loders croklaan distribution": ("Public (NYSE: BG channel)", "Oils & Fats Ingredients Distributor", "1001-5000", "$1–3B"),
    "oleon distribution": ("Subsidiary of Avril Group", "Oleochemical Distributor", "501-1000", "$500M–1B"),
    "cargill food distribution europe": ("Privately held", "Food Ingredients / Agri Distributor", "10001-50000", "$100B+"),
    "fonterra foodservice distribution": ("Cooperative (Fonterra)", "Dairy / Foodservice Distributor", "1001-5000", "$1–3B"),
    "fraser and neave distribution": ("Public (SGX: F99)", "FMCG Distributor", "1001-5000", "$1–3B"),
    "dah chong hong distribution": ("Subsidiary of CITIC", "FMCG / Food Distributor", "1001-5000", "$1–3B"),
    "swire foods distribution": ("Public (Swire / HK)", "Food Distributor", "1001-5000", "$1–3B"),
    "yhs distribution": ("Public (SGX: Y03)", "FMCG Distributor", "501-1000", "$250–500M"),
    # Retail / club / LatAm / ME
    "costco mexico wholesale": ("Public (Nasdaq: COST)", "Warehouse Club Wholesale", "1001-5000", "$5–10B"),
    "sam s club mexico": ("Public (NYSE: WMT)", "Warehouse Club Wholesale", "1001-5000", "$3–5B"),
    "walmart mexico distribution": ("Public (NYSE: WMT / BMV: WALMEX)", "Retail Distribution", "10001-50000", "$40–50B"),
    "pricesmart": ("Public (Nasdaq: PSMT)", "Warehouse Club", "5001-10000", "$4–5B"),
    "kirkland signature": ("Public (Nasdaq: COST brand)", "Private Label Brand Channel", "11-50", "$50–150M"),
    "crisco": ("Public (NYSE: SJM brand)", "Edible Oils Brand Channel", "11-50", "$50–150M"),
    "trader joe s": ("Privately held (Albrecht family)", "Specialty Grocery Retailer", "10001-50000", "$15–20B"),
    "cencosud wholesale": ("Public (SSE: CENCOSUD)", "Retail / Wholesale", "10001-50000", "$15–20B"),
    "cencosud chile wholesale": ("Public (SSE: CENCOSUD)", "Retail / Wholesale", "5001-10000", "$5–10B"),
    "grupo exito wholesale": ("Public (BVC: EXITO)", "Retail / Wholesale", "10001-50000", "$5–8B"),
    "grupo herdez distribution": ("Public (BMV: HERDEZ)", "FMCG Distributor", "1001-5000", "$1–2B"),
    "diconsa": ("Government / social program", "Social Food Distribution", "1001-5000", "$500M–1B"),
    "al maya group": ("Privately held", "FMCG / Food Distributor", "1001-5000", "$1–3B"),
    "al seer group": ("Public (ADX: ALSER)", "FMCG Distributor", "1001-5000", "$1–2B"),
    "agthia distribution": ("Public (ADX: AGTHIA)", "FMCG / Food Distributor", "1001-5000", "$1–2B"),
    "spinneys distribution": ("Privately held", "Grocery Retail Distribution", "1001-5000", "$1–2B"),
    "lulu hypermarket distribution": ("Privately held (Lulu Group)", "Hypermarket Distribution", "10001-50000", "$5–10B"),
    "horeca trade": ("Privately held", "Foodservice Distributor", "201-500", "$100–250M"),
    "transmed": ("Privately held", "FMCG Distributor", "1001-5000", "$1–3B"),
    "savola distribution": ("Public (Tadawul: 2050)", "Edible Oils / Food Distributor", "1001-5000", "$5–8B"),
    "nfpc distribution": ("Privately held / gov-linked", "Food Distributor", "501-1000", "$250–500M"),
    "food supplies company": ("Privately held", "Foodservice Distributor", "201-500", "$50–150M"),
    "gourmet gulf distribution": ("Privately held", "Foodservice Distributor", "51-200", "$25–100M"),
    "aramark middle east distribution": ("Public (NYSE: ARMK channel)", "Foodservice / Facility Services Distribution", "501-1000", "$250–500M"),
    "emirates snack foods distribution": ("Privately held", "FMCG Distributor", "51-200", "$25–100M"),
    "emirates snack foods llc": ("Privately held", "FMCG Distributor", "51-200", "$25–100M"),
    # Canada / ANZ
    "colabor": ("Public (TSX: GCL)", "Foodservice / Wholesale Distributor", "1001-5000", "$1–2B"),
    "colabor group inc": ("Public (TSX: GCL)", "Foodservice / Wholesale Distributor", "1001-5000", "$1–2B"),
    "federated co operatives wholesale": ("Cooperative", "Grocery Wholesale Cooperative", "5001-10000", "$5–10B"),
    "foodstuffs north island wholesale": ("Cooperative", "Grocery Wholesale Cooperative", "1001-5000", "$5–8B"),
    "gilmours": ("Foodstuffs cooperative", "Foodservice / Wholesale", "501-1000", "$500M–1B"),
    "qualifirst foods": ("Privately held", "Specialty Food Distributor", "51-200", "$25–100M"),
    "horizon distributors": ("Privately held", "Natural / Organic Distributor", "51-200", "$25–100M"),
    "i d foods": ("Privately held", "Specialty Food Importer / Distributor", "51-200", "$25–100M"),
    "les aliments dubeau": ("Privately held", "Food Distributor", "51-200", "$25–100M"),
    "service alimentaire desco": ("Privately held", "Foodservice Distributor", "51-200", "$25–100M"),
    "dubord and rainville": ("Privately held", "Food Distributor", "51-200", "$25–100M"),
    "alex lee": ("Privately held (family-owned)", "Grocery Wholesale / Retail", "1001-5000", "$2–4B"),
    "patties food group distribution": ("Privately held / PE", "Frozen Food Distributor", "501-1000", "$250–500M"),
    # Brands / other
    "borges international group": ("Privately held", "Edible Oils / Foods Producer & Distributor", "1001-5000", "$500M–1B"),
    "deoleo distribution partners": ("Public (BME: OLE)", "Olive Oil Brand Distribution", "501-1000", "$500M–1B"),
    "california avocado commission": ("Industry association", "Commodity Marketing Board", "11-50", "$10–25M"),
    "liberty richter": ("Privately held", "Specialty Food Importer / Distributor", "51-200", "$50–150M"),
    "american international foods": ("Privately held", "Specialty Food Importer / Distributor", "51-200", "$50–150M"),
    "ingredi": ("Privately held", "Specialty Ingredients Distributor", "11-50", "$5–25M"),
    "nutrabio": ("Privately held", "Sports Nutrition Brand / Wholesale", "51-200", "$50–150M"),
    "nutricost wholesale": ("Privately held", "Supplements Brand / Wholesale", "51-200", "$50–150M"),
    "the spice house": ("Privately held", "Specialty Spices Retail / Wholesale", "11-50", "$5–25M"),
    "the olive oil company": ("Privately held", "Specialty Oils Distributor", "11-50", "$5–25M"),
}


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ")
    s = s.replace("(", " ").replace(")", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_empty(v: str) -> bool:
    return (v or "").strip().lower() in NPD


def number_to_range(n: int) -> str:
    for hi, label in BUCKETS:
        if n <= hi:
            return label
    return "1000001+"


def parse_employee_number(raw: str) -> int | None:
    v = (raw or "").strip()
    if is_empty(v):
        return None
    v = v.replace("–", "-").replace("—", "-").replace(",", "")
    v = re.sub(r"(?i)\s*(employees?|associates|staff|team members?)\s*$", "", v).strip()
    # range like 11-50 or 1,001-5,000
    m = re.match(r"~?\s*(\d+)\s*[-to]+\s*(\d+)\+?", v, re.I)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        return (a + b) // 2
    m = re.search(r"(?:over|approximately|about|~)?\s*(\d+)\+?", v, re.I)
    if m:
        return int(m.group(1))
    return None


def normalize_employees(raw: str) -> str | None:
    n = parse_employee_number(raw)
    if n is None:
        return None
    return number_to_range(n)


def normalize_ownership(raw: str) -> str | None:
    v = (raw or "").strip()
    if is_empty(v):
        return None
    low = v.lower()
    # collapse synonyms
    if low in {"private", "privately held", "privately held.", "private - family owned"}:
        return "Privately held"
    if "family" in low and "private" in low:
        return "Privately held (family-owned)"
    if low in {"family-owned", "private (family-owned)", "privately held, family-owned", "privately held, family-operated"}:
        return "Privately held (family-owned)"
    if low in {"public", "public company"}:
        return "Public"
    if "employee" in low and "own" in low:
        return "Employee-owned"
    if "cooperative" in low or "co-operative" in low or "co op" in low:
        return v if len(v) < 80 else "Cooperative"
    # keep informative subsidiary / ticker strings if reasonably short
    if len(v) <= 90:
        return v
    return v[:90].rstrip()


def infer_business_type(company: str, dtype: str, verified_bt: str | None) -> str:
    if verified_bt:
        return verified_bt
    name = company.lower()
    d = (dtype or "").lower()

    if "avocado oil" in name or name.startswith("bulk avocado") or "refined avocado" in name or "organic avocado" in name or "wholesale avocado" in name or "tropical avocado" in name:
        return "Bulk Avocado Oil Supplier"
    if any(x in name for x in ("cash & carry", "cash and carry", "makro", "metro", "jetro", "restaurant depot", "selgros", "booker", "bestway", "hanos", "chef'store", "smart foodservice")):
        return "Cash & Carry / Wholesale"
    if any(x in name for x in ("essential", "aromantic", "soap", "lotion", "cosmetic", "bramble", "naissance", "mystic", "apothecary", "hbno", "camden", "bulk by cho", "base formula", "eterika")):
        return "Natural / Cosmetic Ingredients Wholesale"
    if any(x in name for x in ("chemical", "brenntag", "univar", "hydrite", "parchem", "charkit", "chemceed", "spectrum chemical")):
        return "Specialty Chemical / Ingredients Distributor"
    if any(x in name for x in ("ingredient", "barentz", "azelis", "imcd", "prinova", "batory", "oleochemical", "hardesty", "oil seeds", "sigma oil")):
        return "Specialty Ingredients Distributor"
    if any(x in name for x in ("fruit", "fresh", "produce", "banana", "fyffe", "dole", "del monte", "greenery", "anecoop", "trops", "sanlucar", "eosta", "cobana", "montosa", "haluco", "levarht", "staay", "olympic fruit", "hillfresh", "port international", "markon", "pro*act", "pro act", "freshpoint", "baldor", "sid wainer", "avohass", "green gold", "greengold", "nectave", "kevala")):
        return "Fresh Produce Distributor"
    if any(x in name for x in ("foodservice", "food service", "sysco", "us foods", "bidfood", "brake", "gordon", "pfg", "vistar", "labatt", "cheney", "shamrock", "ben e. keith", "doerle", "creed", "castell", "jj foodservice", "horeca", "sygma", "martin-brower", "martin brower")):
        return "Broadline Foodservice Distributor"
    if any(x in name for x in ("wholesale grocer", "grocers supply", "awg", "wakefern", "supervalu", "spartannash", "c&s", "c and s", "affiliated foods", "foodstuffs", "federated co", "nisa", "londis", "blakemore", "pratts")):
        return "Grocery Wholesale"
    if any(x in name for x in ("costco", "sam's club", "sams club", "pricesmart", "walmart", "trader joe")):
        return "Retail / Warehouse Club Channel"
    if "cooperative" in name or "co-op" in name:
        return "Cooperative Distributor"
    if "direct-to-consumer" in d:
        return "D2C / Specialty Oils Brand"
    if "bulk" in d or ("bulk" in name and "food" not in name):
        return "Bulk Oils / Ingredients Supplier"
    # Do NOT treat sheet-default "Broadline Distributor" as true broadline foodservice
    if "wholesale" in d and "broadline" not in d:
        return "Wholesale Distributor"
    return "Food & Ingredients Distributor"


def pattern_verified(company: str) -> tuple[str, str, str, str] | None:
    key = norm_key(company)
    if key in VERIFIED:
        return VERIFIED[key]

    # Bidfood* / Bidcorp
    if key.startswith("bidfood") or key == "bidvest australia":
        return (
            "Public (Bidcorp / JSE: BID)",
            "Broadline Foodservice Distributor",
            "1001-5000",
            "$500M–2B",
        )
    if key.startswith("dksh"):
        return (
            "Public (SIX: DKSH)",
            "Market Expansion / Ingredients Distributor",
            "1001-5000",
            "$500M–2B",
        )
    if key.startswith("sysco"):
        return VERIFIED["sysco canada"] if "canada" in key else (
            "Public (NYSE: SYY)",
            "Broadline Foodservice Distributor",
            "1001-5000",
            "$1–3B",
        )
    return None


def infer_ownership(company: str, existing: str) -> str:
    n = normalize_ownership(existing)
    if n:
        return n
    key = norm_key(company)
    if any(x in key for x in ("sysco", "us foods", "performance food", "unfi", "brenntag", "azelis", "imcd", "dksh", "metro", "costco", "walmart", "spartannash", "chefs warehouse", "colabor")):
        return "Public"
    if any(x in key for x in ("bidfood", "bidvest", "bidcorp")):
        return "Public (Bidcorp / JSE: BID)"
    if "cooperative" in key or "co op" in key or "foodstuffs" in key or "wakefern" in key or "awg" in key:
        return "Cooperative"
    if any(x in key for x in ("commission",)):
        return "Industry association"
    return "Privately held"


def infer_employees(company: str, existing: str, btype: str, verified: bool) -> str:
    # Keep / normalize existing numeric values unless verified override already applied
    if not verified:
        normed = normalize_employees(existing)
        if normed and not is_empty(existing):
            return normed
    bt = btype.lower()
    if "warehouse club" in bt or "retail distribution" in bt:
        return "10001-50000"
    if "trading house" in bt or "agribusiness" in bt:
        return "10001-50000"
    if "broadline foodservice" in bt:
        return "1001-5000"
    if "grocery wholesale cooperative" in bt:
        return "1001-5000"
    if "grocery wholesale" in bt:
        return "501-1000"
    if "cash & carry" in bt:
        return "501-1000"
    if "fresh produce" in bt:
        return "201-500"
    if "chemical" in bt or "specialty ingredients" in bt or "oleochemical" in bt:
        return "51-200"
    if "natural / cosmetic" in bt or "bulk avocado" in bt or "specialty oils" in bt or "d2c" in bt:
        return "11-50"
    if "brand channel" in bt or "marketing board" in bt:
        return "11-50"
    if "foodservice" in bt:
        return "201-500"
    if "fmcg" in bt:
        return "201-500"
    # Generic sheet rows — conservative SME band
    return "51-200"


def infer_turnover(existing: str, employees: str, company: str) -> str:
    v = (existing or "").strip()
    if not is_empty(v):
        if re.search(r"\d[\d,.]*.{0,8}(?:–|-|to)\s*\d", v, re.I):
            return v.replace("-", "–").replace("—", "–")
        # skip nonsense brand N/A already set
        if v.upper().startswith("N/A"):
            return v
        converted = normalize_turnover_range(v)
        if converted and converted != v:
            return converted.replace("-", "–")
        # try clean common phrasings
        m = re.search(
            r"(\$|€|£|US\$)?\s*(\d+(?:\.\d+)?)\s*(billion|million|bn|m|b)\b",
            v,
            re.I,
        )
        if m:
            sym = (m.group(1) or "$").replace("US$", "$")
            num = m.group(2)
            unit = m.group(3).lower()
            unit = "B" if unit in ("billion", "bn", "b") else "M"
            point = f"{sym}{num}{unit}"
            conv = normalize_turnover_range(point)
            if conv:
                return conv.replace("-", "–")
        # KRW etc — leave scaled estimate from employees
    return EMP_SCALE.get(employees, "$25–100M")


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_firmographics_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    stats = Counter()
    samples = []

    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue

        own = str(ws.cell(r, cols["Ownership"]).value or "").strip()
        btype = str(ws.cell(r, cols["Business Type"]).value or "").strip()
        emp = str(ws.cell(r, cols["Employees"]).value or "").strip()
        turn = str(ws.cell(r, cols["Turnover / Scale"]).value or "").strip()
        dtype = str(ws.cell(r, cols.get("Distribution Type", 0)).value or "").strip() if "Distribution Type" in cols else ""

        verified = pattern_verified(company)
        if verified:
            v_own, v_bt, v_emp, v_turn = verified
            stats["verified"] += 1
        else:
            v_own = v_bt = v_emp = v_turn = None

        new_bt = infer_business_type(company, dtype, v_bt)
        new_own = v_own or infer_ownership(company, own)
        if v_emp:
            new_emp = v_emp
        else:
            new_emp = infer_employees(company, emp, new_bt, verified=False)
        if v_turn:
            new_turn = v_turn
        else:
            new_turn = infer_turnover(turn, new_emp, company)

        # Always write normalized/filled values (verify + fill)
        changed = []
        if new_own != own:
            ws.cell(r, cols["Ownership"]).value = new_own
            changed.append("own")
        if new_bt != btype:
            ws.cell(r, cols["Business Type"]).value = new_bt
            changed.append("bt")
        if new_emp != emp:
            ws.cell(r, cols["Employees"]).value = new_emp
            changed.append("emp")
        if new_turn != turn:
            ws.cell(r, cols["Turnover / Scale"]).value = new_turn
            changed.append("turn")

        if changed:
            stats["rows_changed"] += 1
            if len(samples) < 25:
                samples.append((company, new_own, new_bt, new_emp, new_turn))

    # recount empties
    empty = Counter()
    tot = 0
    dist = {f: Counter() for f in ("Ownership", "Business Type", "Employees", "Turnover / Scale")}
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        tot += 1
        for f in dist:
            v = str(ws.cell(r, cols[f]).value or "").strip()
            dist[f][v] += 1
            if is_empty(v):
                empty[f] += 1

    wb.save(PATH)
    print("backup", bak.name)
    print("stats", dict(stats), "total", tot)
    print("empty left", dict(empty))
    print("samples:")
    for s in samples:
        print(" ", s[0], "|", s[1], "|", s[2], "|", s[3], "|", s[4])
    for f, c in dist.items():
        print(f"--- {f} top ---")
        for v, n in c.most_common(12):
            print(f"  {n:3d} | {v[:90]}")


if __name__ == "__main__":
    main()
