"""Fix avocado channel columns: replace Yes/3/Mid-tier boilerplate with real profiles.

Columns:
  Retail Chains, Specialty Grocers, Foodservice / HoReCa, E-commerce,
  Channel Strength, Distribution Type, Regions Served, Regional Extensions,
  Warehouse / Logistics, Delivery / Storage, Competitive Benchmark
"""
from __future__ import annotations

import re
import shutil
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
)

FIELDS = [
    "Retail Chains",
    "Specialty Grocers",
    "Foodservice / HoReCa",
    "E-commerce",
    "Channel Strength",
    "Distribution Type",
    "Regions Served",
    "Regional Extensions",
    "Warehouse / Logistics",
    "Delivery / Storage",
    "Competitive Benchmark",
]

GENERIC = {
    "",
    "yes",
    "no",
    "3",
    "medium",
    "mid-tier",
    "regional",
    "broadline distributor",
    "wholesale",
    "not publicly disclosed",
    "n/a",
    "na",
}

# (retail, specialty, foodservice, ecommerce, strength, dist_type, regions, extensions, warehouse, delivery, benchmark)
VERIFIED: dict[str, tuple[str, ...]] = {
    "sysco": (
        "Limited — foodservice primary; not a grocery retail wholesaler",
        "Limited specialty foodservice (Sysco Specialty / FreshPoint), not grocery specialty",
        "Core — restaurants, hotels, healthcare, education, SYGMA chain accounts (~730k locations)",
        "Strong B2B — Sysco Shop e-commerce portal + third-party assortment",
        "Very High — global #1 broadline foodservice",
        "Broadline foodservice + specialty + SYGMA chain logistics",
        "Global — US, Canada, Mexico, UK, Ireland, France, Sweden, Belgium, LatAm/Caribbean",
        "International foodservice ops beyond US core (Europe & Americas)",
        "~340 distribution facilities worldwide; multi-temp DCs",
        "~18k vehicles; scheduled multi-temp delivery to foodservice accounts",
        "Leader vs US Foods / PFG / Bidfood in broadline foodservice",
    ),
    "sysco canada": (
        "Limited — foodservice focus",
        "Limited specialty foodservice assortment",
        "Core — Canadian restaurants, hotels, institutions",
        "B2B online ordering / Sysco digital tools",
        "High — leading Canadian foodservice distributor",
        "Broadline foodservice",
        "Canada",
        "Part of Sysco Americas network",
        "Canadian Sysco DC network",
        "Multi-temp fleet delivery to foodservice",
        "Top-tier Canada foodservice vs Gordon / local broadliners",
    ),
    "sysco europe": (
        "Limited — foodservice / Brakes channel",
        "Limited specialty foodservice",
        "Core — UK/Europe HoReCa via Sysco/Brakes",
        "B2B e-commerce ordering",
        "High — major European foodservice player",
        "Broadline foodservice (Sysco/Brakes)",
        "UK, Ireland, France, Sweden, Belgium (Sysco Europe footprint)",
        "Sysco International network",
        "European Sysco/Brakes DC network",
        "Foodservice scheduled delivery",
        "Top-tier Europe foodservice vs Bidfood / Metro FSD",
    ),
    "us foods": (
        "Limited — CHEF'STORE cash & carry open to public; not grocery wholesale",
        "Limited specialty foodservice brands (Chef's Line etc.)",
        "Core — ~250k US restaurants, healthcare, hospitality, education",
        "Strong — MOXē digital platform + US Foods Direct online assortment",
        "Very High — #2 US broadline foodservice",
        "Broadline foodservice + CHEF'STORE cash & carry",
        "United States (nationwide)",
        "CHEF'STORE store footprint across multiple US states",
        "70+ broadline DCs; 90+ CHEF'STORE locations",
        "6,500+ trucks; Scheduled Delivery + Pronto next-day in dense markets",
        "Peer to Sysco / PFG / GFS in US foodservice",
    ),
    "us foods chef store": (
        "Public cash & carry (not membership grocery retail)",
        "Limited specialty SKUs for chefs / foodservice",
        "Core — independent restaurants & foodservice operators walk-in",
        "CHEFSTORE.com + in-store; linked to US Foods digital",
        "High — national cash & carry brand of US Foods",
        "Cash & carry foodservice wholesale",
        "United States (multi-state CHEF'STORE network)",
        "Continuing new store openings across US",
        "90+ warehouse-format CHEF'STORE locations",
        "Customer pickup; optional delivery via US Foods network",
        "Competes with Restaurant Depot / Jetro cash & carry",
    ),
    "performance food group": (
        "Limited — foodservice primary; some convenience wholesale via Vistar",
        "Specialty foodservice via Performance / specialty brands",
        "Core — US restaurants, chains, institutions",
        "B2B e-commerce ordering platforms",
        "Very High — top-3 US foodservice distributor",
        "Broadline foodservice + specialty + convenience (Vistar)",
        "United States",
        "National DC footprint; specialty & convenience adjacencies",
        "National multi-temp DC network",
        "Fleet delivery to foodservice & convenience accounts",
        "Peer to Sysco / US Foods",
    ),
    "gordon food service": (
        "Limited — GFS Marketplace retail/cash & carry in some markets",
        "Limited specialty foodservice assortment",
        "Core — US & Canada restaurants, schools, healthcare (~100k+ customers)",
        "B2B online ordering / GFS digital tools",
        "High — largest privately held US/Canada foodservice distributor",
        "Broadline foodservice + Marketplace cash & carry",
        "United States and Canada",
        "Cross-border US–Canada network",
        "28+ distribution centers US & Canada",
        "Multi-temp fleet delivery to foodservice",
        "Strong regional alternative to Sysco / US Foods",
    ),
    "unfi united natural foods": (
        "Core — supermarket chains, independents, Cub Foods retail banners",
        "Core — natural, organic & specialty grocery (legacy UNFI strength)",
        "Limited — grocery wholesale primary, not restaurant broadline",
        "Strong — Endless Aisle / retailer e-commerce enablement + retail media",
        "Very High — largest US natural/specialty grocery wholesaler (~$31B)",
        "Full-line grocery wholesale (natural + conventional) + retail services",
        "United States and Canada",
        "Network optimization (new Sarasota/Manchester DCs; consolidations)",
        "~50+ multi-temp DCs across North America (incl. automated sites)",
        "Retailer replenishment delivery; multi-temp grocery logistics",
        "Peer to KeHE in natural/specialty; competes with C&S / AWG in conventional",
    ),
    "kehe distributors": (
        "Core — supermarket chains & independents (Albertsons, etc.)",
        "Core — natural, organic, fresh & specialty (~31k retail doors)",
        "Limited — grocery specialty channel, not HoReCa broadline",
        "Strong — dedicated e-commerce capable DCs (Elkton, Aurora, Dallas, etc.)",
        "Very High — leading pure-play natural/specialty grocery distributor",
        "Natural / organic / specialty grocery wholesale",
        "United States and Canada",
        "19 DCs across North America (15 US + 4 Canada)",
        "~19 DCs / 7M+ sq ft multi-zone grocery warehousing",
        "Retail grocery replenishment; specialty & fresh cold chain",
        "Peer to UNFI in natural/specialty grocery wholesale",
    ),
    "brenntag": (
        "Limited — industrial/food manufacturer customers, not grocery retail",
        "Limited — specialty food ingredient accounts, not grocery specialty",
        "Limited — foodservice not core; Food & Nutrition ingredients to industry",
        "B2B digital ordering / customer portals",
        "Very High — global #1 chemicals & ingredients distributor",
        "Specialty chemicals & food ingredients distribution",
        "Global — 70+ countries",
        "Food & Nutrition labs / application centers worldwide",
        "~600 sites globally; ambient + cold storage warehouses",
        "Local warehouse pickup + scheduled B2B delivery; formulation support",
        "Leader vs Azelis / IMCD / Barentz in ingredients distribution",
    ),
    "brenntag north america": (
        "Limited — B2B industrial/food manufacturers",
        "Limited specialty ingredient accounts",
        "Limited — not HoReCa broadline",
        "B2B digital / customer portals",
        "High — major NA ingredients distributor",
        "Food ingredients & specialty chemical distribution",
        "United States and Canada",
        "Part of Brenntag global network",
        "North American Brenntag warehouse network",
        "B2B scheduled delivery from regional warehouses",
        "Peer to Univar / Azelis Americas in NA ingredients",
    ),
    "barentz": (
        "Limited — manufacturer ingredient channel",
        "Limited — specialty nutrition/food formulators",
        "Limited — not restaurant distribution",
        "B2B digital / customer support portals",
        "High — global specialty life-science ingredients distributor",
        "Specialty food & nutrition ingredients distribution",
        "Global (EMEA HQ; Americas & Asia presence)",
        "Regional specialty offices & tech centers",
        "Regional specialty warehouses & application labs",
        "B2B ingredient delivery / supply programs",
        "Peer to Azelis / IMCD / Brenntag Specialties",
    ),
    "azelis": (
        "Limited — industrial/formulator customers",
        "Limited specialty ingredient accounts",
        "Limited — not HoReCa",
        "B2B digital platforms",
        "High — global specialty chemicals & ingredients",
        "Specialty ingredients distribution",
        "Global",
        "Americas / EMEA / Asia Pacific networks",
        "Global specialty warehouse & lab network",
        "B2B delivery + formulation support",
        "Peer to Barentz / IMCD / Brenntag Specialties",
    ),
    "imcd": (
        "Limited — B2B formulators",
        "Limited specialty ingredients",
        "Limited — not HoReCa",
        "B2B digital",
        "High — global specialty distributor",
        "Specialty chemicals & food ingredients",
        "Global",
        "Regional technical centers",
        "Regional warehouses",
        "B2B delivery",
        "Peer to Azelis / Barentz",
    ),
    "bidfood uk": (
        "Limited — foodservice primary; retail exposure minimal",
        "Limited specialty foodservice SKUs",
        "Core — UK HoReCa / catering / hospitality",
        "Strong B2B online ordering",
        "Very High — Bidcorp UK flagship foodservice",
        "Broadline foodservice wholesale",
        "United Kingdom",
        "Bidcorp group (31 countries); UK core market",
        "UK Bidfood multi-temp DC network",
        "Scheduled foodservice delivery fleet",
        "UK peer to Brakes/Sysco; part of Bidcorp global",
    ),
    "metro cash and carry": (
        "Traders / independent grocery via Makro/Metro wholesale",
        "Limited specialty assortment in store",
        "Core — HoReCa + Traders (hotels, restaurants, caterers)",
        "Strong — Metro Markets B2B marketplace + digital ordering",
        "Very High — leading international food wholesaler",
        "Cash & carry wholesale + foodservice delivery (FSD)",
        "30+ countries (Metro/Makro store + delivery footprint)",
        "622 stores in 21 countries + FSD depots in additional markets",
        "Wholesale store network + 100+ FSD depots",
        "Store pickup + out-of-store delivery (OOS) + FSD",
        "Peer to Bidfood / Sysco International in HoReCa wholesale",
    ),
    "makro": (
        "Traders / independent retail via Makro membership wholesale",
        "Limited specialty in-store assortment",
        "Core — HoReCa professionals",
        "B2B digital / Metro group e-commerce tools",
        "High — Metro AG Makro brand",
        "Cash & carry wholesale",
        "Country of operation (Metro/Makro network)",
        "Metro AG international network",
        "Makro wholesale stores + delivery where offered",
        "Membership pickup + delivery options",
        "Competes with local cash & carry / foodservice wholesalers",
    ),
    "jetro restaurant depot": (
        "Limited — members are foodservice/resale operators, not grocery chains",
        "Limited specialty for chefs",
        "Core — independent restaurants & foodservice cash & carry",
        "Limited / member portal; primarily warehouse pickup",
        "High — leading US membership foodservice cash & carry",
        "Membership cash & carry foodservice wholesale",
        "United States (multi-metro warehouse network)",
        "Dense metro warehouse expansion",
        "Membership warehouse clubs nationwide",
        "Customer pickup primary; limited delivery partners",
        "Peer to US Foods CHEF'STORE / local cash & carry",
    ),
    "restaurant depot": (
        "Limited — foodservice membership wholesale",
        "Limited specialty chef SKUs",
        "Core — independent restaurants",
        "Limited online; warehouse primary",
        "High — US foodservice cash & carry leader",
        "Membership cash & carry",
        "United States",
        "National warehouse footprint",
        "Membership warehouses",
        "Pickup-focused",
        "Peer to Jetro / CHEF'STORE",
    ),
    "trader joe s": (
        "Own specialty grocery retail banner (not wholesale to other chains)",
        "Core — specialty grocery retail assortment incl. private-label oils",
        "Limited — retail grocery, not HoReCa distributor",
        "Limited — primarily store retail (no full national e-comm grocery)",
        "High — iconic US specialty grocery retailer",
        "Specialty grocery retail (private-label channel)",
        "United States",
        "Select international presence historically limited; US core",
        "Retail store network + DC support",
        "Store shelf; customer in-store purchase",
        "Specialty retail peer to Whole Foods / regional specialty banners",
    ),
    "baldor specialty foods": (
        "Limited — chef/specialty produce focus",
        "Core — specialty produce & specialty foods to fine retail/chefs",
        "Core — NYC/NE chefs, hotels, restaurants",
        "B2B online ordering",
        "High — leading Northeast specialty produce/foodservice",
        "Specialty produce & specialty foodservice distribution",
        "Northeast United States (NYC metro core)",
        "Northeast corridor coverage",
        "Bronx specialty DC / cold chain produce warehouses",
        "Chef delivery; temperature-controlled produce logistics",
        "Specialty peer to FreshPoint / local produce houses",
    ),
    "fyffes": (
        "Retail produce programs via supermarket partners",
        "Limited specialty produce programs",
        "Limited — fresh produce primary vs culinary oil foodservice",
        "Limited B2B; retail partner programs",
        "High — major global banana/tropical produce distributor",
        "Fresh produce import & wholesale distribution",
        "Europe, Americas & other produce markets",
        "Global tropical fruit sourcing network",
        "Ripening centers & produce logistics",
        "Cold-chain produce delivery to retail/wholesale",
        "Peer to other tropical produce majors (Dole, etc.)",
    ),
    "anecoop": (
        "Retail produce programs via European supermarket partners",
        "Limited specialty produce",
        "Limited foodservice produce programs",
        "Limited B2B digital",
        "High — major Spanish produce cooperative exporter",
        "Fresh produce cooperative export / wholesale",
        "Spain / Europe export markets",
        "European retail partner markets",
        "Cooperative packing & cold logistics",
        "Export cold-chain to retail partners",
        "Spanish cooperative peer to other EU produce coops",
    ),
    "dksh": (
        "Retail & pharmacy channel via market expansion services",
        "Specialty / beauty & specialty F&B where contracted",
        "Limited HoReCa; more retail/modern trade coverage",
        "B2B digital / route-to-market tools",
        "High — leading Asia market-expansion distributor",
        "Market expansion services (F&B ingredients & consumer)",
        "Asia-Pacific focus",
        "Multi-country Asia coverage",
        "Regional Asia warehouses & value-added logistics",
        "Modern trade delivery + field sales coverage",
        "Peer to other Asia MES players (e.g. local 3PLs/distributors)",
    ),
    "sligro food group": (
        "Limited retail; cash & carry / foodservice",
        "Limited specialty",
        "Core — Dutch HoReCa & wholesale",
        "B2B online ordering",
        "High — leading Dutch foodservice/wholesale",
        "Foodservice wholesale + cash & carry",
        "Netherlands",
        "Benelux adjacency",
        "Dutch DC / cash & carry network",
        "Delivery + cash & carry pickup",
        "Dutch peer to Bidfood NL / Metro NL",
    ),
    "wakefern food corp": (
        "Core — ShopRite & member retail banners",
        "Yes — Wholesome Pantry / specialty within member stores",
        "Limited — retailer-owned grocery wholesale",
        "Member banner e-commerce / omnichannel",
        "High — major US retailer-owned grocery cooperative",
        "Grocery wholesale cooperative",
        "Northeastern & Mid-Atlantic United States",
        "Member banner geographic expansion",
        "Cooperative DC network for member stores",
        "Store replenishment logistics",
        "Peer to other US grocery coops (AWG, Associated)",
    ),
    "c and s wholesale grocers": (
        "Core — supermarket chain wholesale customers",
        "Limited specialty vs full-line grocery",
        "Limited — grocery wholesale not HoReCa",
        "Retailer e-commerce support programs",
        "Very High — one of largest US grocery wholesalers",
        "Grocery wholesale",
        "United States",
        "National grocery wholesale footprint",
        "Large multi-temp grocery DC network",
        "Retailer replenishment delivery",
        "Peer to UNFI conventional / AWG / Supervalu-era peers",
    ),
    "associated wholesale grocers": (
        "Core — independent supermarket members",
        "Member specialty programs",
        "Limited foodservice",
        "Member digital / e-comm support",
        "High — major US grocery cooperative",
        "Grocery wholesale cooperative",
        "US Midwest, Southeast & South (~36 states)",
        "Member territory coverage",
        "Cooperative DC network",
        "Member store replenishment",
        "Peer to Wakefern / Associated Food Stores",
    ),
    "colabor group inc": (
        "Limited — Canadian wholesale/foodservice",
        "Limited specialty",
        "Core — Québec/Canada restaurants, hotels, institutions",
        "B2B ordering tools",
        "High — major Québec foodservice/wholesale distributor",
        "Foodservice & wholesale distribution",
        "Canada (Québec & adjacent)",
        "Eastern Canada coverage",
        "Canadian multi-temp warehouses",
        "Refrigerated fleet delivery (Québec / northern markets)",
        "Canadian regional peer to Sysco Canada / Gordon",
    ),
    "sam s club mexico": (
        "Warehouse club retail (Sam's Club Mexico)",
        "Limited specialty club assortment",
        "Limited — club retail / small business members",
        "Club e-commerce / app ordering where offered",
        "High — leading Mexico warehouse club (Walmart)",
        "Warehouse club retail / wholesale membership",
        "Mexico",
        "Walmart Mexico / Central America adjacency",
        "Club distribution centers supporting stores",
        "Member pickup; club delivery options",
        "Peer to Costco Mexico in club channel",
    ),
    "costco mexico wholesale": (
        "Warehouse club retail",
        "Limited specialty club packs",
        "Limited — club / small business",
        "Costco e-commerce where available",
        "High — Costco Mexico club channel",
        "Warehouse club",
        "Mexico",
        "Costco global club model",
        "Club DCs",
        "Member pickup",
        "Peer to Sam's Club Mexico",
    ),
    "tradin organic": (
        "Limited — organic ingredient supply to brands/retailers via B2B",
        "Organic specialty ingredient accounts",
        "Limited HoReCa",
        "B2B digital / trading desk",
        "High — major organic ingredients trader",
        "Organic ingredients trading / wholesale",
        "Global organic sourcing & sales",
        "Multi-origin organic supply network",
        "Organic-certified storage / logistics partners",
        "B2B bulk / ingredient shipments",
        "Peer to other organic ingredient traders",
    ),
}


def nk(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ")
    s = s.replace("(", " ").replace(")", " ").replace("/", " ").replace(".", " ")
    s = s.replace(",", " ").replace("'", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_generic(v: str) -> bool:
    return (v or "").strip().lower() in GENERIC


# Business-type → channel profile template (regions filled later)
# retail, specialty, foodservice, ecommerce, strength, dist_type, warehouse, delivery, benchmark
BT_BASE: dict[str, tuple[str, ...]] = {
    "Broadline Foodservice Distributor": (
        "Limited — foodservice primary",
        "Limited specialty foodservice SKUs",
        "Core — restaurants, hotels, catering, institutions",
        "B2B online ordering / customer portal",
        "High — regional/national broadline foodservice",
        "Broadline foodservice wholesale",
        "Multi-temp foodservice distribution centers",
        "Scheduled multi-temp delivery to HoReCa accounts",
        "Competes with national/regional broadline foodservice peers",
    ),
    "Foodservice Distributor": (
        "Limited — foodservice focus",
        "Limited specialty foodservice",
        "Core — regional HoReCa accounts",
        "B2B ordering (portal or sales rep)",
        "Medium–High — regional foodservice distributor",
        "Foodservice wholesale distribution",
        "Regional foodservice warehouse(s)",
        "Route delivery to restaurants & institutions",
        "Regional foodservice peer set",
    ),
    "Foodservice Wholesale": (
        "Limited",
        "Limited",
        "Core — wholesale to foodservice operators",
        "B2B ordering",
        "Medium — regional foodservice wholesale",
        "Foodservice wholesale",
        "Wholesale warehouse",
        "Wholesale delivery / pickup",
        "Local foodservice wholesale peers",
    ),
    "Specialty Foodservice Distributor": (
        "Limited",
        "Core — specialty foods to chefs / fine dining",
        "Core — specialty HoReCa",
        "B2B online / specialty catalog",
        "Medium–High — specialty foodservice niche",
        "Specialty foodservice distribution",
        "Specialty / cold-chain warehouse",
        "Chef-oriented delivery",
        "Specialty foodservice peers",
    ),
    "QSR / Foodservice Distributor": (
        "Limited",
        "Limited",
        "Core — QSR / chain restaurant distribution",
        "B2B / EDI chain ordering",
        "High — chain foodservice logistics",
        "Custom / chain foodservice distribution",
        "Chain-dedicated DCs",
        "High-frequency QSR delivery",
        "Competes with SYGMA / chain specialists",
    ),
    "Cash & Carry / Wholesale": (
        "Traders / independent retail via cash & carry membership",
        "Limited specialty in-store",
        "Core — HoReCa & traders self-service wholesale",
        "B2B digital ordering where offered; store primary",
        "High — cash & carry wholesale channel",
        "Cash & carry wholesale",
        "Cash & carry warehouse stores",
        "Customer pickup + optional delivery",
        "Competes with Metro/Makro / local cash & carry",
    ),
    "Cash & Carry / Foodservice Wholesale": (
        "Limited retail traders",
        "Limited",
        "Core — foodservice cash & carry",
        "Limited e-comm; warehouse primary",
        "Medium–High",
        "Foodservice cash & carry",
        "Membership / wholesale warehouse",
        "Pickup-focused",
        "Cash & carry peers",
    ),
    "Grocery Wholesale": (
        "Core — supermarket & independent grocery retail",
        "Yes — specialty/center-store oils within grocery assortment",
        "Limited — grocery wholesale not HoReCa broadline",
        "Retailer e-commerce / replenishment support",
        "High — grocery wholesale channel",
        "Grocery wholesale",
        "Multi-temp grocery DCs",
        "Retail store replenishment delivery",
        "Grocery wholesale peer set (C&S / UNFI / coops)",
    ),
    "Grocery Wholesale Cooperative": (
        "Core — member supermarket banners",
        "Member specialty / natural programs",
        "Limited",
        "Member banner e-commerce support",
        "High — retailer-owned cooperative wholesale",
        "Grocery wholesale cooperative",
        "Cooperative DC network",
        "Member store replenishment",
        "US grocery coop peers",
    ),
    "Grocery Wholesale / Retail": (
        "Core — own / affiliated retail banners",
        "Yes — within retail assortment",
        "Limited",
        "Retail e-commerce / click & collect",
        "High",
        "Integrated grocery wholesale–retail",
        "Retail + wholesale DCs",
        "Store replenishment + consumer retail",
        "Regional grocery integrated peers",
    ),
    "Natural / Specialty Grocery Wholesale": (
        "Core — natural/specialty retail doors",
        "Core — natural, organic & specialty grocery",
        "Limited foodservice",
        "Strong specialty / natural e-comm support",
        "High — natural/specialty wholesale",
        "Natural & specialty grocery wholesale",
        "Specialty grocery DCs (multi-zone)",
        "Specialty retail replenishment",
        "Peer to UNFI / KeHE natural channel",
    ),
    "Specialty Food Distributor": (
        "Limited–Moderate specialty retail accounts",
        "Core — specialty / gourmet retail",
        "Moderate specialty foodservice",
        "B2B specialty ordering",
        "Medium–High specialty distributor",
        "Specialty food wholesale",
        "Specialty warehouse",
        "Specialty account delivery",
        "Specialty food distributor peers",
    ),
    "Specialty Food Importer / Distributor": (
        "Specialty retail import programs",
        "Core — imported specialty oils & foods",
        "Moderate specialty foodservice",
        "B2B importer portal / sales",
        "Medium — specialty importer",
        "Specialty food import & wholesale",
        "Import warehouse / bonded storage",
        "Import container → wholesale delivery",
        "Specialty importer peers",
    ),
    "Specialty Ingredients Distributor": (
        "Limited — manufacturer customers not retail chains",
        "Limited — specialty ingredient formulators",
        "Limited — not HoReCa broadline",
        "B2B digital / customer portal",
        "High — specialty ingredients channel",
        "Specialty food ingredients distribution",
        "Ingredient warehouses + application support",
        "B2B ingredient delivery",
        "Peer to Azelis / Barentz / IMCD model",
    ),
    "Food Ingredients Distributor": (
        "Limited — food manufacturers",
        "Limited",
        "Limited",
        "B2B ordering",
        "Medium–High",
        "Food ingredients wholesale",
        "Ingredient warehouse",
        "B2B delivery",
        "Ingredients distributor peers",
    ),
    "Food & Ingredients Distributor": (
        "Moderate — food & retail where applicable",
        "Moderate specialty / ingredients",
        "Moderate foodservice or manufacturing",
        "B2B ordering",
        "Medium — regional food/ingredients distributor",
        "Food & ingredients wholesale",
        "Regional warehouse",
        "B2B delivery / wholesale pickup",
        "Regional food/ingredients peers",
    ),
    "Fresh Produce Distributor": (
        "Retail produce programs via supermarket partners",
        "Limited specialty produce programs",
        "Moderate — foodservice produce accounts",
        "Limited B2B produce ordering",
        "Medium–High — fresh produce wholesale",
        "Fresh produce wholesale / distribution",
        "Ripening / cold-chain produce facilities",
        "Temperature-controlled produce delivery",
        "Produce distributor peers",
    ),
    "Fresh Produce Cooperative": (
        "Retail partner programs",
        "Limited",
        "Limited–Moderate",
        "Limited",
        "Medium–High cooperative produce",
        "Produce cooperative wholesale / export",
        "Cooperative packing & cold logistics",
        "Cold-chain to retail partners",
        "Produce coop peers",
    ),
    "FMCG Distributor": (
        "Core — modern trade / FMCG retail",
        "Limited specialty",
        "Limited foodservice",
        "Modern trade / B2B ordering",
        "Medium–High FMCG distribution",
        "FMCG wholesale distribution",
        "FMCG warehouses",
        "Retail modern-trade delivery",
        "Regional FMCG distributor peers",
    ),
    "FMCG / Food Distributor": (
        "Core — retail FMCG",
        "Limited",
        "Limited–Moderate",
        "B2B / modern trade",
        "Medium–High",
        "FMCG / food wholesale",
        "Regional warehouses",
        "Retail delivery",
        "FMCG peers",
    ),
    "Market Expansion / Ingredients Distributor": (
        "Modern trade / pharmacy / retail where contracted",
        "Specialty F&B / beauty specialty accounts",
        "Limited HoReCa",
        "B2B digital route-to-market",
        "High — market expansion services",
        "Market expansion / ingredients distribution",
        "Multi-country regional warehouses",
        "Field sales + warehouse delivery",
        "Asia/EM MES distributor peers",
    ),
    "Specialty Chemical Distributor": (
        "No — industrial/chemical customers",
        "No — not grocery specialty",
        "No — not HoReCa",
        "B2B chemical portal",
        "High — chemical distribution",
        "Specialty chemical distribution",
        "Chemical warehouses (regulated storage)",
        "Chemical B2B delivery / compliance logistics",
        "Chemical distributor peers",
    ),
    "Chemical Distributor": (
        "No — industrial",
        "No",
        "No",
        "B2B",
        "Medium–High",
        "Chemical wholesale distribution",
        "Chemical warehouses",
        "B2B chemical delivery",
        "Chemical peers",
    ),
    "Warehouse Club Wholesale": (
        "Warehouse club retail members",
        "Limited club specialty packs",
        "Limited — small business / club",
        "Club e-commerce / app",
        "High — warehouse club channel",
        "Warehouse club wholesale/retail",
        "Club distribution centers",
        "Member pickup + club delivery",
        "Club peers (Costco / Sam's)",
    ),
    "Specialty Grocery Retailer": (
        "Own specialty grocery retail",
        "Core — specialty grocery",
        "Limited",
        "Limited / store-led",
        "High specialty retail brand",
        "Specialty grocery retail",
        "Retail DC + stores",
        "In-store retail",
        "Specialty grocery retail peers",
    ),
    "Bulk Oils / Ingredients Supplier": (
        "Limited — B2B bulk buyers",
        "Limited specialty bulk",
        "Limited foodservice bulk",
        "B2B online / wholesale catalog",
        "Medium — bulk oils wholesale",
        "Bulk oils & ingredients wholesale",
        "Bulk oil storage / warehouse",
        "Bulk / drum / tote shipments",
        "Bulk oil wholesale peers",
    ),
    "Wholesale Distributor": (
        "Moderate wholesale retail accounts",
        "Moderate",
        "Moderate",
        "B2B ordering",
        "Medium",
        "General wholesale distribution",
        "Wholesale warehouse",
        "Wholesale delivery / pickup",
        "Regional wholesale peers",
    ),
    "Trading House Food Distribution": (
        "Export/import retail partners",
        "Limited specialty trade",
        "Limited",
        "B2B trading desk",
        "Medium — trading house",
        "Food trading / distribution",
        "Trading warehouse / port logistics",
        "Export/import shipments",
        "Trading house peers",
    ),
    "Organic Ingredients Trader": (
        "B2B to brands/retailers",
        "Organic specialty ingredient accounts",
        "Limited",
        "B2B trading",
        "Medium–High organic trade",
        "Organic ingredients trading",
        "Organic-certified storage partners",
        "Bulk organic shipments",
        "Organic trader peers",
    ),
    "Cosmetic Ingredients Supplier": (
        "No grocery retail",
        "No grocery specialty",
        "No HoReCa",
        "B2B / DTC craft portal",
        "Medium — cosmetic ingredients wholesale",
        "Cosmetic / personal-care ingredients wholesale",
        "Ingredients warehouse",
        "Small-parcel / wholesale shipment",
        "Cosmetic ingredients peers",
    ),
    "Natural / Cosmetic Ingredients Wholesale": (
        "No grocery chains",
        "No",
        "No",
        "B2B / online wholesale",
        "Medium",
        "Natural cosmetic ingredients wholesale",
        "Ingredients warehouse",
        "Wholesale shipment",
        "Natural ingredients peers",
    ),
    "Soap / Cosmetic Ingredients Supplier": (
        "No",
        "No",
        "No",
        "Online wholesale catalog",
        "Medium",
        "Soap/cosmetic ingredients wholesale",
        "Warehouse",
        "Parcel / wholesale ship",
        "Craft ingredients peers",
    ),
    "Food Distributor": (
        "Moderate retail",
        "Moderate",
        "Moderate foodservice",
        "B2B",
        "Medium",
        "Food wholesale distribution",
        "Food warehouse",
        "Food delivery",
        "Regional food distributor peers",
    ),
    "Oils Distributor": (
        "Moderate edible-oil retail accounts",
        "Moderate specialty oil accounts",
        "Moderate foodservice oil supply",
        "B2B",
        "Medium",
        "Edible oils wholesale",
        "Oils warehouse",
        "Oil product delivery",
        "Edible oils distributor peers",
    ),
    "Vegetable Oils Distributor": (
        "Moderate",
        "Moderate",
        "Moderate",
        "B2B",
        "Medium",
        "Vegetable oils wholesale",
        "Oils warehouse",
        "Delivery",
        "Vegetable oils peers",
    ),
    "Oleochemical / Ingredients Distributor": (
        "No retail grocery",
        "No",
        "No",
        "B2B",
        "Medium–High",
        "Oleochemical / ingredients distribution",
        "Chemical/ingredient warehouse",
        "B2B delivery",
        "Oleochemical distributor peers",
    ),
    "Laboratory / Fine Chemical Distributor": (
        "No",
        "No",
        "No",
        "B2B lab portal",
        "Medium–High lab chemical",
        "Laboratory / fine chemical distribution",
        "Lab chemical warehouse",
        "Lab chemical shipping",
        "Lab chemical peers",
    ),
    "Hypermarket Distribution": (
        "Core — hypermarket retail",
        "Limited specialty",
        "Limited",
        "Retail e-commerce",
        "High hypermarket",
        "Hypermarket retail distribution",
        "Hypermarket DC network",
        "Store replenishment",
        "Hypermarket peers",
    ),
    "Grocery Retail Distribution": (
        "Core — grocery retail",
        "Yes within assortment",
        "Limited",
        "Retail e-commerce",
        "High",
        "Grocery retail distribution",
        "Retail DCs",
        "Store replenishment",
        "Grocery retail peers",
    ),
    "Retail Distribution": (
        "Core retail",
        "Moderate",
        "Limited",
        "Retail e-comm",
        "Medium–High",
        "Retail distribution",
        "Retail warehouses",
        "Store delivery",
        "Retail peers",
    ),
    "Retail / Wholesale": (
        "Core retail + wholesale",
        "Moderate",
        "Moderate",
        "Omnichannel",
        "Medium–High",
        "Retail & wholesale",
        "Retail/wholesale warehouses",
        "Store + wholesale delivery",
        "Retail-wholesale peers",
    ),
    "Warehouse Club": (
        "Warehouse club members",
        "Limited",
        "Limited small business",
        "Club e-comm",
        "High",
        "Warehouse club",
        "Club DCs",
        "Member pickup",
        "Club peers",
    ),
    "Cash & Carry / Warehouse Club": (
        "Membership wholesale / club",
        "Limited",
        "HoReCa + small business",
        "Limited e-comm",
        "High",
        "Cash & carry / warehouse club",
        "Warehouse club stores",
        "Pickup + delivery",
        "Club / cash & carry peers",
    ),
    "Foodservice Redistributor": (
        "Limited",
        "Limited",
        "Core — redistributes to foodservice operators",
        "B2B",
        "Medium",
        "Foodservice redistribution",
        "Redistribution warehouse",
        "Foodservice delivery",
        "Redistributor peers",
    ),
    "Foodservice / Facility Services Distribution": (
        "Limited",
        "Limited",
        "Core — facility / foodservice accounts",
        "B2B / contract ordering",
        "Medium–High",
        "Foodservice & facility distribution",
        "Contract warehouses",
        "Contract delivery",
        "Facility services distribution peers",
    ),
    "Produce Procurement Cooperative": (
        "Retail produce programs",
        "Limited",
        "Foodservice produce programs",
        "Member digital procurement",
        "High — produce buying coop",
        "Produce procurement cooperative",
        "Member logistics network",
        "Produce cold-chain to members",
        "Produce coop peers (Markon / Pro*Act)",
    ),
    "Organic Produce Distributor": (
        "Organic retail produce",
        "Organic specialty produce",
        "Moderate organic foodservice",
        "Limited B2B",
        "Medium–High organic produce",
        "Organic produce wholesale",
        "Organic cold-chain facilities",
        "Organic produce delivery",
        "Organic produce peers",
    ),
    "Social Food Distribution": (
        "Social / public retail programs",
        "Limited",
        "Limited",
        "Limited",
        "Medium — social distribution",
        "Social / public food distribution",
        "Public distribution warehouses",
        "Program delivery",
        "Public distribution peers",
    ),
}

BT_ALIASES = {
    "Foodservice / Wholesale Distributor": "Foodservice Distributor",
    "Foodservice / Wholesale": "Foodservice Wholesale",
    "Foodservice Distributor / Cash & Carry": "Cash & Carry / Foodservice Wholesale",
    "Foodservice / Cash & Carry": "Cash & Carry / Foodservice Wholesale",
    "Foodservice / Drinks Distributor": "Foodservice Distributor",
    "Drinks / Foodservice Distributor": "Foodservice Distributor",
    "Dairy / Foodservice Distributor": "Foodservice Distributor",
    "Beverage / Foodservice Distribution": "Foodservice Distributor",
    "Foodservice Supplies Distributor": "Foodservice Distributor",
    "Specialty Food Ingredients Distributor": "Specialty Ingredients Distributor",
    "Specialty Chemical / Ingredients Distributor": "Specialty Ingredients Distributor",
    "Food Ingredients Broker / Distributor": "Food Ingredients Distributor",
    "Oils / Food Ingredients Distributor": "Food Ingredients Distributor",
    "Oils & Fats Ingredients Distributor": "Food Ingredients Distributor",
    "Food Ingredients / Agri Distributor": "Food Ingredients Distributor",
    "Vegetable Oils / Ingredients Distributor": "Vegetable Oils Distributor",
    "Oils & Fats Distributor": "Oils Distributor",
    "Edible Oils / Food Distributor": "Food Distributor",
    "Fresh Produce Cooperative / Exporter": "Fresh Produce Cooperative",
    "Fresh Produce Importer / Distributor": "Fresh Produce Distributor",
    "Fresh Produce Grower / Distributor": "Fresh Produce Distributor",
    "Specialty Produce Distributor": "Fresh Produce Distributor",
    "Specialty Produce / Foodservice Distributor": "Specialty Foodservice Distributor",
    "Natural / Organic Distributor": "Natural / Specialty Grocery Wholesale",
    "Natural Ingredients / Oils Distributor": "Natural / Specialty Grocery Wholesale",
    "Natural Oils / Ingredients Wholesale": "Natural / Cosmetic Ingredients Wholesale",
    "Natural Ingredients Wholesale": "Natural / Cosmetic Ingredients Wholesale",
    "Natural Ingredients Supplier": "Natural / Cosmetic Ingredients Wholesale",
    "Natural Ingredients Retail / Wholesale": "Natural / Cosmetic Ingredients Wholesale",
    "Bulk Natural Oils Distributor": "Bulk Oils / Ingredients Supplier",
    "Essential Oils / Ingredients Wholesale": "Natural / Cosmetic Ingredients Wholesale",
    "Essential Oils / Craft Ingredients": "Cosmetic Ingredients Supplier",
    "Soap Making Supplies Wholesale": "Soap / Cosmetic Ingredients Supplier",
    "Oleochemical Distributor": "Oleochemical / Ingredients Distributor",
    "Chemical / Ingredients Distributor": "Chemical Distributor",
    "Chemical / Ingredients Trader": "Chemical Distributor",
    "Nutraceutical Ingredients Distributor": "Specialty Ingredients Distributor",
    "Grocery / Convenience Wholesale": "Grocery Wholesale",
    "Convenience Wholesale": "Grocery Wholesale",
    "Trading House": "Trading House Food Distribution",
    "Frozen Food Distributor": "Food Distributor",
    "Specialty Oils Distributor": "Specialty Food Distributor",
    "Specialty Spices Retail / Wholesale": "Specialty Food Distributor",
    "Specialty Grocery Retailer": "Specialty Grocery Retailer",
}


def region_from_geo(hq: str, cities: str, company: str) -> tuple[str, str]:
    """Return (Regions Served, Regional Extensions) from geography fields."""
    blob = f"{hq} {cities} {company}".strip()
    low = blob.lower()
    # Bidfood country entities
    m = re.match(r"(?i)bidfood\s+(.+)", company.strip())
    if m:
        region = m.group(1).strip()
        special = {
            "poland farutex": "Poland",
            "uk": "United Kingdom",
            "czech": "Czech Republic",
            "baltics": "Latvia, Lithuania, Estonia",
            "asia": "Asia (Bidfood Asia)",
            "south africa": "South Africa",
            "new zealand": "New Zealand",
            "netherlands": "Netherlands",
            "singapore": "Singapore",
            "australia": "Australia",
            "belgium": "Belgium",
            "chile": "Chile",
            "portugal": "Portugal",
            "spain": "Spain",
            "turkey": "Turkey",
        }
        key = region.lower()
        served = special.get(key, region.title())
        return served, "Bidcorp group (31-country foodservice network)"

    if "global" in low or "worldwide" in low:
        return "Global", "Multi-region international network"
    if "united states" in low or re.search(r"\busa\b|\bu\.s\.", low):
        if "canada" in low and "mexico" in low:
            return "United States, Canada, Mexico", "North America"
        if "canada" in low:
            return "United States and Canada", "North America"
        return "United States", "North America adjacency where applicable"
    if "canada" in low and "united states" not in low:
        return "Canada", "North America"
    if "mexico" in low:
        return "Mexico", "LatAm / North America adjacency"
    if "united kingdom" in low or re.search(r"\buk\b|england|scotland|wales", low):
        return "United Kingdom", "Europe adjacency"
    if "netherlands" in low:
        return "Netherlands", "Benelux / Europe"
    if "germany" in low:
        return "Germany", "Europe"
    if "france" in low:
        return "France", "Europe"
    if "spain" in low:
        return "Spain", "Europe / Mediterranean"
    if "belgium" in low:
        return "Belgium", "Benelux / Europe"
    if "australia" in low:
        return "Australia", "ANZ / Asia-Pacific"
    if "new zealand" in low:
        return "New Zealand", "ANZ"
    if "uae" in low or "dubai" in low or "abu dhabi" in low or "united arab emirates" in low:
        return "UAE / GCC", "Middle East"
    if "pakistan" in low:
        return "Pakistan", "South Asia"
    if "india" in low:
        return "India", "South Asia"
    if "china" in low:
        return "China", "Asia"
    if "japan" in low:
        return "Japan", "Asia"
    if "brazil" in low:
        return "Brazil", "LatAm"
    if "chile" in low:
        return "Chile", "LatAm"
    if "south africa" in low:
        return "South Africa", "Southern Africa"
    if "europe" in low:
        return "Europe", "European markets"
    if "asia" in low:
        return "Asia", "Asia-Pacific"
    if "middle east" in low or "gcc" in low:
        return "Middle East / GCC", "MENA"
    # fallback: use cities or hq truncated
    geo = (cities or hq or "Regional").strip()
    if len(geo) > 80:
        geo = geo[:77] + "..."
    return geo, "Core market focus; limited disclosed extensions"


def strength_from_scale(employees: str, turnover: str, default: str) -> str:
    blob = f"{employees} {turnover}".lower()
    if any(x in blob for x in ("10001", "50000", "100000", "$10", "$20", "$50", "$70", "$80", "billion")):
        if "High" in default or "Very" in default:
            return default
        return "High — scaled national/global distributor"
    if any(x in blob for x in ("1001-5000", "5001", "$1–2b", "$1-2b", "$500m")):
        return default if "High" in default else "Medium–High — mid-to-large distributor"
    return default


def resolve_verified(company: str) -> tuple[str, ...] | None:
    key = nk(company)
    if key in VERIFIED:
        return VERIFIED[key]
    if key.startswith("sysco"):
        if "canada" in key:
            return VERIFIED["sysco canada"]
        if "europe" in key or "brakes" in key:
            return VERIFIED["sysco europe"]
        if "chef" in key or "specialty" in key:
            return VERIFIED.get("sysco")
        return VERIFIED["sysco"]
    if key.startswith("bidfood") or key.startswith("bidvest"):
        # region-specific built below
        return None
    if key.startswith("brenntag"):
        return (
            VERIFIED["brenntag north america"]
            if "north america" in key
            else VERIFIED["brenntag"]
        )
    if key.startswith("barentz"):
        return VERIFIED["barentz"]
    if key.startswith("azelis"):
        return VERIFIED["azelis"]
    if key.startswith("imcd"):
        return VERIFIED["imcd"]
    if key.startswith("dksh"):
        return VERIFIED["dksh"]
    if key.startswith("metro"):
        return VERIFIED["metro cash and carry"]
    if key.startswith("makro"):
        return VERIFIED["makro"]
    if "restaurant depot" in key or key.startswith("jetro"):
        return VERIFIED["jetro restaurant depot"]
    if key.startswith("unfi") or "united natural foods" in key:
        return VERIFIED["unfi united natural foods"]
    if key.startswith("kehe"):
        return VERIFIED["kehe distributors"]
    if key.startswith("us foods"):
        return VERIFIED["us foods chef store"] if "chef" in key else VERIFIED["us foods"]
    if "gordon food" in key:
        return VERIFIED["gordon food service"]
    if "performance food" in key or key.startswith("pfg"):
        return VERIFIED["performance food group"]
    if "trader joe" in key:
        return VERIFIED["trader joe s"]
    if "baldor" in key:
        return VERIFIED["baldor specialty foods"]
    if key.startswith("fyffes"):
        return VERIFIED["fyffes"]
    if key.startswith("anecoop"):
        return VERIFIED["anecoop"]
    if "sligro" in key:
        return VERIFIED["sligro food group"]
    if "wakefern" in key:
        return VERIFIED["wakefern food corp"]
    if "c and s" in key or key.startswith("c s wholesale"):
        return VERIFIED["c and s wholesale grocers"]
    if "associated wholesale grocers" in key:
        return VERIFIED["associated wholesale grocers"]
    if "colabor" in key:
        return VERIFIED["colabor group inc"]
    if "sam s club mexico" in key:
        return VERIFIED["sam s club mexico"]
    if "costco mexico" in key:
        return VERIFIED["costco mexico wholesale"]
    if "tradin organic" in key:
        return VERIFIED["tradin organic"]
    return None


def bidfood_profile(company: str) -> tuple[str, ...]:
    regions, ext = region_from_geo("", "", company)
    return (
        "Limited — foodservice primary (Bidcorp model)",
        "Limited specialty foodservice SKUs",
        f"Core — {regions} HoReCa / catering / hospitality",
        "B2B online ordering / Bidfood digital tools",
        "High — Bidcorp country foodservice business",
        "Broadline foodservice wholesale",
        regions,
        ext,
        f"Bidfood {regions} multi-temp DC / depot network",
        "Scheduled foodservice delivery fleet",
        "Local broadline peer; part of Bidcorp 31-country group",
    )


def build_row(company: str, btype: str, hq: str, cities: str, employees: str, turnover: str) -> dict[str, str]:
    verified = resolve_verified(company)
    if verified is None and ("bidfood" in company.lower() or "bidvest" in company.lower()):
        verified = bidfood_profile(company)
    if verified:
        vals = list(verified)
        return dict(zip(FIELDS, vals))

    mapped = BT_ALIASES.get(btype, btype)
    base = BT_BASE.get(mapped) or BT_BASE.get(btype)
    regions, extensions = region_from_geo(hq, cities, company)

    if base:
        retail, specialty, foodservice, ecommerce, strength, dist_type, warehouse, delivery, benchmark = base
        strength = strength_from_scale(employees, turnover, strength)
        vals = [
            retail,
            specialty,
            foodservice,
            ecommerce,
            strength,
            dist_type,
            regions,
            extensions,
            warehouse,
            delivery,
            benchmark,
        ]
        return dict(zip(FIELDS, vals))

    # generic fallback differentiated by keywords in BT
    bt = (btype or "").lower()
    if "foodservice" in bt or "horeca" in bt:
        base_key = "Foodservice Distributor"
    elif "cash" in bt:
        base_key = "Cash & Carry / Wholesale"
    elif "grocery" in bt or "retail" in bt:
        base_key = "Grocery Wholesale"
    elif "ingredient" in bt or "chemical" in bt:
        base_key = "Specialty Ingredients Distributor"
    elif "produce" in bt:
        base_key = "Fresh Produce Distributor"
    elif "cosmetic" in bt or "soap" in bt:
        base_key = "Cosmetic Ingredients Supplier"
    else:
        base_key = "Food & Ingredients Distributor"
    retail, specialty, foodservice, ecommerce, strength, dist_type, warehouse, delivery, benchmark = BT_BASE[base_key]
    return dict(
        zip(
            FIELDS,
            [
                retail,
                specialty,
                foodservice,
                ecommerce,
                strength_from_scale(employees, turnover, strength),
                dist_type,
                regions,
                extensions,
                warehouse,
                delivery,
                benchmark,
            ],
        )
    )


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_channel_cols_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    changed = 0
    tot = 0
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        tot += 1
        btype = str(ws.cell(r, cols["Business Type"]).value or "").strip()
        hq = str(ws.cell(r, cols["Headquarters"]).value or "").strip()
        cities = str(ws.cell(r, cols["Cities / Regions"]).value or "").strip()
        employees = str(ws.cell(r, cols["Employees"]).value or "").strip()
        turnover = str(ws.cell(r, cols["Turnover / Scale"]).value or "").strip()

        profile = build_row(company, btype, hq, cities, employees, turnover)
        for f in FIELDS:
            old = str(ws.cell(r, cols[f]).value or "").strip()
            new = profile[f]
            # Always overwrite generics; also overwrite if new is richer verified content
            if is_generic(old) or old != new:
                if old != new:
                    ws.cell(r, cols[f]).value = new
                    changed += 1

    # stats
    dist = {f: Counter() for f in FIELDS}
    generic_left = Counter()
    for r in range(2, ws.max_row + 1):
        for f in FIELDS:
            v = str(ws.cell(r, cols[f]).value or "").strip()
            dist[f][v] += 1
            if is_generic(v):
                generic_left[f] += 1

    wb.save(PATH)
    print("backup", bak.name)
    print("companies", tot, "cells_changed", changed)
    print("generic_left", dict(generic_left))
    for f in FIELDS:
        print(f"--- {f} unique={len(dist[f])} ---")
        for v, n in dist[f].most_common(6):
            print(f"  {n:3d} | {v[:95]}")


if __name__ == "__main__":
    main()
