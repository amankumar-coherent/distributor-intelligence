"""Normalize Ownership and Business Type for top-10 expanded distributor Excels."""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

# Canonical business types
BT_DISTRIBUTOR = "Distributor"
BT_WHOLESALER = "Wholesaler"
BT_RETAILER = "Retailer"
BT_ONLINE_RETAILER = "Online Retailer"
BT_IMPORTER_DIST = "Importer / Distributor"
BT_MFG_DIST = "Manufacturer / Distributor"
BT_MARKETPLACE = "Marketplace"
BT_AUTH_DIST = "Authorized Distributor"
BT_BRAND_NET = "Brand Distribution Network"
BT_WHOLESALE_RETAIL = "Wholesaler / Retailer"

# Exact company overrides: (Ownership, Business Type)
OVERRIDE: dict[str, tuple[str, str]] = {
    # Public / major medical
    "McKesson": ("McKesson Corporation (public)", BT_WHOLESALER),
    "McKesson Medical-Surgical": ("McKesson Corporation (public)", BT_WHOLESALER),
    "McKesson (Consumer Health)": ("McKesson Corporation (public)", BT_WHOLESALER),
    "Cardinal Health": ("Cardinal Health, Inc. (public)", BT_WHOLESALER),
    "Owens & Minor": ("Owens & Minor, Inc. (public)", BT_WHOLESALER),
    "Henry Schein": ("Henry Schein, Inc. (public)", BT_DISTRIBUTOR),
    "Medline Industries,": ("Medline Industries, LP (private)", BT_MFG_DIST),
    "Medline Industries": ("Medline Industries, LP (private)", BT_MFG_DIST),
    "Independence Medical (Cardinal specialty)": (
        "Cardinal Health, Inc. (public)",
        BT_DISTRIBUTOR,
    ),
    "Bunzl Healthcare / related MRO medical": ("Bunzl plc (public)", BT_WHOLESALER),
    "Paul Hartmann AG distribution": ("Paul Hartmann AG (public)", BT_BRAND_NET),
    "Mölnlycke Health Care distribution": ("Mölnlycke Health Care (private)", BT_BRAND_NET),
    "Amazon Medical Supplies": ("Amazon.com, Inc. (public)", BT_MARKETPLACE),
    "Walmart Mobility / OTC": ("Walmart Inc. (public)", BT_RETAILER),
    "CVS Health DME/OTC": ("CVS Health Corporation (public)", BT_RETAILER),
    "Walgreens DME/OTC": ("Walgreens Boots Alliance, Inc. (public)", BT_RETAILER),
    "3M": ("3M Company (public)", BT_MFG_DIST),
    "Essity": ("Essity AB (public)", BT_MFG_DIST),
    "BSN medical / Essity distribution": ("Essity AB (public)", BT_BRAND_NET),
    "BSN medical / clinic consumables channel": ("Essity AB (public)", BT_BRAND_NET),
    "Ottobock dealer network": ("Ottobock SE & Co. KGaA (private)", BT_BRAND_NET),
    "Össur distribution partners": ("Össur hf. (public)", BT_BRAND_NET),
    "Invacare dealer network": ("Invacare Holdings Corporation (public)", BT_BRAND_NET),
    "Sunrise Medical": ("Sunrise Medical (private)", BT_MFG_DIST),
    "Sunrise Medical dealer network": ("Sunrise Medical (private)", BT_BRAND_NET),
    "Drive DeVilbiss Healthcare distribution": (
        "Drive DeVilbiss Healthcare (private)",
        BT_BRAND_NET,
    ),
    "Pride Mobility dealer network": ("Pride Mobility Products (private)", BT_BRAND_NET),
    "AliMed": ("AliMed, Inc. (private)", BT_DISTRIBUTOR),
    "ScripHessco": ("ScripHessco (private)", BT_DISTRIBUTOR),
    "Concordance Healthcare Solutions": (
        "Concordance Healthcare Solutions (private)",
        BT_WHOLESALER,
    ),
    "Allegro Medical": ("Allegro Medical Supplies, Inc. (private)", BT_ONLINE_RETAILER),
    "Health Products For You": ("Quaisar Enterprises, LLC (private)", BT_ONLINE_RETAILER),
    "Activeforever": ("ActiveForever (private)", BT_ONLINE_RETAILER),
    "School Health Corporation": ("School Health Corporation (private)", BT_DISTRIBUTOR),
    "School Health balance/screening supplies": (
        "School Health Corporation (private)",
        BT_DISTRIBUTOR,
    ),
    "Direct Healthcare Group / UK DME": (
        "Direct Healthcare Group (private)",
        BT_MFG_DIST,
    ),
    "Aidapt / UK mobility wholesale": ("Aidapt Limited (private)", BT_WHOLESALER),
    "Days Healthcare distribution": (
        "Direct Healthcare Group / Days Healthcare (private)",
        BT_BRAND_NET,
    ),
    # Baijiu / spirits
    "JD.com": ("JD.com, Inc. (public)", BT_ONLINE_RETAILER),
    "JD.com supermarket / Jingdong": ("JD.com, Inc. (public)", BT_ONLINE_RETAILER),
    "Alibaba / Tmall supermarket": ("Alibaba Group Holding Limited (public)", BT_MARKETPLACE),
    "Young's Market Company": ("Young's Market Company (private)", BT_WHOLESALER),
    "Breakthru Beverage Group": ("Breakthru Beverage Group, LLC (private)", BT_WHOLESALER),
    "Republic National Distributing Company (RNDC)": (
        "Republic National Distributing Company (private)",
        BT_WHOLESALER,
    ),
    "Winebow Group": ("Winebow Group (private)", BT_WHOLESALER),
    "Opici Family Distributing": ("Opici Family Distributing (private / family)", BT_WHOLESALER),
    "CNS Imports": ("CNS Imports (private)", BT_IMPORTER_DIST),
    "Cnsimports": ("CNS Imports (private)", BT_IMPORTER_DIST),
    "DrinkUpNY": ("Borisal Liquor and Wine / DrinkUpNY (private)", BT_RETAILER),
    "Shiwandistillery": ("Guangdong Shiwan Wine Group Co., Ltd. (private)", BT_MFG_DIST),
    "Pernod Ricard China / Asia distribution": ("Pernod Ricard SA (public)", BT_BRAND_NET),
    "Diageo Greater China distribution": ("Diageo plc (public)", BT_BRAND_NET),
    "Remy Cointreau China network": ("Rémy Cointreau SA (public)", BT_BRAND_NET),
    "Duty Free Americas / travel retail spirits": (
        "Duty Free Americas (private)",
        BT_RETAILER,
    ),
    "DFS Group (travel retail)": ("DFS Group (LVMH) (public parent)", BT_RETAILER),
    "China Duty Free Group": ("China Tourism Group Duty Free (state-owned / public)", BT_RETAILER),
    "BevMo!": ("GoPuff / BevMo! (private)", BT_RETAILER),
    "Astor Wines & Spirits": ("Astor Wines & Spirits (private)", BT_RETAILER),
    "Summergate Fine Wines & Spirits": ("Summergate (private)", BT_IMPORTER_DIST),
    "ASC Fine Wines (China on-trade/off-trade)": ("ASC Fine Wines (private)", BT_IMPORTER_DIST),
    "Jazz Trading (London) Ltd": ("Jazz Trading (London) Ltd (private)", BT_IMPORTER_DIST),
    "MHW Ltd.": ("MHW Ltd. (private)", BT_IMPORTER_DIST),
    # Vertigo / clinical
    "AmerisourceBergen": ("Cencora, Inc. (public)", BT_WHOLESALER),
    "AmerisourceBergen / Cencora": ("Cencora, Inc. (public)", BT_WHOLESALER),
    "Thermo Fisher": ("Thermo Fisher Scientific Inc. (public)", BT_DISTRIBUTOR),
    "Avantor": ("Avantor, Inc. (public)", BT_DISTRIBUTOR),
    "Stryker": ("Stryker Corporation (public)", BT_MFG_DIST),
    "Medtronic": ("Medtronic plc (public)", BT_MFG_DIST),
    "Medtronic ENT distribution": ("Medtronic plc (public)", BT_BRAND_NET),
    "Karl Storz distribution partners": ("KARL STORZ SE & Co. KG (private)", BT_BRAND_NET),
    "Olympus Medical distribution": ("Olympus Corporation (public)", BT_BRAND_NET),
    "Natus Medical / Otometrics distribution": (
        "Natus Medical / ArchiMed (private equity)",
        BT_BRAND_NET,
    ),
    "GN Otometrics / Natus channel partners": (
        "Natus Medical / ArchiMed (private equity)",
        BT_BRAND_NET,
    ),
    "Interacoustics distributor network": ("Demant A/S (public)", BT_BRAND_NET),
    "Interacoustics USA / audiology dealers": ("Demant A/S (public)", BT_BRAND_NET),
    "Amplivox distributors": ("Amplivox Ltd (private)", BT_BRAND_NET),
    "MAICO Diagnostics distribution": ("MAICO Diagnostics GmbH (private)", BT_BRAND_NET),
    "e3 Diagnostics": ("e3 Diagnostics (private)", BT_DISTRIBUTOR),
    "Cedar Gate Technologies": ("Cedar Gate Technologies (private)", "Healthcare Technology / Channel"),
    "Allied Medical": ("Allied Medical (private)", BT_DISTRIBUTOR),
    # Elderberry / CPG
    "KeHE Distributors": ("KeHE Distributors (employee-owned)", BT_WHOLESALER),
    "UNFI (United Natural Foods)": ("United Natural Foods, Inc. (public)", BT_WHOLESALER),
    "C&S Wholesale Grocers": ("C&S Wholesale Grocers (private)", BT_WHOLESALER),
    "Costco": ("Costco Wholesale Corporation (public)", BT_RETAILER),
    "Costco Wholesale": ("Costco Wholesale Corporation (public)", BT_RETAILER),
    "Walmart": ("Walmart Inc. (public)", BT_RETAILER),
    "CVS Health": ("CVS Health Corporation (public)", BT_RETAILER),
    "Walgreens": ("Walgreens Boots Alliance, Inc. (public)", BT_RETAILER),
    "Amazon": ("Amazon.com, Inc. (public)", BT_MARKETPLACE),
    "iHerb Wholesale / B2B channel": ("iHerb, LLC (private)", BT_ONLINE_RETAILER),
    "The Vitamin Shoppe Distribution": ("Vitamin Shoppe (private)", BT_RETAILER),
    "Vitacost": ("The Kroger Co. (public)", BT_ONLINE_RETAILER),
    "Vitacost (Kroger)": ("The Kroger Co. (public)", BT_ONLINE_RETAILER),
    "Garden of Life": ("Nestlé Health Science (Nestlé S.A.) (public)", BT_MFG_DIST),
    "NOW Foods": ("NOW Health Group (private)", BT_MFG_DIST),
    "Herb Pharm": ("Herb Pharm (private)", BT_MFG_DIST),
    "Herbalife": ("Herbalife Ltd. (public)", BT_MFG_DIST),
    "Sundown Naturals": ("Nestlé Health Science / Nature's Bounty (public parent)", BT_MFG_DIST),
    "Sundown Naturals Distribution": (
        "Nestlé Health Science / Nature's Bounty (public parent)",
        BT_BRAND_NET,
    ),
    "Elder Farms": ("Elder Farm LLC (private / family farm)", BT_MFG_DIST),
    "Heartland Elderberry Farms": ("Heartland Elderberry Farms (private / family)", BT_MFG_DIST),
    "Emerson Ecologics / Fullscript": ("Fullscript (private)", BT_DISTRIBUTOR),
    "Natural Partners": ("Natural Partners (private)", BT_DISTRIBUTOR),
    "Wellevate / Prevention Wellness Dist.": ("Fullscript / Wellevate (private)", BT_DISTRIBUTOR),
    # Instant noodles / grocery
    "Gordon Food Service": ("Gordon Food Service (private)", BT_WHOLESALER),
    "McLane Company": ("Berkshire Hathaway Inc. (public) / McLane", BT_WHOLESALER),
    "Performance Food Group": ("Performance Food Group Company (public)", BT_WHOLESALER),
    "Performance Food Group (PFG)": ("Performance Food Group Company (public)", BT_WHOLESALER),
    "Sysco": ("Sysco Corporation (public)", BT_WHOLESALER),
    "Aeon Group wholesale/retail": ("Aeon Co., Ltd. (public)", BT_WHOLESALE_RETAIL),
    "Nestlé Maggi distribution": ("Nestlé S.A. (public)", BT_BRAND_NET),
    "Hethstia": ("Sichuan Sentaiyuan Biotechnology Co., Ltd. (private)", BT_MFG_DIST),
    "Rakuten / Japanese e-grocery channels": ("Rakuten Group, Inc. (public)", BT_MARKETPLACE),
    # MTB
    "Chain Reaction Cycles": ("Frasers Group plc (public) / WiggleCRC legacy", BT_ONLINE_RETAILER),
    "Wiggle": ("Frasers Group plc (public)", BT_ONLINE_RETAILER),
    "Bilstein": ("thyssenkrupp AG (public) / Bilstein", BT_MFG_DIST),
    "Koni": ("ICG / KONI (private)", BT_MFG_DIST),
    "Pon.Bike distribution (Cannondale, Santa Cruz parent ops)": (
        "Pon Holdings B.V. (private)",
        BT_BRAND_NET,
    ),
    "Shimano distributor network": ("Shimano Inc. (public)", BT_BRAND_NET),
    "Bicycle Warehouse": ("Bicycle Warehouse (private)", BT_RETAILER),
    "Performance Bicycle": ("Performance Inc. legacy (private / closed retail)", BT_RETAILER),
    "Universal Cycles": ("Universal Cycles (private)", BT_WHOLESALER),
    "Worldwide Cyclery": ("Worldwide Cyclery (private)", BT_ONLINE_RETAILER),
    "BikeDiscount / Mantel": ("Bike-Discount / Mantel (private)", BT_ONLINE_RETAILER),
    "99 Bikes (AU retail chain)": ("99 Bikes (private)", BT_RETAILER),
    "BikesOnline / DTC AU-NZ": ("BikesOnline (private)", BT_ONLINE_RETAILER),
    "DVO Suspension distributors": ("DVO Suspension (private)", BT_BRAND_NET),
    "KHS Distribution / KHS Bicycles": ("KHS Bicycles (private)", BT_DISTRIBUTOR),
    # Muslin / fabric
    "JOANN (Fabric & Craft)": ("JOANN Inc. (private equity / restructured)", BT_RETAILER),
    "Joann Stores": ("JOANN Inc. (private equity / restructured)", BT_RETAILER),
    "Hobby Lobby Fabric": ("Hobby Lobby Stores, Inc. (private)", BT_RETAILER),
    "Fabric.com (Amazon)": ("Amazon.com, Inc. (public)", BT_ONLINE_RETAILER),
    "Fabric.com": ("Amazon.com, Inc. (public)", BT_ONLINE_RETAILER),
    "Mood Fabrics": ("Mood Fabrics (private)", BT_WHOLESALE_RETAIL),
    "Sewing Machine Plus": ("SewingMachinesPlus.com (private / family)", BT_RETAILER),
    "SyFabrics": ("Sytradingpost Inc. (private)", BT_WHOLESALER),
    "Casa Fashions (Karur)": ("Casa Fashions (private)", BT_MFG_DIST),
    "Central India Fabrics": ("Central India Fabrics (private / proprietorship)", BT_MFG_DIST),
    "ROC-LON Fabrics (Rockland Industries)": (
        "Rockland Industries (private)",
        BT_MFG_DIST,
    ),
    "Whaleys (Bradford) Ltd": ("Whaleys (Bradford) Ltd (private)", BT_WHOLESALER),
    "Liberty Fabrics (Wholesale)": ("Liberty London (private equity history)", BT_WHOLESALER),
    "Coats Group Trade Distribution": ("Coats Group plc (public)", BT_BRAND_NET),
    "Prym Consumer USA / Haberdashery Dist.": ("Prym Group (private)", BT_BRAND_NET),
    "Gutermann Distribution Partners": ("Gütermann / A&E Gütermann (private)", BT_BRAND_NET),
    "Spotlight (Australia/NZ)": ("Spotlight Group (private)", BT_RETAILER),
    "Rose Brand": ("Rose Brand (private)", BT_WHOLESALER),
    "OnlineFabricStore": ("OnlineFabricStore (private)", BT_ONLINE_RETAILER),
    "DM Distributors": ("D&M Distributors LLC (private)", BT_WHOLESALER),
    "Doray Kumas": ("Doray Kumas (private)", BT_WHOLESALER),
    "Fabriclore": ("Fabriclore (private)", BT_WHOLESALER),
    "Global Sources Fabrics": ("Global Sources (private / public history)", BT_MARKETPLACE),
    # Nail
    "Sally Beauty": ("Sally Beauty Holdings, Inc. (public)", BT_RETAILER),
    "Armstrong McCall": ("Sally Beauty Holdings, Inc. (public)", BT_WHOLESALER),
    "Beauty Joint": ("Beauty Joint (private)", BT_DISTRIBUTOR),
    "Dearsundays": ("Dear Sundays Inc. (private)", BT_RETAILER),
    "CND (Creative Nail Design) distribution": (
        "Revlon / CND (public parent history)",
        BT_BRAND_NET,
    ),
    "Gelish / Hand & Nail Harmony distributors": (
        "Hand & Nail Harmony (private)",
        BT_BRAND_NET,
    ),
    "Kiara Sky distribution": ("Kiara Sky Professional Nails (private)", BT_BRAND_NET),
    "Young Nails distribution": ("Young Nails (private)", BT_BRAND_NET),
    "Nail Systems International (NSI) distribution": (
        "Nail Systems International (private)",
        BT_BRAND_NET,
    ),
    "Light Elegance distribution": ("Light Elegance (private)", BT_BRAND_NET),
    "Burmax Beauty Distributors": ("Burmax Company (private)", BT_WHOLESALER),
    "Marianna Industries distribution": ("Marianna Industries (private)", BT_WHOLESALER),
    "Beauty Manufacturing Solutions (BMS Nail)": (
        "Beauty Manufacturing Solutions Corp (private)",
        BT_MFG_DIST,
    ),
    "The Nailhub / pro e-wholesalers": ("The Nail Hub (private)", BT_ONLINE_RETAILER),
    "Kaganails": ("M8 jor International Limited (private)", BT_IMPORTER_DIST),
    "Wholesale Nail Supply": ("Wholesale Nail Supply (private)", BT_WHOLESALER),
    "Douglas (Europe)": ("Douglas Group (public / private equity)", BT_RETAILER),
    "Douglas Beauty Europe": ("Douglas Group (public / private equity)", BT_RETAILER),
    "Beauty Bay": ("THG plc / Beauty Bay (public parent history)", BT_ONLINE_RETAILER),
    "Beauty Bay sheet masks": ("THG plc / Beauty Bay (public parent history)", BT_ONLINE_RETAILER),
    "Lookfantastic / THG Beauty": ("THG plc (public)", BT_ONLINE_RETAILER),
    "Cult Beauty": ("THG plc / Cult Beauty (public parent)", BT_ONLINE_RETAILER),
    "Notino": ("Notino (private)", BT_ONLINE_RETAILER),
    "Chemistwarehouse": ("Chemist Warehouse (private)", BT_RETAILER),
    "Jessicacosmetics": ("Jessica Cosmetics (private)", BT_MFG_DIST),
    "Essie / L'Oréal distribution": ("L'Oréal S.A. (public)", BT_BRAND_NET),
    "Born Pretty wholesale": ("Born Pretty (private)", BT_ONLINE_RETAILER),
    "Makartt wholesale": ("Makartt (private)", BT_ONLINE_RETAILER),
    "Modelones distribution": ("Modelones (private)", BT_ONLINE_RETAILER),
    # Sheet masks / beauty
    "Innisfree": ("Amorepacific Corporation (public)", BT_RETAILER),
    "Innisfree distribution (Amorepacific)": ("Amorepacific Corporation (public)", BT_BRAND_NET),
    "Laneige / Amorepacific distribution": ("Amorepacific Corporation (public)", BT_BRAND_NET),
    "Dr. Jart+ distribution (Estée Lauder)": (
        "The Estée Lauder Companies Inc. (public)",
        BT_BRAND_NET,
    ),
    "Tony Moly distribution": ("Tony Moly (private)", BT_BRAND_NET),
    "Doctor Althea Global": ("Dr. Althea (private)", BT_ONLINE_RETAILER),
    "Peach & Lily": ("Peach & Lily (private)", BT_ONLINE_RETAILER),
    "Peach & Lily distribution": ("Peach & Lily (private)", BT_ONLINE_RETAILER),
    "Soko Glam": ("Soko Glam (private)", BT_ONLINE_RETAILER),
    "Sokoglam": ("Soko Glam (private)", BT_ONLINE_RETAILER),
    "Soko Glam wholesale partners": ("Soko Glam (private)", BT_ONLINE_RETAILER),
    "YesStyle": ("YesStyle.com Ltd (private)", BT_ONLINE_RETAILER),
    "YesStyle Wholesale inquiries / B2B": ("YesStyle.com Ltd (private)", BT_ONLINE_RETAILER),
    "Beautylish": ("Beautylish (private)", BT_ONLINE_RETAILER),
    "Beauty Bridge": ("Beauty Bridge (private)", BT_ONLINE_RETAILER),
    "SkinStore": ("THG plc / SkinStore (public parent history)", BT_ONLINE_RETAILER),
    "Sephora K-beauty assortment": ("LVMH / Sephora (public)", BT_RETAILER),
    "Watsons / AS Watson beauty retail": (
        "CK Hutchison / AS Watson (public parent)",
        BT_RETAILER,
    ),
    "Guardian / AS Watson": ("CK Hutchison / AS Watson (public parent)", BT_RETAILER),
    "Boots UK beauty": ("Walgreens Boots Alliance, Inc. (public)", BT_RETAILER),
    "Superdrug": ("AS Watson / CK Hutchison (public parent)", BT_RETAILER),
    "Jumia": ("Jumia Technologies AG (public)", BT_MARKETPLACE),
    "Anand Agencies (Chennai)": ("Anand Agencies (private)", BT_DISTRIBUTOR),
    # Shunt / electronics
    "Avnet": ("Avnet, Inc. (public)", BT_DISTRIBUTOR),
    "Digi-Key Electronics": ("Digi-Key Corporation (private)", BT_DISTRIBUTOR),
    "Mouser Electronics": ("Mouser Electronics / TTI Inc. (private)", BT_DISTRIBUTOR),
    "Future Electronics": ("Future Electronics (private)", BT_DISTRIBUTOR),
    "RS Components": ("RS Group plc (public)", BT_DISTRIBUTOR),
    "RS Components / RS Group": ("RS Group plc (public)", BT_DISTRIBUTOR),
    "TE Connectivity": ("TE Connectivity Ltd. (public)", BT_MFG_DIST),
    "Sager Electronics": ("Hisco / Sager Electronics (private)", BT_DISTRIBUTOR),
    "Master Electronics": ("Master Electronics (private)", BT_DISTRIBUTOR),
    "Onlinecomponents.com": ("Master Electronics (private)", BT_ONLINE_RETAILER),
    "Distrelec": ("Distrelec Group (private)", BT_DISTRIBUTOR),
    "Chip One Stop": ("Chip One Stop (private)", BT_DISTRIBUTOR),
    "JLCPCB Parts / JiaLiChuang": ("JLCPCB / JiaLiChuang (private)", BT_ONLINE_RETAILER),
    "Prosemic": (
        "Suzhou Prosemi Micro-electronic Technology Co., Ltd. (private)",
        BT_MFG_DIST,
    ),
    "Findchips": ("Supply Frame / Digi-Key ecosystem (private)", BT_MARKETPLACE),
    "Quest Components": ("Quest Components (private)", BT_DISTRIBUTOR),
    "Classic Components": ("Classic Components Corp. (private)", BT_DISTRIBUTOR),
    "Sourceability / Sourcengine": ("Sourceability (private)", BT_DISTRIBUTOR),
    "Source": ("Top Source / sourcing platform (private)", BT_MARKETPLACE),
    "Electropeak": ("Electropeak (private)", BT_ONLINE_RETAILER),
}


def normalize_business_type(raw: str, company: str) -> str:
    s = (raw or "").strip()
    low = s.lower()
    c = company.lower()

    if company in OVERRIDE:
        return OVERRIDE[company][1]

    # Name heuristics
    if any(x in c for x in ("amazon", "alibaba", "tmall", "jumia", "findchips", "thomasnet")):
        return BT_MARKETPLACE
    if any(x in c for x in ("dealer network", "distribution partners", "distributor network", "brand distribution")):
        return BT_BRAND_NET
    if "marketplace" in low or "aggregator" in low:
        return BT_MARKETPLACE
    if "authorized" in low:
        return BT_AUTH_DIST
    if "manufacturer" in low and "distributor" in low:
        return BT_MFG_DIST
    if "importer" in low:
        return BT_IMPORTER_DIST
    if "online retailer" in low or ("online" in low and "retail" in low):
        return BT_ONLINE_RETAILER
    if "wholesaler / retailer" in low or "retailer / wholesaler" in low:
        return BT_WHOLESALE_RETAIL
    if "retailer / distributor" in low or "distributor / retailer" in low:
        return BT_WHOLESALE_RETAIL
    if low in {"retailer", "retailer / importer channel", "retailer / distributor channel"}:
        return BT_RETAILER
    if "online" in c and any(x in c for x in ("shop", "store", "bay", "style", "glam")):
        return BT_ONLINE_RETAILER
    if low in {"var"}:
        return BT_DISTRIBUTOR
    if "wholesale" in low and "retail" not in low:
        return BT_WHOLESALER
    if low in {"distributor", "distributor / wholesaler", "reseller"}:
        # Prefer Distributor for channel partners; Wholesaler if name says wholesale
        if "wholesale" in c:
            return BT_WHOLESALER
        return BT_DISTRIBUTOR
    if "wholesaler" in low:
        return BT_WHOLESALER
    if "retailer" in low:
        return BT_RETAILER
    if "distributor" in low:
        return BT_DISTRIBUTOR
    if "gpo" in low or "professional network" in low:
        return BT_DISTRIBUTOR
    if "ingredient" in low or "contract manufacturer" in low:
        return BT_MFG_DIST
    if "dtc" in low:
        return BT_ONLINE_RETAILER
    if "b2b marketplace" in low:
        return BT_MARKETPLACE

    # Default for this dataset
    return BT_DISTRIBUTOR if s else BT_DISTRIBUTOR


def normalize_ownership(raw: str, company: str, business_type: str) -> str:
    if company in OVERRIDE:
        return OVERRIDE[company][0]

    s = (raw or "").strip()
    low = s.lower()
    c = company.lower()

    # Already good public/parent labels
    if s and low not in {
        "independent",
        "private",
        "independent / privately held (public filings not always available)",
        "not confirmed publicly",
        "n/a",
        "unknown",
    }:
        # Clean trailing noise
        if "(public" in low or "(private" in low or "plc" in low or "inc." in low or "ltd" in low:
            return s
        # Known parent-only values already set
        if any(
            x in s
            for x in (
                "Corporation",
                "Inc.",
                "plc",
                "Ltd",
                "AG",
                "SA",
                "Group",
                "Holdings",
                "Amazon",
                "Walmart",
                "Nestlé",
                "LVMH",
                "Amorepacific",
                "Sally Beauty",
                "Alibaba",
                "McKesson",
                "Cardinal",
                "Avnet",
            )
        ):
            if "(public)" not in low and "(private)" not in low and "public" not in low:
                # leave as-is if already a company name
                return s
            return s

    # Heuristics for public parents from name
    public_hints = {
        "walmart": "Walmart Inc. (public)",
        "amazon": "Amazon.com, Inc. (public)",
        "cvs": "CVS Health Corporation (public)",
        "walgreens": "Walgreens Boots Alliance, Inc. (public)",
        "costco": "Costco Wholesale Corporation (public)",
        "target": "Target Corporation (public)",
        "mckesson": "McKesson Corporation (public)",
        "cardinal": "Cardinal Health, Inc. (public)",
        "henry schein": "Henry Schein, Inc. (public)",
        "nestlé": "Nestlé S.A. (public)",
        "nestle": "Nestlé S.A. (public)",
        "alibaba": "Alibaba Group Holding Limited (public)",
        "jd.com": "JD.com, Inc. (public)",
        "sephora": "LVMH / Sephora (public)",
        "l'oréal": "L'Oréal S.A. (public)",
        "loreal": "L'Oréal S.A. (public)",
        "amorepacific": "Amorepacific Corporation (public)",
        "sally beauty": "Sally Beauty Holdings, Inc. (public)",
        "thg": "THG plc (public)",
        "boots": "Walgreens Boots Alliance, Inc. (public)",
        "watsons": "CK Hutchison / AS Watson (public parent)",
        "guardian": "CK Hutchison / AS Watson (public parent)",
        "douglas": "Douglas Group (public / private equity)",
        "rs components": "RS Group plc (public)",
        "avnet": "Avnet, Inc. (public)",
        "te connectivity": "TE Connectivity Ltd. (public)",
        "digi-key": "Digi-Key Corporation (private)",
        "mouser": "Mouser Electronics / TTI Inc. (private)",
        "sysco": "Sysco Corporation (public)",
        "unfi": "United Natural Foods, Inc. (public)",
        "performance food": "Performance Food Group Company (public)",
        "mclane": "Berkshire Hathaway Inc. (public) / McLane",
        "shimano": "Shimano Inc. (public)",
        "coats": "Coats Group plc (public)",
        "essity": "Essity AB (public)",
        "3m": "3M Company (public)",
        "joann": "JOANN Inc. (private equity / restructured)",
        "hobby lobby": "Hobby Lobby Stores, Inc. (private)",
        "frasers": "Frasers Group plc (public)",
        "wiggle": "Frasers Group plc (public)",
        "pon.": "Pon Holdings B.V. (private)",
        "aeon": "Aeon Co., Ltd. (public)",
        "rakuten": "Rakuten Group, Inc. (public)",
        "cencora": "Cencora, Inc. (public)",
        "amerisource": "Cencora, Inc. (public)",
        "medtronic": "Medtronic plc (public)",
        "stryker": "Stryker Corporation (public)",
        "thermo fisher": "Thermo Fisher Scientific Inc. (public)",
        "avantor": "Avantor, Inc. (public)",
        "herbalife": "Herbalife Ltd. (public)",
        "diageo": "Diageo plc (public)",
        "pernod": "Pernod Ricard SA (public)",
        "moutai": "Kweichow Moutai Co., Ltd. (public)",
        "wuliangye": "Wuliangye Yibin Co., Ltd. (public)",
        "cofco": "COFCO (state-owned)",
        "bunzl": "Bunzl plc (public)",
        "owens": "Owens & Minor, Inc. (public)",
        "jumia": "Jumia Technologies AG (public)",
        "estée lauder": "The Estée Lauder Companies Inc. (public)",
        "estee lauder": "The Estée Lauder Companies Inc. (public)",
        "lvmh": "LVMH (public)",
    }
    for key, val in public_hints.items():
        if key in c:
            return val

    # Brand distribution networks owned by brand parent if present in ownership already
    if "distribution" in c or "dealer network" in c or "distributor network" in c:
        brand = re.split(r"\s+(distribution|dealer|distributor|partners|network|/)", company, maxsplit=1)[0].strip()
        if brand and len(brand) > 2:
            return f"{brand} (brand parent; channel partners vary)"

    return "Privately held"


def process(path: Path) -> dict:
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
    oi = header.index("Ownership")
    bi = header.index("Business Type")

    changed = 0
    for r in rows[hdr_i + 1 :]:
        company = str(r[ci].value or "").strip()
        if not company:
            continue
        old_o = str(r[oi].value or "").strip()
        old_b = str(r[bi].value or "").strip()
        new_b = normalize_business_type(old_b, company)
        new_o = normalize_ownership(old_o, company, new_b)
        if new_o != old_o or new_b != old_b:
            r[oi].value = new_o
            r[bi].value = new_b
            changed += 1

    if "Run Summary" in wb.sheetnames:
        wb["Run Summary"]["A14"] = f"Ownership/Business Type normalized: {changed} rows"

    wb.save(path)
    wb.close()
    return {"changed": changed}


def main() -> None:
    from collections import Counter

    total = 0
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        st = process(path)
        total += st["changed"]
        print(f"{path.parent.name}: updated {st['changed']}")

    own_c: Counter[str] = Counter()
    biz_c: Counter[str] = Counter()
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        data = list(ws.iter_rows(values_only=True))
        hdr_i = next(i for i, r in enumerate(data[:12]) if r and str(r[0] or "").strip() == "Company")
        header = [str(h or "").strip() for h in data[hdr_i]]
        oi, bi = header.index("Ownership"), header.index("Business Type")
        for r in data[hdr_i + 1 :]:
            if not r or not r[0]:
                continue
            own_c[str(r[oi] or "").strip()] += 1
            biz_c[str(r[bi] or "").strip()] += 1
        wb.close()

    print(f"\nTOTAL updated {total}")
    print("\nBusiness Type distribution:")
    for k, v in biz_c.most_common():
        print(f"  {v:4}  {k}")
    print("\nOwnership sample (top 25):")
    for k, v in own_c.most_common(25):
        print(f"  {v:4}  {k[:80]}")


if __name__ == "__main__":
    main()
