"""Fill/verify Founded, Headquarters, Cities/Regions on top-10 expanded Excels.

Sources: company contact/About pages, Wikipedia, IR releases, prior web verification
(e.g. AliMed Dedham; Digi-Key Thief River Falls; McKesson Irving; Henry Schein Melville;
Sally Beauty Plano; QBP Bloomington). LinkedIn MCP was not available for this pass.
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

WEAK_HQ = EMPTY | {
    "usa",
    "u.s.",
    "u.s.a.",
    "united states",
    "united kingdom",
    "uk",
    "germany",
    "china",
    "europe",
    "south korea",
    "korea",
    "hong kong",
    "japan",
    "taiwan",
    "india",
    "australia",
    "france",
    "italy",
    "sweden",
    "netherlands",
    "global",
    "usa / global",
    "europe; global",
    "asia",
    "international online",
    "multi-country",
    "california, usa",
    "texas, usa",
    "new york, usa",
}


def norm(s: str) -> str:
    s = (s or "").lower().strip()
    s = s.replace("’", "'").replace("–", "-")
    return re.sub(r"\s+", " ", s)


def is_empty(v) -> bool:
    return v is None or str(v).strip().lower() in EMPTY


def is_weak_hq(v) -> bool:
    if is_empty(v):
        return True
    s = str(v).strip().lower()
    if s in WEAK_HQ:
        return True
    # bare country / region only
    if len(s) < 12 and "," not in s and s in WEAK_HQ:
        return True
    return False


def is_weak_cities(v) -> bool:
    if is_empty(v):
        return True
    s = str(v).strip().lower()
    return s in {
        "usa",
        "global",
        "global / multi-region (confirm coverage)",
        "usa nationwide",
        "usa online",
        "global online",
        "europe",
        "asia",
    } or s.startswith("global / multi-region")


# key (lowercase substring, longest match) -> (founded|None, hq, cities)
# founded None = do not overwrite founded
GEO: dict[str, tuple[str | None, str, str]] = {
    # Medical mega
    "henry schein": (None, "135 Duryea Road, Melville, NY 11747, USA", "Melville, New York, USA; 34 countries"),
    "mckesson": (None, "6555 State Highway 161, Irving, TX 75039, USA", "Irving, Texas; Richmond, Virginia; Columbus, Ohio; Canada"),
    "cardinal health": (None, "7000 Cardinal Place, Dublin, OH 43017, USA", "Dublin, Ohio, USA; international operations"),
    "medline": (None, "Three Lakes Drive, Northfield, IL 60093, USA", "Northfield, Illinois, USA; global"),
    "owens & minor": (None, "Mechanicsville, Virginia, USA", "Mechanicsville, Virginia, USA; Europe; Asia Pacific"),
    "amerisourcebergen": (None, "1 West First Avenue, Conshohocken, PA 19428, USA", "Conshohocken, Pennsylvania, USA; global"),
    "cencora": (None, "1 West First Avenue, Conshohocken, PA 19428, USA", "Conshohocken, Pennsylvania, USA; global"),
    "alimed": (None, "297 High Street, Dedham, MA 02026, USA", "Dedham, Massachusetts, USA"),
    "sunrise medical": (None, "Fresno, California, USA", "Fresno, California, USA; Europe; international"),
    "invacare": (None, "Elyria, Ohio, USA", "Elyria, Ohio, USA; Europe; Asia"),
    "drive devilbiss": (None, "Port Washington, NY, USA", "Port Washington, New York, USA; global"),
    "graham-field": (None, "Atlanta, Georgia, USA", "Atlanta, Georgia, USA"),
    "ottobock": (None, "Duderstadt, Germany", "Duderstadt, Germany; global"),
    "össur": (None, "Reykjavik, Iceland", "Reykjavik, Iceland; global"),
    "ossur": (None, "Reykjavik, Iceland", "Reykjavik, Iceland; global"),
    "bound tree": (None, "Dublin, Ohio, USA", "Dublin, Ohio, USA"),
    "performance health": (None, "Warrenville, Illinois, USA", "Warrenville, Illinois, USA; international"),
    "school health": (None, "Chicago / Midwest, Illinois, USA", "Illinois, USA; national school channel"),
    "rehabmart": (None, "Georgia, USA", "Georgia, USA; online national"),
    "allegro medical": (None, "USA", "USA online national"),
    "vitality medical": (None, "Salt Lake City area, Utah, USA", "Utah, USA; online national"),
    "pride mobility": (None, "Exeter, Pennsylvania, USA", "Exeter, Pennsylvania, USA; international dealers"),
    "bunzl": (None, "London, United Kingdom", "London, UK; global distribution"),
    "paul hartmann": (None, "Heidenheim, Germany", "Heidenheim, Germany; Europe; global"),
    "mölnlycke": (None, "Gothenburg, Sweden", "Gothenburg, Sweden; global"),
    "molnlycke": (None, "Gothenburg, Sweden", "Gothenburg, Sweden; global"),
    "3m medical": (None, "St. Paul, Minnesota, USA", "St. Paul, Minnesota, USA; global"),
    "scriphessco": (None, "West Chester, Ohio, USA", "West Chester, Ohio, USA"),
    "health products for you": (None, "USA", "USA online; multi-warehouse"),
    "activeforever": (None, "Scottsdale, Arizona, USA", "Scottsdale, Arizona, USA"),
    "dme supply usa": (None, "USA", "USA"),
    "hugo mobility": (None, "USA", "USA retail / DME dealers"),
    "ndc (national distribution": (None, "Nashville, Tennessee, USA", "Nashville, Tennessee, USA"),
    "essity": (None, "Stockholm, Sweden", "Stockholm, Sweden; global"),
    "bsn medical": (None, "Hamburg / Charlotte (Essity BSN)", "Europe; North America"),
    "direct healthcare group": (None, "United Kingdom", "UK; Europe"),
    "aidapt": (None, "United Kingdom", "United Kingdom"),
    "days healthcare": (None, "United Kingdom", "UK; export markets"),
    "cascade healthcare": (None, "USA", "USA"),
    "concordance healthcare": (None, "USA", "USA"),
    "independence medical": (None, "USA", "USA specialty medical"),
    "emergency medical products": (None, "Waukesha, Wisconsin, USA", "Waukesha, Wisconsin, USA"),
    "allied medical": (None, "Australia", "Australia; New Zealand"),
    "nova medical": (None, "Carson, California, USA", "Carson, California, USA"),
    "promed products": (None, "USA", "USA"),
    # Vestibular / audiology
    "interacoustics": (None, "Middelfart, Denmark", "Denmark; global clinical dealers"),
    "natus": (None, "Middleton, Wisconsin, USA", "Middleton, Wisconsin, USA; global"),
    "otometrics": (None, "Taastrup, Denmark / Natus", "Denmark; USA; global"),
    "inventis": (None, "Padova, Italy", "Italy; Europe; global"),
    "micromedical": (None, "Chatham, Illinois, USA", "Chatham, Illinois, USA; export"),
    "neuro kinetics": (None, "Pittsburgh, Pennsylvania, USA", "Pittsburgh, Pennsylvania, USA"),
    "amplivox": (None, "United Kingdom", "UK; export"),
    "maico": (None, "Berlin, Germany", "Berlin, Germany; global"),
    "grason-stadler": (None, "Eden Prairie, Minnesota, USA", "Eden Prairie, Minnesota, USA; global"),
    "e3 diagnostics": (None, "USA", "USA audiology dealers"),
    "oaktree products": (None, "St. Louis area, Missouri, USA", "Missouri, USA"),
    "medrx": (None, "Florida / USA", "USA; export"),
    "synapsys": (None, "France", "France; Europe"),
    "karl storz": (None, "Tuttlingen, Germany", "Tuttlingen, Germany; global"),
    "olympus medical": (None, "Tokyo, Japan", "Tokyo, Japan; global"),
    "medtronic": (None, "Dublin, Ireland / Minneapolis, USA ops", "Global (Ireland HQ; major US ops)"),
    "stryker": (None, "Kalamazoo, Michigan, USA", "Kalamazoo, Michigan, USA; global"),
    "thermo fisher": (None, "Waltham, Massachusetts, USA", "Waltham, Massachusetts, USA; global"),
    "avantor": (None, "Radnor, Pennsylvania, USA", "Radnor, Pennsylvania, USA; global"),
    "vwr /": (None, "Radnor, Pennsylvania, USA (Avantor)", "USA; global"),
    "cole-parmer": (None, "Vernon Hills, Illinois, USA", "Vernon Hills, Illinois, USA; global"),
    "cedar gate": (None, "USA", "USA health-tech"),
    "vestibular first": (None, "USA", "USA clinical"),
    "auditdata": (None, "Denmark", "Denmark; global hearing clinics"),
    "westone": (None, "Colorado Springs, Colorado, USA", "Colorado Springs, Colorado, USA; export"),
    "hal-hen": (None, "Garden City Park, New York, USA", "New York, USA"),
    "path medical": (None, "Germany", "Germany; Europe; global"),
    "homoth": (None, "Germany", "Germany; Europe"),
    "audiology systems": (None, "USA", "USA"),
    # Electronics
    "digi-key": (None, "701 Brooks Avenue South, Thief River Falls, MN 56701, USA", "Thief River Falls, Minnesota, USA; global shipping"),
    "mouser": (None, "1000 North Main Street, Mansfield, TX 76063, USA", "Mansfield, Texas, USA; global"),
    "arrow electronics": (None, "9201 E Dry Creek Road, Centennial, CO 80112, USA", "Centennial, Colorado, USA; global"),
    "avnet": (None, "2211 South 47th Street, Phoenix, AZ 85034, USA", "Phoenix, Arizona, USA; global"),
    "future electronics": (None, "Montreal, Quebec, Canada", "Montreal, Canada; global"),
    "tti, inc": (None, "Fort Worth, Texas, USA", "Fort Worth, Texas, USA; global"),
    "tti": (None, "Fort Worth, Texas, USA", "Fort Worth, Texas, USA; global"),
    "rs components": (None, "London / Corby, United Kingdom (RS Group)", "UK; Europe; Asia; Americas"),
    "rs group": (None, "London, United Kingdom", "UK; global"),
    "farnell": (None, "Leeds, United Kingdom", "UK; Europe; Asia"),
    "element14": (None, "Leeds, United Kingdom", "UK; Asia; global"),
    "newark": (None, "Chicago, Illinois, USA", "Chicago, Illinois, USA; Americas"),
    "rochester electronics": (None, "Newburyport, Massachusetts, USA", "Newburyport, Massachusetts, USA; global"),
    "heilind": (None, "Wilmington, Massachusetts, USA", "Wilmington, Massachusetts, USA; Americas; Europe; Asia"),
    "sager electronics": (None, "Middleborough, Massachusetts, USA", "Middleborough, Massachusetts, USA; Americas"),
    "master electronics": (None, "Santa Monica / USA", "USA; Americas"),
    "lcsc": (None, "Shenzhen, China", "Shenzhen, China; global online"),
    "jlcpcb": (None, "Shenzhen / Guangdong, China", "China; global"),
    "chip one stop": (None, "Yokohama / Japan", "Japan; Asia; global"),
    "wpg holdings": (None, "Taipei, Taiwan", "Taipei, Taiwan; Asia-Pacific"),
    "wt microelectronics": (None, "Taipei, Taiwan", "Taiwan; Asia"),
    "ebv elektronik": (None, "Poing, Germany", "Germany; EMEA"),
    "silica": (None, "Munich area, Germany", "EMEA"),
    "rutronik": (None, "Ispringen, Germany", "Germany; Europe; global"),
    "schukat": (None, "Monheim am Rhein, Germany", "Germany; Europe"),
    "conrad electronic": (None, "Hirschau, Germany", "Germany; Europe"),
    "tme (transfer": (None, "Łódź, Poland", "Poland; Europe; global shipping"),
    "distrelec": (None, "Nänikon / Switzerland / Europe", "Europe"),
    "allied electronics": (None, "Fort Worth, Texas, USA", "Fort Worth, Texas, USA; Americas"),
    "automationdirect": (None, "Cumming, Georgia, USA", "Cumming, Georgia, USA; Americas"),
    "mcmaster-carr": (None, "Elmhurst, Illinois, USA", "Elmhurst, Illinois, USA; USA national"),
    "grainger": (None, "Lake Forest, Illinois, USA", "Lake Forest, Illinois, USA; USA; Canada; MRO global"),
    "hisco": (None, "Houston, Texas, USA", "Houston, Texas, USA; Americas"),
    "pei-genesis": (None, "Southampton, Pennsylvania, USA", "Pennsylvania, USA; global"),
    "bisco industries": (None, "Anaheim, California, USA", "Anaheim, California, USA; Americas"),
    "te connectivity": (None, "Schaffhausen, Switzerland / Berwyn, PA ops", "Switzerland HQ; major USA/global ops"),
    "vishay": (None, "Malvern, Pennsylvania, USA", "Malvern, Pennsylvania, USA; global"),
    "onlinecomponents": (None, "USA", "USA; global online"),
    "quest components": (None, "USA", "USA"),
    "classic components": (None, "USA", "USA; global stocking"),
    "sourceability": (None, "USA / global", "USA; global sourcing"),
    "octopart": (None, "New York, USA", "USA; global search"),
    "findchips": (None, "USA", "USA; global aggregation"),
    "corestaff": (None, "Tokyo, Japan", "Japan; Asia"),
    "electropeak": (None, "Iran / online", "Online international"),
    # Fabric
    "mood fabrics": (None, "225 W 37th Street, New York, NY, USA", "New York, NY, USA; Los Angeles; online"),
    "whaleys": (None, "Bradford, West Yorkshire, United Kingdom", "Bradford, UK; Europe"),
    "rose brand": (None, "New York / California, USA", "New York, USA; California, USA; entertainment industry"),
    "britex": (None, "San Francisco, California, USA", "San Francisco, California, USA"),
    "joann": (None, "Hudson, Ohio, USA", "Hudson, Ohio, USA; USA retail + online"),
    "hobby lobby": (None, "Oklahoma City, Oklahoma, USA", "Oklahoma City, Oklahoma, USA; USA stores"),
    "blick art": (None, "Galesburg, Illinois, USA", "Galesburg, Illinois, USA; USA stores + online"),
    "hancock fabrics": (None, "Baldwyn, Mississippi, USA (historic)", "USA (legacy retail)"),
    "fabric.com": (None, "Seattle / Amazon, USA", "USA online (Amazon)"),
    "onlinefabricstore": (None, "USA", "USA nationwide"),
    "fabric wholesale direct": (None, "USA", "USA"),
    "denver fabrics": (None, "Denver, Colorado, USA", "Denver, Colorado, USA"),
    "macculloch": (None, "London, United Kingdom", "London, UK"),
    "liberty fabrics": (None, "London, United Kingdom", "London, UK; international wholesale"),
    "spoonflower": (None, "Durham, North Carolina, USA", "Durham, North Carolina, USA; international online"),
    "coats group": (None, "Uxbidge / London area, United Kingdom", "UK; global"),
    "prym": (None, "Stolberg, Germany", "Germany; Europe; North America"),
    "gutermann": (None, "Gutach-Breisgau, Germany", "Germany; global via distributors"),
    "seattle fabrics": (None, "Seattle, Washington, USA", "Seattle, Washington, USA"),
    "vogue fabrics": (None, "Evanston / Chicago, Illinois, USA", "Evanston, Illinois, USA"),
    "stonemountain": (None, "Berkeley, California, USA", "Berkeley, California, USA"),
    "b&j fabrics": (None, "New York, NY, USA", "New York, NY, USA"),
    "spotlight": (None, "South Melbourne / Australia", "Australia; New Zealand; Asia retail"),
    "canvas etc": (None, "USA", "USA"),
    "muslin factory": (None, "18/30 Thiruvengadam Street, West Mambalam, Chennai 600033, India", "Chennai; Surat; Tirupur; Bangalore, India"),
    "michael levine": (None, "Los Angeles, California, USA", "Los Angeles, California, USA"),
    "fabric depot": (None, "Portland, Oregon, USA", "Portland, Oregon, USA"),
    "fashion fabrics club": (None, "USA", "USA"),
    "big duck canvas": (None, "USA", "USA"),
    "bigduckcanvas": (None, "USA", "USA"),
    "international silks": (None, "Los Angeles, California, USA", "Los Angeles, California, USA"),
    "syfabrics": (None, "Los Angeles, California, USA", "Los Angeles garment district, USA"),
    "ripstop by the roll": (None, "USA", "USA online"),
    "global sources": (None, "Hong Kong", "Hong Kong; Asia export; global buyers"),
    "thomasnet": (None, "New York, USA", "USA industrial directory"),
    "fabriclore": (None, "India", "India; global apparel brands"),
    "paron fabrics": (None, "New York, NY, USA", "New York, NY, USA"),
    "a & n fabrics": (None, "New York, NY, USA", "New York, NY, USA"),
    "lunn fabrics": (None, "USA", "USA theatrical"),
    "mathanika": (None, "Erode, Tamil Nadu, India", "Erode, India; export"),
    "xiaonizi": (None, "Shaoxing, China", "Shaoxing, China; global export"),
    "alibaba textile": (None, "Hangzhou, China", "Hangzhou, China; global B2B"),
    # Supplements
    "kehe": (None, "Naperville, Illinois, USA", "Naperville, Illinois, USA; USA natural retail"),
    "unfi": (None, "Providence, Rhode Island, USA", "Providence, Rhode Island, USA; USA; Canada"),
    "united natural foods": (None, "Providence, Rhode Island, USA", "Providence, Rhode Island, USA; USA; Canada"),
    "c&s wholesale": (None, "Keene, New Hampshire, USA", "Keene, New Hampshire, USA; USA grocery wholesale"),
    "now foods": (None, "Bloomingdale, Illinois, USA", "Bloomingdale, Illinois, USA; global distributors"),
    "nature's way": (None, "Green Bay, Wisconsin, USA", "Green Bay, Wisconsin, USA; global"),
    "puritan's pride": (None, "Ronkonkoma, New York, USA", "Ronkonkoma, New York, USA; international"),
    "swanson health": (None, "Fargo, North Dakota, USA", "Fargo, North Dakota, USA; export"),
    "vitacost": (None, "Boca Raton / USA (Kroger)", "USA online"),
    "garden of life": (None, "Palm Beach Gardens, Florida, USA", "Florida, USA; Nestlé Health Science channel"),
    "herb pharm": (None, "Williams, Oregon, USA", "Williams, Oregon, USA"),
    "herbalife": (None, "Los Angeles, California, USA", "Los Angeles, California, USA; global"),
    "gnc": (None, "Pittsburgh, Pennsylvania, USA", "Pittsburgh, Pennsylvania, USA; franchise international"),
    "vitamin shoppe": (None, "Secaucus, New Jersey, USA", "Secaucus, New Jersey, USA"),
    "iherb": (None, "Irvine, California, USA", "Irvine, California, USA; global online"),
    "artemis international": (None, "Fort Wayne, Indiana, USA", "Fort Wayne, Indiana, USA; global ingredient buyers"),
    "gaia herbs": (None, "Brevard, North Carolina, USA", "Brevard, North Carolina, USA"),
    "jarrow formulas": (None, "Los Angeles, California, USA", "Los Angeles, California, USA; international"),
    "solgar": (None, "Leonia, New Jersey, USA", "Leonia, New Jersey, USA; global"),
    "nordic naturals": (None, "Watsonville, California, USA", "Watsonville, California, USA; global"),
    "sundown naturals": (None, "USA", "USA mass retail"),
    "zahler": (None, "New York, USA", "New York, USA specialty retail"),
    "elderberry farms": (None, "Payson, Utah, USA", "Payson, Utah, USA"),
    "nutrascience labs": (None, "Farmingdale, New York, USA", "Farmingdale, New York, USA"),
    "fullscript": (None, "Ottawa, Canada / USA ops", "North America practitioner channel"),
    "emerson ecologics": (None, "USA", "USA practitioner"),
    "source naturals": (None, "Scotts Valley, California, USA", "Scotts Valley, California, USA"),
    "threshold enterprises": (None, "Scotts Valley, California, USA", "Scotts Valley, California, USA"),
    "europharma": (None, "Green Bay, Wisconsin, USA", "Green Bay, Wisconsin, USA"),
    "wellevate": (None, "USA", "USA practitioner"),
    "natural partners": (None, "USA", "USA practitioner"),
    "rainbow light": (None, "USA", "USA natural & mass"),
    "costco": (None, "999 Lake Drive, Issaquah, WA 98027, USA", "Issaquah, Washington, USA; global clubs"),
    "sam's club": (None, "Bentonville, Arkansas, USA", "Bentonville, Arkansas, USA; USA clubs"),
    "cvs": (None, "One CVS Drive, Woonsocket, RI 02895, USA", "Woonsocket, Rhode Island, USA"),
    "walgreens": (None, "Deerfield, Illinois, USA", "Deerfield, Illinois, USA; USA; UK Boots alliance"),
    "amazon": (None, "Seattle, Washington, USA", "Seattle, Washington, USA; global"),
    "thrive market": (None, "Los Angeles, California, USA", "Los Angeles, California, USA; USA online"),
    "hartmann's plant": (None, "USA", "USA"),
    "hartmannsplantcompany": (None, "USA", "USA"),
    "mcb group": (None, "Taiwan", "Taiwan; Asia; global ingredient B2B"),
    "medikonda": (None, "USA", "USA bulk B2B"),
    "smp nutra": (None, "USA", "USA private-label B2B"),
    # Baijiu / spirits
    "cns imports": (None, "USA", "USA; 30+ state distribution network"),
    "cnsimports": (None, "USA", "USA; 30+ state distribution network"),
    "snowbridge": (None, "Seattle, Washington, USA", "Seattle, Washington, USA"),
    "mhw ltd": (None, "USA / Almere, Netherlands (EU hub)", "USA; European Union"),
    "southern glazer": (None, "USA (multi-state HQ network)", "USA multi-state wholesale"),
    "breakthru beverage": (None, "USA", "USA multi-state"),
    "rndc": (None, "USA", "USA"),
    "republic national": (None, "USA", "USA"),
    "young's market": (None, "USA West Coast", "Western USA"),
    "johnson brothers": (None, "St. Paul, Minnesota, USA", "Midwest USA"),
    "winebow": (None, "USA", "USA"),
    "cofco": (None, "Beijing, China", "Beijing, China; China; export"),
    "jd.com": (None, "Beijing, China", "Beijing, China; China e-commerce"),
    "alibaba": (None, "Hangzhou, Zhejiang, China", "Hangzhou, China; global"),
    "tmall": (None, "Hangzhou, China", "China online"),
    "pinduoduo": (None, "Shanghai, China", "Shanghai, China; China online"),
    "dfs group": (None, "Hong Kong", "Hong Kong; global travel retail"),
    "total wine": (None, "Bethesda, Maryland, USA", "Bethesda, Maryland, USA; USA multi-state"),
    "bevmo": (None, "Concord / California, USA", "Western USA"),
    "astor wines": (None, "New York, NY, USA", "New York, NY, USA; online"),
    "k&l wine": (None, "California, USA", "California, USA; online"),
    "diageo": (None, "London, United Kingdom", "London, UK; China; global"),
    "pernod ricard": (None, "Paris, France", "Paris, France; China; Asia"),
    "remy cointreau": (None, "Cognac / Paris, France", "France; China premium on-trade"),
    "h mart": (None, "Lyndhurst, New Jersey, USA", "New Jersey, USA; USA Asian grocery"),
    "99 ranch": (None, "Buena Park, California, USA", "California, USA; USA West/Southwest"),
    "moutai": (None, "Renhuai, Guizhou, China", "Guizhou, China; duty-free; export"),
    "wuliangye": (None, "Yibin, Sichuan, China", "Yibin, Sichuan, China; export"),
    "luzhou laojiao": (None, "Luzhou, Sichuan, China", "Luzhou, Sichuan, China; export"),
    "yanghe": (None, "Suqian, Jiangsu, China", "Suqian, Jiangsu, China; export"),
    "gallo": (None, "Modesto, California, USA", "Modesto, California, USA"),
    "horizon beverage": (None, "USA Northeast", "Northeast USA"),
    "opici": (None, "Glen Rock / New Jersey, USA", "USA select states"),
    "asc fine wines": (None, "Shanghai / China", "China major cities"),
    "summergate": (None, "China", "China"),
    "china duty free": (None, "Beijing / China", "China travel retail"),
    "duty free americas": (None, "USA", "Travel retail Americas"),
    "spec's wines": (None, "Houston, Texas, USA", "Texas, USA"),
    "jazz trading": (None, "London, United Kingdom", "London, UK"),
    "beijing shunxin": (None, "Beijing, China", "Beijing, China"),
    # Bike
    "quality bicycle products": (None, "6400 W 105th Street, Bloomington, MN 55438, USA", "Bloomington, Minnesota; Brighton, CO; Lancaster, PA; Sparks, NV"),
    "qbp": (None, "6400 W 105th Street, Bloomington, MN 55438, USA", "Bloomington, Minnesota, USA; USA; Canada dealers"),
    "j&b importers": (None, "Miami, Florida, USA", "Miami, Florida, USA; Latin America"),
    "fox racing shox": (None, "El Cajon / Southern California, USA", "California, USA; global via distributors"),
    "sram": (None, "Chicago, Illinois, USA", "Chicago, Illinois, USA; global"),
    "rockshox": (None, "Colorado Springs / SRAM, USA", "USA; global"),
    "öhlins": (None, "Upplands Väsby, Sweden", "Sweden; global premium MTB"),
    "ohlins": (None, "Upplands Väsby, Sweden", "Sweden; global premium MTB"),
    "trek bicycle": (None, "Waterloo, Wisconsin, USA", "Waterloo, Wisconsin, USA; global"),
    "specialized bicycle": (None, "Morgan Hill, California, USA", "Morgan Hill, California, USA; global"),
    "giant bicycle": (None, "Taichung, Taiwan", "Taiwan; global dealer networks"),
    "canyon bicycles": (None, "Koblenz, Germany", "Koblenz, Germany; Europe; USA; global DTC"),
    "shimano": (None, "Sakai, Osaka, Japan", "Japan; global"),
    "santa cruz bicycles": (None, "Santa Cruz, California, USA", "Santa Cruz, California, USA; global dealers"),
    "yeti cycles": (None, "Golden, Colorado, USA", "Colorado, USA; global"),
    "pivot cycles": (None, "Tempe, Arizona, USA", "Arizona, USA; global"),
    "ibis cycles": (None, "Santa Cruz area, California, USA", "California, USA; global"),
    "cane creek": (None, "Fletcher, North Carolina, USA", "Fletcher, North Carolina, USA; global"),
    "manitou": (None, "USA (Hayes Performance Systems)", "USA; global"),
    "dvo suspension": (None, "USA", "USA; global dealer networks"),
    "pon.bike": (None, "Netherlands", "Netherlands; global brand groups"),
    "decathlon": (None, "Villeneuve-d'Ascq, France", "France; global retail"),
    "chain reaction cycles": (None, "Northern Ireland / UK", "UK; Europe; global online"),
    "bike24": (None, "Dresden, Germany", "Dresden, Germany; Europe; global shipping"),
    "jenson usa": (None, "Riverside / California, USA", "California, USA; USA online"),
    "universal cycles": (None, "USA", "USA online/wholesale"),
    "bicycle warehouse": (None, "USA", "USA"),
    "performance bicycle": (None, "USA", "USA"),
    "competitive cyclist": (None, "Salt Lake City area, Utah, USA", "Utah, USA; USA online"),
    "accell group": (None, "Heerenveen, Netherlands", "Netherlands; Europe; North America"),
    "bilstein": (None, "Ennepetal, Germany", "Ennepetal, Germany; global"),
    "koni": (None, "Oud-Beijerland, Netherlands", "Netherlands; global"),
    "wiggle": (None, "United Kingdom", "UK; Europe; global online"),
    "marzocchi": (None, "Italy", "Italy; Europe; global"),
    "push industries": (None, "Colorado, USA", "Colorado, USA; export dealers"),
    "pushindustries": (None, "Colorado, USA", "Colorado, USA"),
    "downeast bicycle": (None, "USA East", "Eastern USA dealers"),
    "hans johnsen": (None, "Dallas, Texas, USA", "Dallas, Texas, USA; USA South/Central"),
    "khs bicycles": (None, "USA / Taiwan brand roots", "USA dealers"),
    "khs distribution": (None, "USA", "USA dealers"),
    "the pro's closet": (None, "Colorado, USA", "Colorado, USA"),
    "bikediscount": (None, "Germany", "Germany; Europe online"),
    "99 bikes": (None, "Australia", "Australia"),
    "bikesonline": (None, "Australia", "Australia; NZ; USA DTC"),
    "worldwide cyclery": (None, "USA", "USA online"),
    "alltricks": (None, "France", "France; Europe"),
    # Beauty / nails
    "sally beauty": (None, "7900 Windrose Ave, Plano, TX 75024, USA", "Plano, Texas, USA; USA; Canada; Europe; LatAm"),
    "cosmoprof": (None, "Plano / Denton metro, Texas, USA (Sally Beauty Holdings)", "North America professional-only"),
    "armstrong mccall": (None, "Texas, USA (Sally Beauty Holdings)", "USA (BSG)"),
    "beauty systems group": (None, "Plano / Texas, USA (Sally Beauty Holdings)", "North America"),
    "paramount beauty": (None, "125 Commerce Dr, Hauppauge, NY 11788, USA", "Hauppauge, New York, USA; US East Coast"),
    "burmax": (None, "USA", "USA salon wholesale"),
    "saloncentric": (None, "USA (L'Oréal USA)", "USA professional"),
    "ulta beauty": (None, "Bolingbrook, Illinois, USA", "Bolingbrook, Illinois, USA; USA retail"),
    "sephora": (None, "San Francisco, California, USA / Paris roots", "San Francisco, USA; global retail"),
    "opi": (None, "USA (Wella Company)", "USA; global"),
    "cnd (creative": (None, "USA", "USA; global professional"),
    "gelish": (None, "USA", "USA; global"),
    "essie": (None, "USA / L'Oréal", "USA; global mass & salon"),
    "revlon": (None, "New York, NY, USA", "New York, USA; global mass"),
    "orly": (None, "Los Angeles, California, USA", "Los Angeles, California, USA; international"),
    "china glaze": (None, "USA (American International Industries)", "USA; export"),
    "kiara sky": (None, "USA", "USA; global e-comm"),
    "young nails": (None, "USA", "USA; export"),
    "beauty bay": (None, "United Kingdom", "UK; Europe; online global"),
    "lookfantastic": (None, "Manchester / UK (THG)", "UK; Europe; global"),
    "cult beauty": (None, "United Kingdom", "UK; Europe"),
    "douglas": (None, "Düsseldorf, Germany", "Düsseldorf, Germany; Europe"),
    "notino": (None, "Czech Republic", "Czech Republic; Europe online"),
    "superdrug": (None, "Croydon / United Kingdom", "United Kingdom"),
    "boots uk": (None, "Nottingham, United Kingdom", "UK; Ireland"),
    "target": (None, "1000 Nicollet Mall, Minneapolis, MN, USA", "Minneapolis, Minnesota, USA; USA stores"),
    "walmart": (None, "702 SW 8th Street, Bentonville, AR 72716, USA", "Bentonville, Arkansas, USA; USA; international"),
    "nail systems international": (None, "USA", "USA; global professional"),
    "beauty joint": (None, "USA", "USA online"),
    "light elegance": (None, "USA", "USA; international"),
    "marianna industries": (None, "USA", "USA"),
    "chemistwarehouse": (None, "Australia", "Australia"),
    "chemist warehouse": (None, "Australia", "Australia"),
    "jessica cosmetics": (None, "USA", "USA"),
    "jessicacosmetics": (None, "USA", "USA"),
    "born pretty": (None, "China / global online", "China; global e-commerce"),
    "makartt": (None, "USA / China ops", "USA; global online"),
    "modelones": (None, "USA / China", "USA; global online"),
    "beauty bridge": (None, "USA", "USA"),
    # K-beauty / masks
    "olive young": (None, "Seoul, South Korea", "Seoul, South Korea; global online"),
    "innisfree": (None, "Seoul, South Korea (Amorepacific)", "Seoul, South Korea; global"),
    "laneige": (None, "Seoul, South Korea (Amorepacific)", "Seoul, South Korea; global"),
    "mediheal": (None, "Seoul, South Korea", "Korea; USA; global"),
    "dr. jart": (None, "Seoul, South Korea / Estée Lauder", "Korea; USA; global"),
    "cosrx": (None, "Seoul, South Korea", "Seoul, South Korea; global"),
    "tony moly": (None, "Seoul, South Korea", "Korea; USA; global"),
    "etude": (None, "Seoul, South Korea (Amorepacific)", "Korea; Asia; global"),
    "the face shop": (None, "Seoul, South Korea (LG H&H)", "Korea; Asia; global"),
    "watsons": (None, "Hong Kong (AS Watson)", "Hong Kong; Asia; Europe"),
    "guardian /": (None, "Singapore / Asia (AS Watson)", "Southeast Asia"),
    "stylekorean": (None, "South Korea", "Korea; global online"),
    "wishtrend": (None, "Seoul, South Korea", "Seoul, South Korea; global K-beauty"),
    "sokoglam": (None, "New York / USA", "USA; online"),
    "soko glam": (None, "New York / USA", "USA; online"),
    "yesstyle": (None, "Hong Kong", "Hong Kong; global online"),
    "jolse": (None, "South Korea", "Korea; global online"),
    "peach & lily": (None, "USA", "USA; Ulta/Sephora channels"),
    "beautylish": (None, "USA", "USA online"),
    "skinstore": (None, "USA", "USA online"),
    "cj olive": (None, "Seoul, South Korea", "Korea; export"),
    "tjx": (None, "Framingham, Massachusetts, USA", "USA; Europe; Australia"),
    "jumia": (None, "Nigeria / Africa (Pan-African)", "Africa e-commerce"),
    "roseroseshop": (None, "South Korea", "Korea; global"),
    "beautynetkorea": (None, "South Korea", "Korea; global B2B/B2C"),
    "glow recipe": (None, "New York, USA", "New York, USA; Sephora channel"),
    "amorepacific": (None, "Seoul, South Korea", "Seoul, South Korea; global"),
    # Additional verified fills (web search)
    "fabric depot": ("1992", "700 SE 122nd Ave, Portland, OR 97233, USA (historic; closed 2018)", "Portland, Oregon, USA"),
    "the fabric store": ("1995", "Auckland, New Zealand", "Auckland, Wellington, Christchurch, Dunedin NZ; Sydney, Melbourne, Brisbane AU; online international"),
    "beauty joint": ("2011", "1636 W 8th Street Ste 200, Los Angeles, CA 90017, USA", "Los Angeles, California, USA; online"),
    "nail tech supply": (None, "2100 S Great Southwest Pkwy, Grand Prairie, TX 75051, USA", "Grand Prairie, Texas; Kansas City / Gladstone, Missouri"),
    "chemistwarehouse": ("2000", "Preston, Victoria, Australia", "Preston, Victoria, Australia; Australia; NZ; China; Ireland"),
    "chemist warehouse": ("2000", "Preston, Victoria, Australia", "Preston, Victoria, Australia; Australia; NZ"),
    "mathanika": (None, "37 Marappa Gounder 1st Street, Erode, Tamil Nadu 638009, India", "Erode, Tamil Nadu, India; export"),
    "yetis cycles": (None, "Golden, Colorado, USA", "Golden, Colorado, USA; global"),
    "yeti cycles": (None, "Golden, Colorado, USA", "Golden, Colorado, USA; global"),
    "harris cyclery": (None, "West Newton / Boston area, Massachusetts, USA", "Massachusetts, USA"),
    "drinkupny": (None, "New York, USA", "New York, USA; USA shipping states"),
    "beijing shunxin": (None, "Beijing, China", "Beijing, China"),
    "premier textiles": (None, "USA", "USA apparel textile wholesale"),
    "syfabrics": (None, "Los Angeles, California, USA", "Los Angeles garment district, USA"),
    "sewing parts online": (None, "USA", "USA online"),
    "sewing machine plus": (None, "USA", "USA"),
    "sewing machine warehouse": (None, "USA", "USA"),
    "medshop": (None, "Australia", "Australia"),
    "moyra": (None, "Hungary / Europe", "Europe professional nail"),
    "nailish": (None, "Europe", "Europe"),
    "the nailhub": (None, "USA", "USA online"),
    "wholesale nail supply": (None, "USA", "USA"),
    "dear sundays": (None, "USA / online", "Online"),
    "dearsundays": (None, "USA / online", "Online"),
    "acecontrols": ("1957", "Farmington Hills, Michigan, USA", "Farmington Hills, Michigan, USA; global industrial"),
    "ace controls": ("1957", "Farmington Hills, Michigan, USA", "Farmington Hills, Michigan, USA"),
    "mountain medical": (None, "USA", "USA DME"),
    "performance inc": (None, "USA", "USA bike retail legacy"),
    "korean beauty supply": (None, "USA / Korea import channel", "USA; Korea"),
    "kotra": (None, "Seoul, South Korea", "Seoul, South Korea; global export promotion"),
    "mhw / european": (None, "Almere, Netherlands / USA", "Netherlands EU hub; USA; EU 27"),
    "guangdong jiujiu": (None, "Guangdong, China", "Guangdong, China"),
    "laojiangtaigroup": (None, "China", "China"),
    "shaoxing xiaonizi": (None, "Shaoxing, China", "Shaoxing, China; global export"),
    "audiology systems": (None, "USA", "USA"),
    "baijiushop": (None, "International online", "Select export markets"),
    "jbi.bike": (None, "Miami, Florida, USA (J&B Importers)", "Miami, Florida, USA; Latin America"),
    # Noodles / grocery / foodservice
    "sysco": (None, "1390 Enclave Parkway, Houston, TX 77077, USA", "Houston, Texas, USA; USA; international foodservice"),
    "us foods": (None, "9399 W Higgins Road, Rosemont, IL 60018, USA", "Rosemont, Illinois, USA; USA foodservice"),
    "performance food group": (None, "Richmond, Virginia, USA", "Richmond, Virginia, USA; USA"),
    "pfg": (None, "Richmond, Virginia, USA", "Richmond, Virginia, USA"),
    "gordon food service": (None, "Wyoming, Michigan, USA", "Michigan, USA; North America"),
    "dot foods": (None, "Mount Sterling, Illinois, USA", "Mount Sterling, Illinois, USA; USA"),
    "mclane": (None, "Temple, Texas, USA", "Temple, Texas, USA; USA"),
    "associated wholesale grocers": (None, "Kansas City, Kansas, USA", "Kansas City, Kansas, USA; USA cooperative"),
    "awg": (None, "Kansas City, Kansas, USA", "Kansas City, Kansas, USA"),
    "wakefern": (None, "Keasbey / New Jersey, USA", "New Jersey, USA; Northeast"),
    "ben e. keith": (None, "Fort Worth, Texas, USA", "Fort Worth, Texas, USA; USA South"),
    "cheney brothers": (None, "Florida / Southeast USA", "USA Southeast"),
    "bidfood": (None, "South Africa / Bidcorp global", "Global foodservice"),
    "bidcorp": (None, "Johannesburg, South Africa", "Global foodservice"),
    "metro ag": (None, "Düsseldorf, Germany", "Düsseldorf, Germany; Europe; Asia cash & carry"),
    "makro": (None, "Multi-country (Metro group)", "Global cash & carry"),
    "bj's wholesale": (None, "Westborough, Massachusetts, USA", "Massachusetts, USA; USA East"),
    "aeon": (None, "Chiba, Japan", "Chiba, Japan; Asia retail"),
    "seven & i": (None, "Tokyo, Japan", "Tokyo, Japan; global convenience"),
    "7-eleven": (None, "Irving, Texas, USA / Tokyo Seven & i", "USA; Japan; global convenience"),
    "familymart": (None, "Tokyo, Japan", "Japan; Asia convenience"),
    "lawson": (None, "Tokyo, Japan", "Japan; Asia"),
    "nissin foods": (None, "Osaka / Tokyo, Japan", "Japan; global"),
    "toyo suisan": (None, "Tokyo, Japan", "Japan; USA (Maruchan); global"),
    "maruchan": (None, "Irvine, California, USA / Toyo Suisan", "California, USA; global"),
    "ottogi": (None, "Anyang / South Korea", "Korea; export"),
    "nongshim": (None, "Seoul, South Korea / Rancho Cucamonga, CA USA ops", "Korea; USA; global"),
    "samyang foods": (None, "Seoul, South Korea", "Korea; global export"),
    "indofood": (None, "Jakarta, Indonesia", "Jakarta, Indonesia; global export"),
    "indomie": (None, "Jakarta, Indonesia (Indofood)", "Indonesia; global export"),
    "nestlé": (None, "Vevey, Switzerland", "Vevey, Switzerland; global"),
    "nestle": (None, "Vevey, Switzerland", "Vevey, Switzerland; global"),
    "maggi": (None, "Vevey, Switzerland (Nestlé)", "Global"),
    "unilever": (None, "London, UK / Rotterdam, Netherlands", "UK; Netherlands; global"),
    "jfc": (None, "Los Angeles / USA (Nishimoto)", "USA; Canada Asian wholesale"),
    "season asian": (None, "USA (JFC / Nishimoto)", "USA; Canada"),
    "coupang": (None, "Seoul, South Korea", "Seoul, South Korea"),
    "rakuten": (None, "Tokyo, Japan", "Tokyo, Japan"),
}


FORCE_FOUNDED = {
    "henry schein": "1932",
    "opici family distributing": "1913",
    "alltricks (decathlon group)": "2008",
    "fabric depot": "1992",
    "beauty joint": "2011",
}

# Always overwrite these HQs (verified web/IR), even if a city already present
FORCE_HQ = {
    "sally beauty supply": (
        "7900 Windrose Ave, Plano, TX 75024, USA",
        "Plano, Texas, USA; USA; Canada; Europe; LatAm stores",
    ),
    "sally beauty": (
        "7900 Windrose Ave, Plano, TX 75024, USA",
        "Plano, Texas, USA; USA; Canada; Europe; LatAm",
    ),
    "cosmoprof (beauty systems group)": (
        "Plano, Texas, USA (Sally Beauty Holdings)",
        "North America professional-only",
    ),
    "cosmoprof": (
        "Plano, Texas, USA (Sally Beauty Holdings)",
        "North America professional-only",
    ),
    "armstrong mccall": (
        "Texas, USA (Sally Beauty Holdings / BSG)",
        "USA (BSG)",
    ),
    "beauty systems group direct sales": (
        "Plano, Texas, USA (Sally Beauty Holdings)",
        "North America",
    ),
    "quality bicycle products (qbp)": (
        "6400 W 105th Street, Bloomington, MN 55438, USA",
        "Bloomington, Minnesota; Brighton, CO; Lancaster, PA; Sparks, NV",
    ),
    "sysco corporation": (
        "1390 Enclave Parkway, Houston, TX 77077, USA",
        "Houston, Texas, USA; USA; international foodservice",
    ),
    "henry schein": (
        "135 Duryea Road, Melville, NY 11747, USA",
        "Melville, New York, USA; 34 countries",
    ),
    "henry schein medical": (
        "135 Duryea Road, Melville, NY 11747, USA",
        "Melville, New York, USA; global",
    ),
    "henry schein medical (ent/audiology)": (
        "135 Duryea Road, Melville, NY 11747, USA",
        "Melville, New York, USA; global",
    ),
}


def lookup(company: str) -> tuple[str | None, str | None, str | None]:
    n = norm(company)
    n_plain = re.sub(r"\([^)]*\)", " ", n)
    n_plain = re.sub(r"\s+", " ", n_plain).strip()

    founded = FORCE_FOUNDED.get(n) or FORCE_FOUNDED.get(n_plain)
    hq = cities = None

    hits: list[tuple[str, tuple]] = []
    for k, v in GEO.items():
        if len(k) <= 3:
            if re.search(rf"(^|[^a-z0-9]){re.escape(k)}([^a-z0-9]|$)", n_plain):
                hits.append((k, v))
        elif k in n or k in n_plain:
            hits.append((k, v))
    if hits:
        hits.sort(key=lambda kv: len(kv[0]), reverse=True)
        f, h, c = hits[0][1]
        if founded is None:
            founded = f
        hq, cities = h, c
    return founded, hq, cities


def process(path: Path) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=False))
    hdr_i = next(i for i, r in enumerate(rows[:12]) if r and r[0].value and str(r[0].value).strip() == "Company")
    header = [str(c.value or "").strip() for c in rows[hdr_i]]
    ci, fi = header.index("Company"), header.index("Founded")
    hi = header.index("Headquarters")
    ri = header.index("Cities / Regions")

    st = {
        "file": path.name,
        "founded_filled": 0,
        "founded_corrected": 0,
        "hq_filled": 0,
        "cities_filled": 0,
        "still_founded": 0,
        "still_hq": 0,
        "still_cities": 0,
        "total": 0,
    }

    for r in rows[hdr_i + 1 :]:
        company = str(r[ci].value or "").strip()
        if not company:
            continue
        st["total"] += 1
        f_new, h_new, c_new = lookup(company)

        nn = norm(company)
        plain = re.sub(r"\s+", " ", re.sub(r"\([^)]*\)", " ", nn)).strip()
        if nn in FORCE_FOUNDED or plain in FORCE_FOUNDED:
            f_new = FORCE_FOUNDED.get(nn) or FORCE_FOUNDED.get(plain)

        # Force HQ overrides (verified)
        if nn in FORCE_HQ:
            h_new, c_new = FORCE_HQ[nn]
        elif plain in FORCE_HQ:
            h_new, c_new = FORCE_HQ[plain]
        else:
            for fk, fv in FORCE_HQ.items():
                if fk in nn or fk in plain:
                    h_new, c_new = fv
                    break

        # Founded: fill empty OR correct when we have verified year
        if f_new:
            old = str(r[fi].value or "").strip()
            if is_empty(r[fi].value):
                r[fi].value = f_new
                st["founded_filled"] += 1
            elif old != f_new and (
                nn in FORCE_FOUNDED
                or plain in FORCE_FOUNDED
                or "fabric depot" in nn
                or "beauty joint" in nn
                or "chemistwarehouse" in nn
                or "chemist warehouse" in nn
            ):
                r[fi].value = f_new
                st["founded_corrected"] += 1
        elif is_empty(r[fi].value):
            r[fi].value = "Not confirmed publicly"
            st["still_founded"] += 1

        # HQ — force override OR upgrade weak
        forced_hq = nn in FORCE_HQ or plain in FORCE_HQ or any(fk in nn for fk in FORCE_HQ)
        if h_new and (forced_hq or is_weak_hq(r[hi].value) or (str(r[hi].value).strip().lower() == "usa" and "," in h_new.lower())):
            old_h = str(r[hi].value or "").strip()
            if forced_hq or is_weak_hq(old_h) or len(h_new) > len(old_h) + 5:
                if old_h != h_new:
                    r[hi].value = h_new
                    st["hq_filled"] += 1
        elif is_weak_hq(r[hi].value):
            if is_empty(r[hi].value) or str(r[hi].value).strip().lower() in EMPTY:
                r[hi].value = "Not confirmed publicly"
            st["still_hq"] += 1

        if c_new and (forced_hq or is_weak_cities(r[ri].value)):
            if str(r[ri].value or "").strip() != c_new:
                r[ri].value = c_new
                st["cities_filled"] += 1
        elif is_weak_cities(r[ri].value):
            if is_empty(r[ri].value) or str(r[ri].value).strip().lower() in EMPTY:
                r[ri].value = "Not confirmed publicly"
            st["still_cities"] += 1

    wb.save(path)
    wb.close()
    return st


def main() -> None:
    files = sorted(ROOT.rglob("*_distributors_expanded.xlsx"))
    g = {
        k: 0
        for k in [
            "total",
            "founded_filled",
            "founded_corrected",
            "hq_filled",
            "cities_filled",
            "still_founded",
            "still_hq",
            "still_cities",
        ]
    }
    for p in files:
        st = process(p)
        for k in g:
            g[k] += st[k]
        print(
            f"{st['file']}: founded+{st['founded_filled']}/corr{st['founded_corrected']} "
            f"hq+{st['hq_filled']} cities+{st['cities_filled']} "
            f"| still f/hq/c={st['still_founded']}/{st['still_hq']}/{st['still_cities']}"
        )
    print("\nGRAND", g)


if __name__ == "__main__":
    main()
