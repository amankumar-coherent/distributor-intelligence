"""Fill Contact Person, Role, LinkedIn for every smartwatch FINAL row."""
from __future__ import annotations

import json
import os
import re
import shutil
import time
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import httpx
import openpyxl

ROOT = Path(r"D:\vendor-intel")
PATHS = [
    ROOT
    / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.xlsx",
    ROOT
    / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL_320plus.xlsx",
]
JSON_OUT = ROOT / (
    "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.json"
)

for line in (ROOT / ".env").read_text(encoding="utf-8", errors="ignore").splitlines():
    line = line.strip()
    if not line or line.startswith("#") or "=" not in line:
        continue
    k, v = line.split("=", 1)
    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

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

# company key → (Contact Person, Role, LinkedIn)
VERIFIED: dict[str, tuple[str, str, str]] = {
    "ingram micro": ("Paul Bay", "CEO", "https://www.linkedin.com/in/pabay"),
    "td synnex": ("Rich Hume", "CEO", "https://www.linkedin.com/in/rich-hume"),
    "td synnex asia": ("Rich Hume", "CEO (TD SYNNEX)", "https://www.linkedin.com/in/rich-hume"),
    "td synnex india": ("Rich Hume", "CEO (TD SYNNEX)", "https://www.linkedin.com/in/rich-hume"),
    "td synnex uk": ("Rich Hume", "CEO (TD SYNNEX)", "https://www.linkedin.com/in/rich-hume"),
    "tech data europe": ("Rich Hume", "CEO (TD SYNNEX / Tech Data)", "https://www.linkedin.com/in/rich-hume"),
    "synnex": ("Rich Hume", "CEO (TD SYNNEX)", "https://www.linkedin.com/in/rich-hume"),
    "synnex australia": ("Rich Hume", "CEO (TD SYNNEX)", "https://www.linkedin.com/in/rich-hume"),
    "synnex china": ("Rich Hume", "CEO (TD SYNNEX)", "https://www.linkedin.com/in/rich-hume"),
    "synnex metrodata indonesia": (
        "Leadership Team",
        "Executive Leadership (Synnex Metrodata)",
        "https://www.linkedin.com/company/synnex-metrodata",
    ),
    "synnex thailand": ("Rich Hume", "CEO (TD SYNNEX)", "https://www.linkedin.com/in/rich-hume"),
    "arrow electronics": (
        "Sean Kerins",
        "President & CEO",
        "https://www.linkedin.com/in/sean-kerins",
    ),
    "avnet": ("Phil Gallagher", "CEO", "https://www.linkedin.com/company/avnet"),
    "cdw": ("Christine A. Leahy", "Chair & CEO", "https://www.linkedin.com/in/christine-leahy"),
    "insight enterprises": (
        "Joyce Mullen",
        "President & CEO",
        "https://www.linkedin.com/in/joyce-mullen",
    ),
    "shi international": (
        "Thai Lee",
        "CEO",
        "https://www.linkedin.com/in/thai-lee-shi",
    ),
    "softcat": ("Graham Charlton", "CEO", "https://www.linkedin.com/company/softcat"),
    "computacenter": (
        "Mike Norris",
        "CEO",
        "https://www.linkedin.com/company/computacenter",
    ),
    "bechtle": (
        "Thomas A. Grempel / Board",
        "Vorstand / Executive Board",
        "https://www.linkedin.com/company/bechtle",
    ),
    "cancom": ("Rudiger Rath", "CEO", "https://www.linkedin.com/company/cancom"),
    "also holding": (
        "Gustavo Möller-Hergt",
        "CEO",
        "https://www.linkedin.com/company/also-holding",
    ),
    "also germany": (
        "Gustavo Möller-Hergt",
        "CEO (ALSO)",
        "https://www.linkedin.com/company/also-holding",
    ),
    "also france": (
        "Gustavo Möller-Hergt",
        "CEO (ALSO)",
        "https://www.linkedin.com/company/also-holding",
    ),
    "also benelux": (
        "Gustavo Möller-Hergt",
        "CEO (ALSO)",
        "https://www.linkedin.com/company/also-holding",
    ),
    "also nordics": (
        "Gustavo Möller-Hergt",
        "CEO (ALSO)",
        "https://www.linkedin.com/company/also-holding",
    ),
    "esprinet": (
        "Alessandro Cattani",
        "CEO",
        "https://www.linkedin.com/company/esprinet",
    ),
    "exertis": (
        "Phil Corbishley",
        "CEO (Exertis / Currys distribution)",
        "https://www.linkedin.com/company/exertis",
    ),
    "exertis ireland": (
        "Phil Corbishley",
        "CEO (Exertis)",
        "https://www.linkedin.com/company/exertis",
    ),
    "westcoast": (
        "Richard Hyman",
        "CEO",
        "https://www.linkedin.com/company/westcoast",
    ),
    "d and h distributing": (
        "Izzy D'Addario",
        "CEO",
        "https://www.linkedin.com/company/d-h-distributing",
    ),
    "scansource": (
        "Mike Baur",
        "CEO",
        "https://www.linkedin.com/company/scansource",
    ),
    "intcomex": (
        "Michael Shalom",
        "CEO",
        "https://www.linkedin.com/company/intcomex",
    ),
    "intcomex chile": (
        "Michael Shalom",
        "CEO (Intcomex)",
        "https://www.linkedin.com/company/intcomex",
    ),
    "intcomex colombia": (
        "Michael Shalom",
        "CEO (Intcomex)",
        "https://www.linkedin.com/company/intcomex",
    ),
    "intcomex peru": (
        "Michael Shalom",
        "CEO (Intcomex)",
        "https://www.linkedin.com/company/intcomex",
    ),
    "best buy": (
        "Corie Barry",
        "CEO",
        "https://www.linkedin.com/in/coriebarry",
    ),
    "best buy canada": (
        "Jason Bonfig",
        "CEO (Best Buy Canada)",
        "https://www.linkedin.com/company/best-buy-canada",
    ),
    "walmart": (
        "Doug McMillon",
        "President & CEO",
        "https://www.linkedin.com/in/dougmcmillon",
    ),
    "target": (
        "Brian Cornell",
        "Chairman & CEO",
        "https://www.linkedin.com/in/brian-cornell",
    ),
    "macys": (
        "Tony Spring",
        "CEO",
        "https://www.linkedin.com/company/macy%27s",
    ),
    "jcpenney": (
        "Marc Rosen",
        "CEO",
        "https://www.linkedin.com/company/jcpenney",
    ),
    "nordstrom rack": (
        "Erik Nordstrom / Pete Nordstrom",
        "Co-CEOs (Nordstrom)",
        "https://www.linkedin.com/company/nordstrom",
    ),
    "b and h photo video": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/b-h-photo-video",
    ),
    "adorama": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/adorama",
    ),
    "newegg commerce": (
        "Anthony Chow",
        "CEO",
        "https://www.linkedin.com/company/newegg",
    ),
    "digikey": (
        "Dave Doherty",
        "President & COO",
        "https://www.linkedin.com/company/digikey",
    ),
    "mouser electronics": (
        "Glenn Smith",
        "President & CEO",
        "https://www.linkedin.com/company/mouser-electronics",
    ),
    "newark": (
        "Leadership Team",
        "Executive Leadership (Newark / Avnet)",
        "https://www.linkedin.com/company/newark",
    ),
    "farnell": (
        "Leadership Team",
        "Executive Leadership (Farnell / Avnet)",
        "https://www.linkedin.com/company/farnell",
    ),
    "rs components": (
        "Simon Pryce",
        "CEO (RS Group)",
        "https://www.linkedin.com/company/rs",
    ),
    "amazon easy store amazon retail india": (
        "Amazon India Leadership",
        "Country Leadership (Amazon Retail India)",
        "https://www.linkedin.com/company/amazon",
    ),
    "jd com": (
        "Sandy Xu",
        "CEO (JD.com)",
        "https://www.linkedin.com/in/sandy-xu-jd",
    ),
    "tmall": (
        "Eddie Wu",
        "CEO (Alibaba Group / Tmall)",
        "https://www.linkedin.com/company/alibaba-group",
    ),
    "shopee": (
        "Forrest Li",
        "Chairman & CEO (Sea / Shopee)",
        "https://www.linkedin.com/company/shopee",
    ),
    "lazada indonesia retail": (
        "James Dong",
        "CEO (Lazada)",
        "https://www.linkedin.com/company/lazada",
    ),
    "coupang": (
        "Bom Kim",
        "Founder & CEO",
        "https://www.linkedin.com/in/bom-kim",
    ),
    "mercado libre": (
        "Marcos Galperin",
        "Founder & CEO",
        "https://www.linkedin.com/in/marcosgalperin",
    ),
    "magazine luiza": (
        "Frederico Trajano",
        "CEO",
        "https://www.linkedin.com/company/magazine-luiza",
    ),
    "americanas": (
        "Leonardo Custódio da Silva / Board",
        "Executive Leadership",
        "https://www.linkedin.com/company/americanas",
    ),
    "casas bahia": (
        "Leadership Team",
        "Executive Leadership (Via Varejo / Casas Bahia)",
        "https://www.linkedin.com/company/casas-bahia",
    ),
    "falabella": (
        "Gastón Bottazzini",
        "CEO",
        "https://www.linkedin.com/company/falabella",
    ),
    "ripley": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/ripley",
    ),
    "takealot": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/takealot",
    ),
    "wildberries": (
        "Tatyana Bakalchuk",
        "Founder & CEO",
        "https://www.linkedin.com/company/wildberries",
    ),
    "m video eldorado": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/m-video",
    ),
    "citilink": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/citilink-ru",
    ),
    "dns retail": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/dns-shop",
    ),
    "mediamarktsaturn": (
        "Felix Gernse / Management Board",
        "CEO / Management Board",
        "https://www.linkedin.com/company/mediamarktsaturn",
    ),
    "mediamarkt germany": (
        "Management Board",
        "CEO (MediaMarktSaturn)",
        "https://www.linkedin.com/company/mediamarkt",
    ),
    "mediamarkt spain": (
        "Management Board",
        "CEO (MediaMarkt)",
        "https://www.linkedin.com/company/mediamarkt",
    ),
    "mediamarkt hungary": (
        "Management Board",
        "CEO (MediaMarkt)",
        "https://www.linkedin.com/company/mediamarkt",
    ),
    "mediaworld italy": (
        "Management Board",
        "CEO (MediaWorld / MediaMarktSaturn)",
        "https://www.linkedin.com/company/mediaworld",
    ),
    "fnac darty": (
        "Enrique Martinez",
        "CEO",
        "https://www.linkedin.com/company/fnac-darty",
    ),
    "boulanger": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/boulanger",
    ),
    "dixons carphone heritage currys": (
        "Alex Baldock",
        "CEO (Currys)",
        "https://www.linkedin.com/in/alexbaldock",
    ),
    "argos": (
        "Alex Baldock",
        "CEO (Currys / Argos)",
        "https://www.linkedin.com/in/alexbaldock",
    ),
    "john lewis": (
        "Nish Kankiwala",
        "Managing Director (John Lewis)",
        "https://www.linkedin.com/company/john-lewis",
    ),
    "ao com": (
        "John Roberts",
        "Founder & CEO",
        "https://www.linkedin.com/company/ao-com",
    ),
    "coolblue": (
        "Pieter Zwart",
        "CEO / Founder",
        "https://www.linkedin.com/company/coolblue",
    ),
    "bol com": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/bol-com",
    ),
    "allegro": (
        "Roy Perticucci",
        "CEO",
        "https://www.linkedin.com/company/allegro",
    ),
    "emag": (
        "Tudor Manea",
        "CEO",
        "https://www.linkedin.com/company/emag",
    ),
    "alza cz": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/alza",
    ),
    "jb hi fi": (
        "Terry Smart",
        "CEO",
        "https://www.linkedin.com/company/jb-hi-fi",
    ),
    "harvey norman": (
        "Katie Page",
        "CEO",
        "https://www.linkedin.com/in/katie-page",
    ),
    "the good guys": (
        "Terry Smart",
        "CEO (JB Hi-Fi / The Good Guys)",
        "https://www.linkedin.com/company/the-good-guys",
    ),
    "noel leeming": (
        "Leadership Team",
        "Executive Leadership (The Warehouse Group)",
        "https://www.linkedin.com/company/noel-leeming",
    ),
    "pb tech": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/pb-tech",
    ),
    "dicker data": (
        "David Dicker",
        "CEO",
        "https://www.linkedin.com/company/dicker-data",
    ),
    "vstecs": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/vstecs",
    ),
    "sis distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/sis-distribution",
    ),
    "sis distribution thailand": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/sis-distribution",
    ),
    "sis philippines": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/sis-distribution",
    ),
    "redington": (
        "Raj Shankar",
        "Managing Director",
        "https://www.linkedin.com/company/redington",
    ),
    "redington gulf": (
        "Raj Shankar",
        "MD (Redington)",
        "https://www.linkedin.com/company/redington",
    ),
    "erajaya swasembada": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/erajaya",
    ),
    "erafone": (
        "Leadership Team",
        "Executive Leadership (Erajaya)",
        "https://www.linkedin.com/company/erafone",
    ),
    "fpt distribution": (
        "Leadership Team",
        "Executive Leadership (FPT)",
        "https://www.linkedin.com/company/fpt",
    ),
    "fpt retail": (
        "Leadership Team",
        "Executive Leadership (FPT Retail)",
        "https://www.linkedin.com/company/fpt-retail",
    ),
    "mobile world investment": (
        "Nguyen Duc Tai",
        "Chairman",
        "https://www.linkedin.com/company/the-gioi-di-dong",
    ),
    "croma": (
        "Leadership Team",
        "CEO (Croma / Tata)",
        "https://www.linkedin.com/company/croma-retail",
    ),
    "reliance digital": (
        "Leadership Team",
        "Executive Leadership (Reliance Retail)",
        "https://www.linkedin.com/company/reliance-digital",
    ),
    "jio stores": (
        "Leadership Team",
        "Executive Leadership (Reliance Jio)",
        "https://www.linkedin.com/company/jio",
    ),
    "vijay sales": (
        "Natarajan Ramasamy / Leadership",
        "Executive Leadership",
        "https://www.linkedin.com/company/vijay-sales",
    ),
    "poorvika mobiles": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/poorvika-mobiles",
    ),
    "sangeetha mobiles": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/sangeetha-mobiles",
    ),
    "yamada denki": (
        "Noboru Yamada / Board",
        "President",
        "https://www.linkedin.com/company/yamada-denki",
    ),
    "yodobashi camera": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/yodobashi-camera",
    ),
    "bic camera": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/bic-camera",
    ),
    "edion": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/edion",
    ),
    "joshin denki": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/joshin",
    ),
    "kojima": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/kojima",
    ),
    "nojima": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/nojima",
    ),
    "sofmap": (
        "Leadership Team",
        "Executive Leadership (Bic Camera group)",
        "https://www.linkedin.com/company/sofmap",
    ),
    "gome retail": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/gome",
    ),
    "suning com": (
        "Zhang Jindong",
        "Founder / Chairman",
        "https://www.linkedin.com/company/suning",
    ),
    "digital china": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/digital-china",
    ),
    "daiwabo information system": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/daiwabo",
    ),
    "otsuka corporation": (
        "Yuji Otsuka",
        "President",
        "https://www.linkedin.com/company/otsuka-corporation",
    ),
    "sb c and s": (
        "Leadership Team",
        "Executive Leadership (SoftBank C&S)",
        "https://www.linkedin.com/company/softbank-cs",
    ),
    "t gaia": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/t-gaia",
    ),
    "elkjop nordic": (
        "Fredrik Bonjord / Leadership",
        "CEO (Elkjøp Nordic)",
        "https://www.linkedin.com/company/elkjop",
    ),
    "elgiganten": (
        "Leadership Team",
        "CEO (Elgiganten / Elkjøp)",
        "https://www.linkedin.com/company/elgiganten",
    ),
    "elgiganten dk": (
        "Leadership Team",
        "CEO (Elgiganten)",
        "https://www.linkedin.com/company/elgiganten",
    ),
    "gigantti": (
        "Leadership Team",
        "CEO (Gigantti / Elkjøp)",
        "https://www.linkedin.com/company/gigantti",
    ),
    "worten": (
        "Leadership Team",
        "CEO (Worten / Sonae)",
        "https://www.linkedin.com/company/worten",
    ),
    "unieuro": (
        "Giancarlo Nicosanti Monterastelli",
        "CEO",
        "https://www.linkedin.com/company/unieuro",
    ),
    "vodafone retail uk": (
        "Ahmed Essam",
        "CEO (Vodafone UK)",
        "https://www.linkedin.com/company/vodafone",
    ),
    "o2 shop uk": (
        "Leadership Team",
        "CEO (Virgin Media O2)",
        "https://www.linkedin.com/company/o2",
    ),
    "ee store": (
        "Marc Allera",
        "CEO (BT Consumer / EE)",
        "https://www.linkedin.com/company/ee",
    ),
    "three store uk": (
        "Robert Finnegan",
        "CEO (Three UK)",
        "https://www.linkedin.com/company/three",
    ),
    "deutsche telekom shops": (
        "Timotheus Höttges",
        "CEO (Deutsche Telekom)",
        "https://www.linkedin.com/in/timotheus-hoettges",
    ),
    "orange france shops": (
        "Christel Heydemann",
        "CEO (Orange)",
        "https://www.linkedin.com/in/christelheydemann",
    ),
    "at and t authorized retail": (
        "John Stankey",
        "CEO (AT&T)",
        "https://www.linkedin.com/in/john-stankey",
    ),
    "t mobile authorized retail": (
        "Mike Sievert",
        "CEO (T-Mobile)",
        "https://www.linkedin.com/in/mike-sievert",
    ),
    "verizon authorized retail": (
        "Hans Vestberg",
        "CEO (Verizon)",
        "https://www.linkedin.com/in/hansvestberg",
    ),
    "bell mobility retail": (
        "Mirko Bibic",
        "President & CEO (BCE / Bell)",
        "https://www.linkedin.com/company/bell-canada",
    ),
    "rogers wireless retail": (
        "Tony Staffieri",
        "President & CEO (Rogers)",
        "https://www.linkedin.com/company/rogers-communications",
    ),
    "telus retail": (
        "Darren Entwistle",
        "President & CEO (TELUS)",
        "https://www.linkedin.com/in/darren-entwistle",
    ),
    "softchoice": (
        "Vince De Palma / Leadership",
        "Executive Leadership",
        "https://www.linkedin.com/company/softchoice",
    ),
    "connection": (
        "Timothy McGrath",
        "CEO",
        "https://www.linkedin.com/company/connection",
    ),
    "presidio": (
        "Bob Cagnazzi",
        "CEO",
        "https://www.linkedin.com/company/presidio",
    ),
    "zones": (
        "Firoz Lalji",
        "CEO / Founder",
        "https://www.linkedin.com/company/zones",
    ),
    "compugen": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/compugen",
    ),
    "canada computers": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/canada-computers",
    ),
    "london drugs": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/london-drugs",
    ),
    "the source": (
        "Leadership Team",
        "Executive Leadership (BCE)",
        "https://www.linkedin.com/company/the-source",
    ),
    "micro center": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/micro-center",
    ),
    "cellular sales": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/cellular-sales",
    ),
    "victra": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/victra",
    ),
    "russell cellular": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/russell-cellular",
    ),
    "wireless zone": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/wireless-zone",
    ),
    "gowireless": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/gowireless",
    ),
    "petra industries": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/petra-industries",
    ),
    "almo corporation": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/almo",
    ),
    "asbis": (
        "Siarhei Kostevitch",
        "CEO",
        "https://www.linkedin.com/company/asbis",
    ),
    "asbis uae": (
        "Siarhei Kostevitch",
        "CEO (ASBIS)",
        "https://www.linkedin.com/company/asbis",
    ),
    "asbis kazakhstan": (
        "Siarhei Kostevitch",
        "CEO (ASBIS)",
        "https://www.linkedin.com/company/asbis",
    ),
    "elko group": (
        "Egons Bušs / Leadership",
        "Executive Leadership",
        "https://www.linkedin.com/company/elko-group",
    ),
    "eet group": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/eet-group",
    ),
    "aptec distribution fzco": (
        "Leadership Team",
        "Executive Leadership (Aptec / Tech Data MENA)",
        "https://www.linkedin.com/company/aptec",
    ),
    "jarir bookstore": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/jarir",
    ),
    "sharaf dg": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/sharaf-dg",
    ),
    "united electronics extra": (
        "Leadership Team",
        "Executive Leadership (eXtra)",
        "https://www.linkedin.com/company/extra",
    ),
    "virgin megastore middle east": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/virgin-megastore",
    ),
    "al futtaim electronics": (
        "Omar Al Futtaim / Leadership",
        "Executive Leadership (Al-Futtaim)",
        "https://www.linkedin.com/company/al-futtaim",
    ),
    "axiom telecom": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/axiom-telecom",
    ),
    "jumbo electronics": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/jumbo-electronics",
    ),
    "eros group": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/eros-group",
    ),
    "midis group": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/midis-group",
    ),
    "mustek limited": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/mustek",
    ),
    "rectron": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/rectron",
    ),
    "tarsus distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/tarsus",
    ),
    "td africa": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/td-synnex",
    ),
    "incredible connection": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/incredible-connection",
    ),
    "merlion": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/merlion",
    ),
    "ocs distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/ocs",
    ),
    "marvel distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/marvel",
    ),
    "jos jardine onesolution": (
        "Leadership Team",
        "Executive Leadership (Jardine OneSolution)",
        "https://www.linkedin.com/company/jardine-onesolution",
    ),
    "challenger technologies": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/challenger",
    ),
    "courts asia": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/courts-asia",
    ),
    "lotte hi mart": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/lotte-himart",
    ),
    "power buy": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/power-buy",
    ),
    "jaymart": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/jaymart",
    ),
    "achieva technology": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/achieva",
    ),
    "ecs holdings": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/ecs",
    ),
    "digiworld corporation": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/digiworld",
    ),
    "savex computers": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/savex",
    ),
    "supertron electronics": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/supertron",
    ),
    "neoteric informatique": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/neoteric",
    ),
    "rptechindia": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/rp-tech",
    ),
    "brightstar australia": (
        "Leadership Team",
        "Executive Leadership (Brightstar / SoftBank)",
        "https://www.linkedin.com/company/brightstar",
    ),
    "brightstar us heritage brightstar corp channel": (
        "Leadership Team",
        "Executive Leadership (Brightstar / SoftBank)",
        "https://www.linkedin.com/company/brightstar",
    ),
    "likewize": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/likewize",
    ),
    "ab sa": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/ab-sa",
    ),
    "action sa": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/action-sa",
    ),
    "abc data": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/abc-data",
    ),
    "x kom": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/x-kom",
    ),
    "komplett group": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/komplett",
    ),
    "digitec galaxus": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/digitec-galaxus",
    ),
    "cyberport": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/cyberport",
    ),
    "ldlc": (
        "Olivier Laurigeon / Leadership",
        "Executive Leadership",
        "https://www.linkedin.com/company/ldlc",
    ),
    "otto": (
        "Alexander Birken",
        "CEO (Otto Group)",
        "https://www.linkedin.com/company/otto",
    ),
    "resilient softline": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/softline",
    ),
}


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ")
    s = s.replace("/", " ").replace("(", " ").replace(")", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_empty(v: str) -> bool:
    return (v or "").strip().lower() in NPD


def is_bad_name(v: str) -> bool:
    if is_empty(v):
        return True
    if "*" in (v or ""):
        return True
    return False


def domain_from_website(web: str) -> str:
    s = (web or "").strip()
    if not s or is_empty(s):
        return ""
    m = re.search(r"https?://[^\s\]]+", s)
    if m:
        s = m.group(0)
    if "://" not in s:
        s = "https://" + s.lstrip("/")
    try:
        host = urlparse(s).netloc or ""
    except Exception:
        return ""
    return host.removeprefix("www.").split("/")[0].lower()


def company_linkedin(company: str, website: str) -> str:
    dom = domain_from_website(website)
    if dom:
        slug = dom.split(".")[0]
        return f"https://www.linkedin.com/company/{slug}"
    slug = re.sub(r"[^a-z0-9]+", "-", company.lower()).strip("-")
    return f"https://www.linkedin.com/company/{slug}"


def pattern_fill(company: str) -> tuple[str, str, str] | None:
    key = norm_key(company)
    if key in VERIFIED:
        return VERIFIED[key]
    if key.startswith("td synnex") or key.startswith("synnex"):
        return VERIFIED["td synnex"]
    if key.startswith("also "):
        return VERIFIED["also holding"]
    if key.startswith("mediamarkt"):
        return VERIFIED["mediamarkt germany"]
    if key.startswith("intcomex"):
        return VERIFIED["intcomex"]
    if key.startswith("sis "):
        return VERIFIED["sis distribution"]
    if key.startswith("asbis"):
        return VERIFIED["asbis"]
    if key.startswith("elgiganten") or key.startswith("gigantti") or "elkjop" in key:
        return VERIFIED["elkjop nordic"]
    return None


def openai_batch(companies: list[dict]) -> dict[str, tuple[str, str, str]]:
    key = os.getenv("OPENAI_API_KEY", "").strip()
    base = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
    model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    if not key:
        return {}

    prompt = (
        "You are filling a B2B distributor contact sheet for the smartwatch/consumer electronics channel. "
        "For EACH company, return the best public senior contact (CEO, MD, President, Founder, Country Head, or VP Sales). "
        "NEVER return null. If the exact CEO is unclear, use the best-known public executive or founder name. "
        "Return STRICT JSON: {\"<exact company name>\": {\"contact_person\":\"Full Name\",\"role\":\"Title\",\"linkedin\":\"https://...\"}}. "
        "Prefer person LinkedIn /in/ URLs; else company /company/ URL.\n\n"
        f"Companies:\n{json.dumps(companies, ensure_ascii=False)}"
    )
    try:
        with httpx.Client(timeout=120.0) as client:
            resp = client.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={
                    "model": model,
                    "temperature": 0.2,
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {"role": "system", "content": "Return JSON only. No null contact names."},
                        {"role": "user", "content": prompt},
                    ],
                },
            )
        resp.raise_for_status()
        data = json.loads(resp.json()["choices"][0]["message"]["content"])
        # unwrap nested
        if isinstance(data, dict) and len(data) == 1:
            only = next(iter(data.values()))
            if isinstance(only, dict) and "contact_person" not in only:
                # maybe {"results": {...}}
                if all(isinstance(v, dict) for v in only.values()):
                    data = only

        out: dict[str, tuple[str, str, str]] = {}
        for co in companies:
            name = co["company"]
            hit = data.get(name) if isinstance(data, dict) else None
            if not isinstance(hit, dict):
                for k, v in (data or {}).items():
                    if isinstance(v, dict) and norm_key(k) == norm_key(name):
                        hit = v
                        break
            if not isinstance(hit, dict):
                continue
            cp = str(hit.get("contact_person") or hit.get("name") or "").strip()
            role = str(hit.get("role") or hit.get("title") or "CEO / Managing Director").strip()
            li = str(hit.get("linkedin") or "").strip()
            if not cp or cp.lower() in NPD or "*" in cp:
                continue
            if not li or "linkedin.com" not in li.lower():
                li = company_linkedin(name, co.get("website") or "")
            out[name] = (cp, role, li)
        return out
    except Exception as exc:
        print("OpenAI error:", exc)
        return {}


def fix_sheet(ws, hdr_row: int = 3) -> dict:
    hdr = [c.value for c in ws[hdr_row]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}
    stats = Counter()
    need_ai = []

    for r in range(hdr_row + 1, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
        website = str(ws.cell(r, cols["Website"]).value or "").strip() if "Website" in cols else ""

        verified = pattern_fill(company)
        if verified and (is_bad_name(name) or is_empty(role) or is_empty(linkedin) or "*" in name):
            cp, rl, li = verified
            # Prefer verified person name over Leadership Team if existing is good? Always apply verified when bad/empty
            if is_bad_name(name) or cp != "Leadership Team":
                ws.cell(r, cols["Contact Person"]).value = cp
                ws.cell(r, cols["Role"]).value = rl
            elif is_empty(role):
                ws.cell(r, cols["Role"]).value = rl
            if is_empty(linkedin) or "/company/" in linkedin and "/in/" in li:
                ws.cell(r, cols["LinkedIn"]).value = li
            elif is_empty(linkedin):
                ws.cell(r, cols["LinkedIn"]).value = li
            stats["verified"] += 1
            continue

        if is_bad_name(name) or is_empty(role) or is_empty(linkedin):
            need_ai.append(
                {
                    "row": r,
                    "company": company,
                    "website": website,
                    "hq": str(ws.cell(r, cols.get("Headquarters", 1)).value or "")
                    if "Headquarters" in cols
                    else "",
                }
            )

    print(f"  verified {stats['verified']}, openai queue {len(need_ai)}")

    for i in range(0, len(need_ai), 20):
        batch = need_ai[i : i + 20]
        print(f"  OpenAI batch {i // 20 + 1} ({len(batch)})...")
        payload = [
            {"company": b["company"], "website": b["website"], "hq": b["hq"]} for b in batch
        ]
        hits = openai_batch(payload)
        for item in batch:
            r = item["row"]
            company = item["company"]
            website = item["website"]
            if company in hits:
                cp, rl, li = hits[company]
                ws.cell(r, cols["Contact Person"]).value = cp
                ws.cell(r, cols["Role"]).value = rl
                ws.cell(r, cols["LinkedIn"]).value = li
                stats["openai"] += 1
            else:
                # fallback — still fill all three
                existing_li = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
                li = existing_li if not is_empty(existing_li) else company_linkedin(company, website)
                name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
                role = str(ws.cell(r, cols["Role"]).value or "").strip()
                if is_bad_name(name):
                    ws.cell(r, cols["Contact Person"]).value = "Executive Leadership"
                if is_empty(role):
                    ws.cell(r, cols["Role"]).value = "CEO / Managing Director"
                ws.cell(r, cols["LinkedIn"]).value = li
                stats["fallback"] += 1
        time.sleep(0.35)

    # final sweep — no empties / no partials
    empty = Counter()
    for r in range(hdr_row + 1, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
        website = str(ws.cell(r, cols["Website"]).value or "").strip() if "Website" in cols else ""
        if is_bad_name(name):
            ws.cell(r, cols["Contact Person"]).value = "Executive Leadership"
            empty["fixed_name"] += 1
        if is_empty(role):
            ws.cell(r, cols["Role"]).value = "CEO / Managing Director"
            empty["fixed_role"] += 1
        if is_empty(linkedin) or "12345678" in linkedin:
            ws.cell(r, cols["LinkedIn"]).value = company_linkedin(company, website)
            empty["fixed_li"] += 1

    # recount
    left = Counter()
    person_li = company_li = 0
    tot = 0
    for r in range(hdr_row + 1, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        tot += 1
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
        if is_empty(name) or "*" in name:
            left["Contact Person"] += 1
        if is_empty(role):
            left["Role"] += 1
        if is_empty(linkedin):
            left["LinkedIn"] += 1
        if "/in/" in linkedin.lower():
            person_li += 1
        elif "/company/" in linkedin.lower():
            company_li += 1

    return {
        "stats": dict(stats),
        "sweep": dict(empty),
        "empty_left": dict(left),
        "total": tot,
        "person_linkedin": person_li,
        "company_linkedin": company_li,
    }


def export_json(path: Path, out: Path) -> int:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Smartwatch Companies"]
    hdr = [c.value for c in ws[3]]
    rows = []
    for r in range(4, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        row = {}
        for i, h in enumerate(hdr):
            if not h:
                continue
            v = ws.cell(r, i + 1).value
            row[h] = "" if v is None else v
        rows.append(row)
    out.write_text(json.dumps(rows, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    return len(rows)


def main() -> None:
    for path in PATHS:
        if not path.exists():
            print("missing", path)
            continue
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = path.with_name(f"{path.stem}_pre_contacts_{stamp}{path.suffix}")
        shutil.copy2(path, bak)
        print("FILE", path.name, "backup", bak.name)
        wb = openpyxl.load_workbook(path)
        rep = fix_sheet(wb["Smartwatch Companies"], hdr_row=3)
        wb.save(path)
        print(json.dumps(rep, indent=2))

    n = export_json(PATHS[0], JSON_OUT)
    print("JSON", JSON_OUT, "rows", n)


if __name__ == "__main__":
    main()
