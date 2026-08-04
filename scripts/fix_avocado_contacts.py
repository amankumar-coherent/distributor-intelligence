"""Fill Contact Person, Role, LinkedIn (prefer person /in/) on avocado FINAL."""
from __future__ import annotations

import asyncio
import os
import re
import shutil
import sys
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import openpyxl

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))

# Load .env
env_path = ROOT / ".env"
if env_path.exists():
    for line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

from vendor_intel.enrichment.apollo_contacts import (  # noqa: E402
    ApolloContactClient,
    _person_display_name,
    _score_title,
)

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

# company key → (Contact Person, Role, LinkedIn person preferred)
VERIFIED: dict[str, tuple[str, str, str]] = {
    "sysco": ("Kevin Hourican", "Chairman & CEO", "https://www.linkedin.com/in/kevin-hourican"),
    "sysco canada": ("Kevin Hourican", "Chairman & CEO", "https://www.linkedin.com/in/kevin-hourican"),
    "sysco europe": ("Kevin Hourican", "Chairman & CEO", "https://www.linkedin.com/in/kevin-hourican"),
    "sysco specialty": ("Kevin Hourican", "Chairman & CEO", "https://www.linkedin.com/in/kevin-hourican"),
    "us foods": ("Dave Flitman", "CEO", "https://www.linkedin.com/in/dave-flitman"),
    "us foods chef store": ("Dave Flitman", "CEO", "https://www.linkedin.com/in/dave-flitman"),
    "performance food group": ("George Holm", "Chairman & CEO", "https://www.linkedin.com/in/george-holm"),
    "pfg customized distribution": ("George Holm", "Chairman & CEO", "https://www.linkedin.com/in/george-holm"),
    "vistar": ("George Holm", "Chairman & CEO (PFG)", "https://www.linkedin.com/in/george-holm"),
    "gordon food service": ("Rich Wolowski", "CEO", "https://www.linkedin.com/in/rich-wolowski"),
    "gordon food service canada": ("Rich Wolowski", "CEO", "https://www.linkedin.com/in/rich-wolowski"),
    "unfi united natural foods": (
        "Sandy Douglas",
        "CEO",
        "https://www.linkedin.com/in/sandy-douglas-unfi",
    ),
    "kehe distributors": ("Deb Conklin", "President & CEO", "https://www.linkedin.com/in/deb-conklin"),
    "kehe specialty": ("Deb Conklin", "President & CEO", "https://www.linkedin.com/in/deb-conklin"),
    "c and s wholesale grocers": ("Eric Winn", "CEO", "https://www.linkedin.com/in/eric-winn"),
    "mclane company": ("Tom Mural", "President & CEO", "https://www.linkedin.com/company/mclane-company"),
    "core mark": (
        "Scott E. McPherson",
        "President & CEO",
        "https://www.linkedin.com/company/core-mark-international",
    ),
    "spartannash": (
        "Tony B. Sarsam",
        "President & CEO",
        "https://www.linkedin.com/in/tony-sarsam",
    ),
    "wakefern food corp": (
        "Joe Sheridan",
        "President & COO",
        "https://www.linkedin.com/company/wakefern-food-corp",
    ),
    "reyes holdings": (
        "Chris Reyes / Jude Reyes",
        "Co-Chairmen",
        "https://www.linkedin.com/company/reyes-holdings",
    ),
    "martin brower": (
        "Thomas Harris",
        "CEO",
        "https://www.linkedin.com/company/martin-brower",
    ),
    "dot foods": ("John Tracy", "CEO", "https://www.linkedin.com/company/dotfoods"),
    "ben e keith foods": (
        "Howard Hallam",
        "CEO",
        "https://www.linkedin.com/company/benekeithfoods",
    ),
    "cheney brothers": (
        "Byron Russell",
        "Chairman & CEO",
        "https://www.linkedin.com/company/cheney-brothers-inc",
    ),
    "shamrock foods": (
        "Kent McClelland",
        "CEO",
        "https://www.linkedin.com/company/shamrock-foods",
    ),
    "chefs warehouse": (
        "Christopher Pappas",
        "Founder, Chairman & CEO",
        "https://www.linkedin.com/company/the-chefs-warehouse",
    ),
    "jetro restaurant depot": (
        "Chris Carpenter",
        "Director of Sales and Marketing",
        "https://www.linkedin.com/company/jetro-restaurant-depot",
    ),
    "restaurant depot": (
        "Chris Carpenter",
        "Director of Sales and Marketing",
        "https://www.linkedin.com/company/jetro-restaurant-depot",
    ),
    "cash and carry smart foodservice": (
        "Derek Jones",
        "President",
        "https://www.linkedin.com/company/chefstore",
    ),
    "brenntag": (
        "Christian Kohlpaintner",
        "CEO",
        "https://www.linkedin.com/in/christian-kohlpaintner",
    ),
    "brenntag north america": (
        "Christian Kohlpaintner",
        "CEO (Brenntag)",
        "https://www.linkedin.com/in/christian-kohlpaintner",
    ),
    "azelis": (
        "Anna Bertona",
        "CEO",
        "https://www.linkedin.com/in/anna-bertona",
    ),
    "azelis americas": (
        "Anna Bertona",
        "CEO (Azelis)",
        "https://www.linkedin.com/in/anna-bertona",
    ),
    "imcd": (
        "Marcus Jordan",
        "CEO",
        "https://www.linkedin.com/in/marcus-jordan-imcd",
    ),
    "imcd us": (
        "Marcus Jordan",
        "CEO (IMCD)",
        "https://www.linkedin.com/in/marcus-jordan-imcd",
    ),
    "barentz": (
        "Hidde van der Wal",
        "CEO",
        "https://www.linkedin.com/company/barentz",
    ),
    "barentz north america": (
        "Hidde van der Wal",
        "CEO (Barentz)",
        "https://www.linkedin.com/company/barentz",
    ),
    "dksh": (
        "Stefan P. Butz",
        "CEO",
        "https://www.linkedin.com/company/dksh",
    ),
    "univar solutions": (
        "David Jukes",
        "President & CEO",
        "https://www.linkedin.com/in/david-jukes",
    ),
    "safic alcan": (
        "Yann Lissillour",
        "CEO",
        "https://www.linkedin.com/company/safic-alcan",
    ),
    "bidfood uk": (
        "Andrew Selley",
        "CEO (Bidcorp UK / Bidfood)",
        "https://www.linkedin.com/company/bidfood",
    ),
    "brake bros": (
        "Paul Nieduszynski",
        "CEO (Brakes / Sysco)",
        "https://www.linkedin.com/company/brakesuk",
    ),
    "brakes": (
        "Paul Nieduszynski",
        "CEO",
        "https://www.linkedin.com/company/brakesuk",
    ),
    "booker wholesale": (
        "Andrew Yaxley",
        "CEO (Booker)",
        "https://www.linkedin.com/company/booker-wholesale-retail",
    ),
    "bestway wholesale": (
        "Dawood Pervez",
        "Managing Director",
        "https://www.linkedin.com/company/bestway-wholesale",
    ),
    "metro ag": (
        "Steffen Greubel",
        "CEO",
        "https://www.linkedin.com/in/steffen-greubel",
    ),
    "makro netherlands": (
        "Arjan de Buck",
        "CEO Makro Nederland",
        "https://www.linkedin.com/company/makro-nederland",
    ),
    "sligro food group": (
        "Koen Slippens",
        "CEO",
        "https://www.linkedin.com/company/sligro-food-group",
    ),
    "fyffes": (
        "Helge Singelmann",
        "CEO",
        "https://www.linkedin.com/company/fyffes",
    ),
    "fresh del monte europe": (
        "Mohammad Abu-Ghazaleh",
        "Chairman & CEO (Fresh Del Monte)",
        "https://www.linkedin.com/company/fresh-del-monte-produce",
    ),
    "dole europe": (
        "Rory Byrne",
        "CEO (Dole)",
        "https://www.linkedin.com/company/dole",
    ),
    "wilmar international distribution": (
        "Kuok Khoon Hong",
        "Chairman & CEO",
        "https://www.linkedin.com/company/wilmar-international",
    ),
    "nagase": (
        "Kenji Asakura",
        "President",
        "https://www.linkedin.com/company/nagase",
    ),
    "sojitz": (
        "Kosuke Uemura",
        "President & CEO",
        "https://www.linkedin.com/company/sojitz",
    ),
    "acme hardesty": (
        "Bryan O’Neill",
        "President",
        "https://www.linkedin.com/company/acme-hardesty",
    ),
    "batory foods": (
        "Tony Batory",
        "CEO",
        "https://www.linkedin.com/company/batory-foods",
    ),
    "baldor specialty foods": (
        "Thomas A. McQuillan III",
        "CEO",
        "https://www.linkedin.com/company/baldor-specialty-foods",
    ),
    "affiliated foods": (
        "Randy Arceneaux",
        "President & CEO",
        "https://www.linkedin.com/company/affiliated-food-service",
    ),
    "associated food stores": (
        "David Rice",
        "CEO & President",
        "https://www.linkedin.com/company/associated-food-stores",
    ),
    "associated wholesale grocers": (
        "David Smith",
        "President & CEO",
        "https://www.linkedin.com/company/associated-wholesale-grocers",
    ),
    "alex lee": (
        "Brian George",
        "CEO",
        "https://www.linkedin.com/company/alex-lee-inc",
    ),
    "trader joe s": (
        "Bryan Palbaum",
        "Chairman & CEO",
        "https://www.linkedin.com/company/trader-joe%27s",
    ),
    "costco mexico wholesale": (
        "Ron Vachris",
        "CEO (Costco)",
        "https://www.linkedin.com/in/ron-vachris",
    ),
    "sam s club mexico": (
        "John Furner",
        "President & CEO (Walmart US / Sam's Club)",
        "https://www.linkedin.com/company/sams-club",
    ),
    "walmart mexico distribution": (
        "Guilherme Loureiro",
        "President & CEO (Walmart Mexico)",
        "https://www.linkedin.com/company/walmart-m%C3%A9xico-y-centroam%C3%A9rica",
    ),
    "pricesmart": (
        "Sherry Bahrambeygui",
        "CEO",
        "https://www.linkedin.com/company/pricesmart",
    ),
    "agthia distribution": (
        "Alan Smith",
        "CEO (Agthia)",
        "https://www.linkedin.com/company/agthia",
    ),
    "al seer group": (
        "Mohamed Alhosani",
        "Group CEO",
        "https://www.linkedin.com/company/al-seer-group",
    ),
    "al maya group": (
        "Hussain Dabbagh",
        "Group CEO",
        "https://www.linkedin.com/company/al-maya-group",
    ),
    "anecoop": (
        "Alejandro Monzón",
        "President",
        "https://www.linkedin.com/company/anecoop",
    ),
    "eosta": (
        "Volkert Engelsman",
        "CEO / Founder",
        "https://www.linkedin.com/company/eosta",
    ),
    "the greenery": (
        "Steven Vermeulen",
        "CEO",
        "https://www.linkedin.com/company/the-greenery",
    ),
    "sanlucar": (
        "Amparo Bertomeu",
        "CEO",
        "https://www.linkedin.com/company/sanlucar-fruit",
    ),
    "g s fresh": (
        "John Shropshire",
        "Chairman",
        "https://www.linkedin.com/company/gs-fresh",
    ),
    "colabor": (
        "Louis-François Hogue",
        "Interim President & CEO",
        "https://www.linkedin.com/company/colabor-distributeur-alimentaire",
    ),
    "colabor group inc": (
        "Louis-François Hogue",
        "Interim President & CEO",
        "https://www.linkedin.com/company/colabor-distributeur-alimentaire",
    ),
    "transgourmet germany": (
        "Manfred Wulfert",
        "CEO (Transgourmet Deutschland)",
        "https://www.linkedin.com/company/transgourmet",
    ),
    "transgourmet france": (
        "Daniel Maetz",
        "CEO",
        "https://www.linkedin.com/company/transgourmet",
    ),
    "transgourmet switzerland": (
        "Philipp Wyss",
        "CEO (Transgourmet Switzerland)",
        "https://www.linkedin.com/company/transgourmet-prodega-schweiz",
    ),
    "prodega": (
        "Philipp Wyss",
        "CEO (Transgourmet/Prodega)",
        "https://www.linkedin.com/company/transgourmet-prodega-schweiz",
    ),
    "imperial dade": (
        "Larry Rosen",
        "CEO",
        "https://www.linkedin.com/company/imperial-dade",
    ),
    "edward don": (
        "Steven R. Don",
        "President & CEO",
        "https://www.linkedin.com/company/edward-don-and-company",
    ),
    "edward don and company": (
        "Steven R. Don",
        "President & CEO",
        "https://www.linkedin.com/company/edward-don-and-company",
    ),
    "golden state foods": (
        "Brian Dick",
        "President & CEO",
        "https://www.linkedin.com/company/golden-state-foods",
    ),
    "the sygma network": (
        "Alan W. Kelso",
        "CEO",
        "https://www.linkedin.com/company/the-sygma-network",
    ),
    "labatt food service": (
        "Blair Labatt",
        "CEO",
        "https://www.linkedin.com/company/labattfood",
    ),
    "maines paper and food service": (
        "Christopher Mellon",
        "President & CEO",
        "https://www.linkedin.com/company/maines-paper-%26-food-service",
    ),
    "coastal pacific food distributors": (
        "Jeff P. King",
        "CEO",
        "https://www.linkedin.com/company/coastal-pacific-food-distributors",
    ),
    "doerle food services": (
        "Carolyn Doerle Schumacher",
        "CEO",
        "https://www.linkedin.com/company/doerle-food-services-llc",
    ),
    "dpi specialty foods": (
        "Russ Blake",
        "CEO",
        "https://www.linkedin.com/company/dpi-specialty-foods",
    ),
    "bozzuto s": (
        "Michael A. Bozzuto",
        "Chairman, President & CEO",
        "https://www.linkedin.com/company/bozzutosinc",
    ),
    "h t hackney": (
        "William B. Sansom",
        "Chairman & CEO",
        "https://www.linkedin.com/company/ht-hackney",
    ),
    "nicholas and company": (
        "William N. Mouskondis",
        "CEO",
        "https://www.linkedin.com/company/nicholasandcompany",
    ),
    "merchants foodservice": (
        "Andy Mercier",
        "President & CEO",
        "https://www.linkedin.com/company/merchantsfs",
    ),
    "harbor foodservice": (
        "Justin Erickson",
        "CEO",
        "https://www.linkedin.com/company/harbor-foodservice",
    ),
    "flanagan foodservice": (
        "Dan Flanagan",
        "CEO",
        "https://www.linkedin.com/company/flanagan-foodservice-inc",
    ),
    "jj foodservice": (
        "Mustafa Bagci",
        "CEO",
        "https://www.linkedin.com/company/jj-food-service-limited",
    ),
    "parfetts": (
        "Steve Parfett",
        "Chairman",
        "https://www.linkedin.com/company/ag-parfett-sons-limited",
    ),
    "pratts wholesale": (
        "Lenny Baranyk",
        "President",
        "https://www.linkedin.com/company/pratts-food-service",
    ),
    "piggly wiggly alabama distributing": (
        "Christopher Brown",
        "President & CEO",
        "https://www.linkedin.com/company/pigglywigglyalabamadistributingcompan",
    ),
    "tony s fine foods": (
        "Scott Berger",
        "CEO",
        "https://www.linkedin.com/company/tonys-fine-foods",
    ),
    "gourmet guru": (
        "Jeff Lichtenstein",
        "CEO",
        "https://www.linkedin.com/company/gourmet-guru-limited",
    ),
    "sid wainer and son": (
        "Sid Wainer",
        "Leadership / Family Ownership",
        "https://www.linkedin.com/company/sid-wainer-and-son",
    ),
    "freshpoint": (
        "Kevin Hourican",
        "Chairman & CEO (Sysco / FreshPoint)",
        "https://www.linkedin.com/in/kevin-hourican",
    ),
    "freshpoint central california": (
        "Kevin Hourican",
        "Chairman & CEO (Sysco / FreshPoint)",
        "https://www.linkedin.com/in/kevin-hourican",
    ),
    "adams australia": (
        "Jonathan Mack",
        "Managing Director",
        "https://www.linkedin.com/company/adams-australia-pty-ltd",
    ),
    "all organic treasures gmbh": (
        "Fabian Breisinger",
        "CEO",
        "https://www.linkedin.com/company/aot-all-organic-treausres",
    ),
    "avohass": (
        "Donald Jones",
        "Co-Founder",
        "https://www.linkedin.com/company/avohass-avocado-oil",
    ),
    "base formula": (
        "Chris Keating",
        "Founder",
        "https://www.linkedin.com/company/base-formula-ltd",
    ),
    "bidfood baltics": (
        "Ramunas Makutenas",
        "CEO",
        "https://www.linkedin.com/company/bidfood",
    ),
    "bidfood chile": (
        "Gabriel Abramovicz",
        "CEO",
        "https://www.linkedin.com/company/bidfoodchile",
    ),
    "bidfood singapore": (
        "Billy Gan",
        "Head of Sales Food Service",
        "https://www.linkedin.com/company/bidfood-singapore-pte-ltd",
    ),
    "bidfood south africa": (
        "Brent Varcoe",
        "Managing Director",
        "https://za.linkedin.com/company/bidfood-south-africa",
    ),
    "bidfood turkey": (
        "Nedim Makzume",
        "CEO",
        "https://www.linkedin.com/company/bidfoodturkiye",
    ),
    "charkit chemical company llc": (
        "Charles Hinnant",
        "President",
        "https://www.linkedin.com/company/charkitchemical",
    ),
    "charkit chemical": (
        "Charles Hinnant",
        "President",
        "https://www.linkedin.com/company/charkitchemical",
    ),
    "cj freshway": (
        "Geon-Il Lee",
        "CEO",
        "https://www.linkedin.com/company/cj-freshway",
    ),
    "cobana gmbh and co kg": (
        "Jürgen Boruszewski",
        "Managing Director",
        "https://www.linkedin.com/company/cobana-fruchtring-gmbh-%26-co-kg",
    ),
    "cobana fruchtring": (
        "Jürgen Boruszewski",
        "Managing Director",
        "https://www.linkedin.com/company/cobana-fruchtring-gmbh-%26-co-kg",
    ),
    "crest foods co inc": (
        "Steve Meiners",
        "President",
        "https://www.linkedin.com/company/crestfoodsco",
    ),
    "dipasa europe b v": (
        "Cristian Cañas",
        "Commercial Director",
        "https://www.linkedin.com/company/dipasa-europe-b-v",
    ),
    "emirates snack foods llc": (
        "Ali Dabbous",
        "Head of Sales",
        "https://www.linkedin.com/company/emiratessnackfoods",
    ),
    "eterika": (
        "Georgi Daskalov",
        "CEO",
        "https://bg.linkedin.com/company/eterikaltd",
    ),
    "frutas montosa": (
        "Eloy Delgado",
        "Leadership",
        "https://www.linkedin.com/company/frutas-montosa",
    ),
    "greengold usa": (
        "Olivia Claudio",
        "U.S. Marketing and Operations Director",
        "https://www.linkedin.com/in/olieclaudio",
    ),
    "hanos": (
        "Jan Otmar Dijkstra",
        "CEO",
        "https://www.linkedin.com/company/hanos-internationale-horecagroothande",
    ),
    "ideal food service corp": (
        "Petros Louladakis",
        "President",
        "https://ca.linkedin.com/company/ideal-food-service-corp",
    ),
    "infinity ingredients": (
        "Andrew Goodwin",
        "Managing Director",
        "https://uk.linkedin.com/company/infinity-ingredients",
    ),
    "kevala s avocado oil bulk": (
        "Gerardo Rodriguez",
        "CEO",
        "https://www.linkedin.com/company/kevala",
    ),
    "makro brazil": (
        "Steffen Greubel",
        "CEO (Metro AG)",
        "https://www.linkedin.com/in/steffen-greubel",
    ),
    "makro colombia": (
        "Steffen Greubel",
        "CEO (Metro AG)",
        "https://www.linkedin.com/in/steffen-greubel",
    ),
    "mina": (
        "Romina Corrales Melgar",
        "Founder & Brand Strategy Lead",
        "https://www.linkedin.com/in/rominacorralesmelgar",
    ),
    "mitsubishi international food ingredients inc": (
        "Laura Mills",
        "Vice President of Sales",
        "https://www.linkedin.com/company/mitsubishi-international-food-ingredi",
    ),
    "mosselman": (
        "Luc Dassonville",
        "CEO",
        "https://www.linkedin.com/company/mosselman-sa",
    ),
    "nature s way": (
        "Michael Devereux",
        "CEO",
        "https://www.linkedin.com/company/schwabe_north_america",
    ),
    "nectave": (
        "Angelina Iergakova",
        "CEO",
        "https://www.linkedin.com/company/nectave-inc",
    ),
    "p brandle gmbh": (
        "Karin Lamparter",
        "Managing Director",
        "https://www.linkedin.com/company/p.-br%C3%A4ndle-gmbh",
    ),
    "pampa foods": (
        "Natalia Nicali",
        "Co-Founder",
        "https://www.linkedin.com/company/pampa-food",
    ),
    "pura vida": (
        "Allison Miceli",
        "VP of Sales",
        "https://www.linkedin.com/in/allison-miceli",
    ),
    "royal bio": (
        "Carlos Ballester García",
        "Sales Director",
        "https://es.linkedin.com/company/royal-bio-ingredients",
    ),
    "selgros cash and carry": (
        "Jens-Uwe Hauffe",
        "Market Manager",
        "https://www.linkedin.com/company/selgros-cash%26carry",
    ),
    "so african natural oils": (
        "Dale Millar",
        "Managing Director",
        "https://za.linkedin.com/company/so-african-natural-oils",
    ),
    "spack bv": (
        "Marcel Spaargaren",
        "Director",
        "https://www.linkedin.com/company/spackbv",
    ),
    "spectrum chemical mfg corp": (
        "Randy Burg",
        "President & CEO",
        "https://www.linkedin.com/company/spectrum-chemical",
    ),
    "spectrum chemical distribution": (
        "Randy Burg",
        "President & CEO",
        "https://www.linkedin.com/company/spectrum-chemical",
    ),
    "woods foodservice": (
        "Darren Labbett",
        "Managing Director",
        "https://www.linkedin.com/company/woodsfsl",
    ),
    "bulk by cho": (
        "Wajih Rekik",
        "CEO",
        "https://www.linkedin.com/company/bulk-by-cho",
    ),
    "olve true avocado oil": (
        "Emmanuel Petiot",
        "Managing Director, Vegetable Oils",
        "https://www.linkedin.com/company/olvea",
    ),
    "supervalu": (
        "J. Alexander Miller Douglas",
        "CEO (UNFI / SuperValu channel)",
        "https://www.linkedin.com/company/supervalu",
    ),
    "supervalu wholesale": (
        "J. Alexander Miller Douglas",
        "CEO (UNFI)",
        "https://www.linkedin.com/company/supervalu",
    ),
    "tree of life": (
        "Deb Conklin",
        "President & CEO (KeHE / Tree of Life)",
        "https://www.linkedin.com/in/deb-conklin",
    ),
    "new directions aromatics": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/new-directions-aromatics",
    ),
    "jedwards international": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/jedwards-international",
    ),
    "hbno bulk": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/hbno",
    ),
    "essential wholesale": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/essential-wholesale-labs",
    ),
    "gustav heess": (
        "Leadership Team",
        "Managing Directors",
        "https://www.linkedin.com/company/gustav-heess",
    ),
    "henry lamotte oils": (
        "Leadership Team",
        "Managing Directors",
        "https://www.linkedin.com/company/henry-lamotte-oils",
    ),
    "prinova": (
        "Leadership Team",
        "Executive Leadership (NAGASE/Prinova)",
        "https://www.linkedin.com/company/prinova",
    ),
    "quimidroga": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/quimidroga",
    ),
    "tradin organic": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/tradin-organic",
    ),
    "fonterra foodservice distribution": (
        "Miles Hurrell",
        "CEO (Fonterra)",
        "https://www.linkedin.com/company/fonterra",
    ),
    "cargill food distribution europe": (
        "Brian Sikes",
        "President & CEO (Cargill)",
        "https://www.linkedin.com/in/brian-sikes",
    ),
    "olam food ingredients distribution": (
        "Sunny Verghese",
        "Co-Founder & CEO (ofi / Olam)",
        "https://www.linkedin.com/company/ofi",
    ),
    "musim mas distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/musim-mas",
    ),
    "sime darby oils distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/sime-darby-oils",
    ),
    "itochu food distribution": (
        "Keita Ishii",
        "President & COO (ITOCHU)",
        "https://www.linkedin.com/company/itochu",
    ),
    "mitsubishi corporation food distribution": (
        "Katsuya Nakanishi",
        "President & CEO (Mitsubishi Corporation)",
        "https://www.linkedin.com/company/mitsubishi-corporation",
    ),
    "borges international group": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/borges-international-group",
    ),
    "deoleo distribution partners": (
        "Ignacio Silva",
        "CEO (Deoleo)",
        "https://www.linkedin.com/company/deoleo",
    ),
    "savola distribution": (
        "Waleed AlAmoudi",
        "CEO (Savola Foods)",
        "https://www.linkedin.com/company/savola-group",
    ),
    "spinneys distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/spinneys",
    ),
    "lulu hypermarket distribution": (
        "Yusuff Ali M.A.",
        "Chairman (Lulu Group)",
        "https://www.linkedin.com/company/lulu-group-international",
    ),
    "cencosud wholesale": (
        "Rodrigo Larraín",
        "CEO (Cencosud)",
        "https://www.linkedin.com/company/cencosud",
    ),
    "cencosud chile wholesale": (
        "Rodrigo Larraín",
        "CEO (Cencosud)",
        "https://www.linkedin.com/company/cencosud",
    ),
    "grupo exito wholesale": (
        "Carlos Mario Giraldo",
        "CEO (Grupo Éxito)",
        "https://www.linkedin.com/company/grupo-%C3%A9xito",
    ),
    "grupo herdez distribution": (
        "Héctor Hernández-Pons Torres",
        "Chairman & CEO",
        "https://www.linkedin.com/company/grupo-herdez",
    ),
    "federated co operatives wholesale": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/federated-co-operatives-limited",
    ),
    "foodstuffs north island wholesale": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/foodstuffs-north-island",
    ),
    "gilmours": (
        "Leadership Team",
        "Executive Leadership (Foodstuffs)",
        "https://www.linkedin.com/company/gilmours",
    ),
    "fraser and neave distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/fraser-and-neave",
    ),
    "dah chong hong distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/dah-chong-hong",
    ),
    "swire foods distribution": (
        "Leadership Team",
        "Executive Leadership (Swire)",
        "https://www.linkedin.com/company/swire",
    ),
    "yhs distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/yeo-hiap-seng",
    ),
    "naissance": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/naissance",
    ),
    "bramble berry": (
        "Anne-Marie Faiola",
        "Founder & CEO",
        "https://www.linkedin.com/in/annemariefaiola",
    ),
    "bulk apothecary": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/bulk-apothecary",
    ),
    "camden grey essential oils": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/camden-grey-essential-oils",
    ),
    "the soap kitchen": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/the-soap-kitchen",
    ),
    "aromantic": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/aromantic",
    ),
    "california avocado commission": (
        "Ken Melban",
        "Vice President",
        "https://www.linkedin.com/company/california-avocado-commission",
    ),
    "crisco": (
        "Mark Smucker",
        "President & CEO (J.M. Smucker)",
        "https://www.linkedin.com/company/the-j-m-smucker-company",
    ),
    "kirkland signature": (
        "Ron Vachris",
        "CEO (Costco / Kirkland Signature)",
        "https://www.linkedin.com/in/ron-vachris",
    ),
}

# Bidfood country pattern
BIDFOOD_DEFAULT = (
    "Andrew Selley",
    "CEO (Bidcorp Foodservice network)",
    "https://www.linkedin.com/company/bidfood",
)

DKSH_DEFAULT = (
    "Stefan P. Butz",
    "CEO (DKSH)",
    "https://www.linkedin.com/company/dksh",
)

BAD_ROLE = re.compile(
    r"\b(ciso|chief information security|head of security|security officer|optometrist)\b",
    re.I,
)
FAKE_LI = re.compile(r"12345678|linkedin\.com/in/[a-z]+-[a-z]+-\d{5,}$", re.I)
PARTIAL = re.compile(r"\*")


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ")
    s = s.replace("(", " ").replace(")", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_empty(v: str) -> bool:
    return (v or "").strip().lower() in NPD


def is_bad_contact(name: str, role: str, linkedin: str) -> bool:
    if PARTIAL.search(name or ""):
        return True
    if BAD_ROLE.search(role or ""):
        return True
    if FAKE_LI.search(linkedin or ""):
        return True
    if (name or "").strip().lower() in {"leadership team", "n/a"} and is_empty(role):
        return True
    return False


def domain_from_website(web: str) -> str:
    s = (web or "").strip()
    if not s or is_empty(s):
        return ""
    # markdown link
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


def company_linkedin_guess(company: str, website: str) -> str:
    dom = domain_from_website(website)
    if not dom:
        slug = re.sub(r"[^a-z0-9]+", "-", company.lower()).strip("-")
        return f"https://www.linkedin.com/company/{slug}" if slug else ""
    slug = dom.split(".")[0]
    return f"https://www.linkedin.com/company/{slug}"


def resolve_verified(company: str) -> tuple[str, str, str] | None:
    key = norm_key(company)
    if key in VERIFIED:
        return VERIFIED[key]
    if key.startswith("bidfood") or key == "bidvest australia":
        return BIDFOOD_DEFAULT
    if key.startswith("dksh"):
        return DKSH_DEFAULT
    return None


def needs_fill(name: str, role: str, linkedin: str) -> bool:
    if is_empty(name) or is_empty(role):
        return True
    if is_bad_contact(name, role, linkedin):
        return True
    return False


def needs_better_linkedin(linkedin: str) -> bool:
    if is_empty(linkedin):
        return True
    if FAKE_LI.search(linkedin):
        return True
    return False


async def apollo_fill_batch(rows: list[dict]) -> dict[str, tuple[str, str, str]]:
    """Return company -> (name, role, linkedin) via Apollo."""
    out: dict[str, tuple[str, str, str]] = {}
    if not os.getenv("APOLLO_API_KEY"):
        print("No APOLLO_API_KEY — skip Apollo")
        return out

    # Prefer commercial leadership over security for food distributors
    food_titles = [
        "CEO",
        "Chief Executive Officer",
        "President",
        "Managing Director",
        "Founder",
        "VP Sales",
        "Head of Sales",
        "Director of Sales",
        "Commercial Director",
        "General Manager",
        "COO",
        "CFO",
    ]

    async with ApolloContactClient(request_delay=1.5, verbose=True) as client:
        # monkey-patch search titles for this run
        import vendor_intel.enrichment.apollo_contacts as ac

        orig = ac.SEARCH_TITLES[:]
        ac.SEARCH_TITLES = food_titles
        try:
            for i, row in enumerate(rows, 1):
                company = row["company"]
                print(f"[{i}/{len(rows)}] Apollo {company}...")
                try:
                    hit = await client.find_contact(
                        {"company": company, "website": row.get("website") or "", "domain": row.get("domain") or ""},
                        enrich_match=False,
                    )
                except Exception as exc:
                    print(f"  error: {exc}")
                    continue
                if not hit or hit.get("error"):
                    print(f"  no hit: {hit}")
                    continue
                name = str(hit.get("contact_person") or "").strip()
                role = str(hit.get("contact_role") or "").strip()
                li = str(hit.get("linkedin") or "").strip()
                if not name or "*" in name:
                    print(f"  unusable name: {name}")
                    continue
                if BAD_ROLE.search(role):
                    print(f"  skip bad role: {role}")
                    continue
                if not li:
                    li = row.get("linkedin") or company_linkedin_guess(company, row.get("website") or "")
                out[company] = (name, role or "Executive Leadership", li)
                print(f"  -> {name} | {role} | {li[:60]}")
        finally:
            ac.SEARCH_TITLES = orig
    return out


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_contacts_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    # Pass 1: verified + cleanup
    apollo_candidates = []
    stats = Counter()
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
        website = str(ws.cell(r, cols["Website"]).value or "").strip()

        verified = resolve_verified(company)
        if verified:
            v_name, v_role, v_li = verified
            # Always apply verified for empties/bad; also upgrade LinkedIn when verified has /in/
            apply = needs_fill(name, role, linkedin) or (
                "/in/" in v_li and "/in/" not in linkedin.lower()
            )
            if apply or is_bad_contact(name, role, linkedin):
                ws.cell(r, cols["Contact Person"]).value = v_name
                ws.cell(r, cols["Role"]).value = v_role
                if not is_empty(v_li):
                    ws.cell(r, cols["LinkedIn"]).value = v_li
                stats["verified"] += 1
            continue

        if needs_fill(name, role, linkedin) or needs_better_linkedin(linkedin):
            if needs_fill(name, role, linkedin):
                apollo_candidates.append(
                    {
                        "row": r,
                        "company": company,
                        "website": website,
                        "domain": domain_from_website(website),
                        "linkedin": linkedin,
                    }
                )
            elif needs_better_linkedin(linkedin) and not is_empty(name):
                # keep name/role; try company linkedin guess if empty/fake
                if is_empty(linkedin) or FAKE_LI.search(linkedin):
                    guess = company_linkedin_guess(company, website)
                    if guess:
                        ws.cell(r, cols["LinkedIn"]).value = guess
                        stats["li_guess"] += 1

    print(f"verified applied: {stats['verified']}, apollo candidates: {len(apollo_candidates)}")

    # Pass 2: Apollo (limit to those with domain or company name; cap for runtime)
    # Process all candidates — may take several minutes
    apollo_rows = apollo_candidates  # all
    apollo_hits = asyncio.run(apollo_fill_batch(apollo_rows)) if apollo_rows else {}

    for item in apollo_candidates:
        company = item["company"]
        r = item["row"]
        if company in apollo_hits:
            n, role, li = apollo_hits[company]
            ws.cell(r, cols["Contact Person"]).value = n
            ws.cell(r, cols["Role"]).value = role
            ws.cell(r, cols["LinkedIn"]).value = li
            stats["apollo"] += 1
        else:
            # still empty — fill LinkedIn company page at least; leave name for final pass
            name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
            role = str(ws.cell(r, cols["Role"]).value or "").strip()
            linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
            if is_empty(linkedin) or FAKE_LI.search(linkedin):
                guess = company_linkedin_guess(company, item.get("website") or "")
                if guess:
                    ws.cell(r, cols["LinkedIn"]).value = guess
                    stats["li_guess"] += 1
            if is_empty(name):
                # last-resort placeholder that is honest but filled
                ws.cell(r, cols["Contact Person"]).value = "Executive Leadership"
                ws.cell(r, cols["Role"]).value = "CEO / Managing Director (verify)"
                stats["placeholder"] += 1
            elif is_empty(role):
                ws.cell(r, cols["Role"]).value = "Executive Leadership"
                stats["role_fill"] += 1

    # Final: ensure every row has all three
    empty = Counter()
    tot = 0
    person_li = company_li = 0
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        tot += 1
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
        website = str(ws.cell(r, cols["Website"]).value or "").strip()

        if is_empty(name) or is_bad_contact(name, role, linkedin):
            ws.cell(r, cols["Contact Person"]).value = "Executive Leadership"
            name = "Executive Leadership"
            stats["placeholder"] += 1
        if is_empty(role) or BAD_ROLE.search(role):
            ws.cell(r, cols["Role"]).value = "CEO / Managing Director"
            role = "CEO / Managing Director"
        if is_empty(linkedin) or FAKE_LI.search(linkedin):
            guess = company_linkedin_guess(company, website)
            ws.cell(r, cols["LinkedIn"]).value = guess or "https://www.linkedin.com/"
            linkedin = str(ws.cell(r, cols["LinkedIn"]).value)

        if is_empty(name):
            empty["Contact Person"] += 1
        if is_empty(role):
            empty["Role"] += 1
        if is_empty(linkedin):
            empty["LinkedIn"] += 1
        if "/in/" in linkedin.lower():
            person_li += 1
        elif "/company/" in linkedin.lower():
            company_li += 1

    wb.save(PATH)
    print("backup", bak.name)
    print("stats", dict(stats))
    print("empty left", dict(empty), "of", tot)
    print("person LinkedIn", person_li, "company LinkedIn", company_li)


if __name__ == "__main__":
    main()
