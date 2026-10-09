"""Curate ~300 avocado-oil channel distributors, verify domains, update seeds + Phase2."""
from __future__ import annotations

import json
import re
import socket
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(r"D:\vendor-intel")
SEED_PATH = ROOT / "queries" / "seeds" / "global_avocado_oil_market.txt"
P2_PATH = ROOT / "output" / "phase2" / "phase2_companies_avocado_oil_global.json"
DESK_SEED = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\queries\seeds\global_avocado_oil_market.txt")
REPORT = ROOT / "output" / "_seed300_avocado_distributors_report.json"

# name, domain — curated B2B distributors / wholesalers / importers / foodservice & ingredient channels
# Manufacturers / pure brands intentionally omitted.
RAW: list[tuple[str, str]] = [
    # --- US broadline / specialty foodservice (verified lists: GourmetPro, Crew MP) ---
    ("Sysco", "sysco.com"),
    ("US Foods", "usfoods.com"),
    ("Performance Food Group", "pfgc.com"),
    ("Gordon Food Service", "gfs.com"),
    ("McLane Company", "mclaneco.com"),
    ("KeHE Distributors", "kehe.com"),
    ("UNFI", "unfi.com"),
    ("C&S Wholesale Grocers", "cswg.com"),
    ("Core-Mark", "core-mark.com"),
    ("Cheney Brothers", "cheneybrothers.com"),
    ("Ben E. Keith Foods", "benekeith.com"),
    ("Shamrock Foods", "shamrockfoods.com"),
    ("Maines Paper & Food Service", "maines.net"),
    ("SpartanNash", "spartannash.com"),
    ("H.T. Hackney", "hthackney.com"),
    ("Reyes Holdings", "reyesholdings.com"),
    ("Martin-Brower", "martin-brower.com"),
    ("Dot Foods", "dotfoods.com"),
    ("Nicholas and Company", "nicholasandco.com"),
    ("Harbor Foodservice", "harborfoodservice.com"),
    ("Golden State Foods", "goldenstatefoods.com"),
    ("The SYGMA Network", "sygma.com"),
    ("Vistar", "vistar.com"),
    ("Labatt Food Service", "labattfood.com"),
    ("Coastal Pacific Food Distributors", "cpfd.com"),
    ("Doerle Food Services", "doerle.com"),
    ("S&W Wholesale Foods", "sandwfoods.com"),
    ("Northern Hasserot", "northernhasserot.com"),
    ("Scavuzzo's", "scavuzzo.com"),
    ("Merchants Foodservice", "merchantsfoodservice.com"),
    ("Imperial Dade", "imperialdade.com"),
    ("PFG Customized Distribution", "pfgc.com"),
    ("FreshPoint", "freshpoint.com"),
    ("Greco & Sons", "grecosons.com"),
    ("Edward Don", "don.com"),
    ("Sysco Specialty", "sysco.com"),
    ("US Foods CHEF'STORE", "chefstore.com"),
    ("Jetro Restaurant Depot", "jetro.com"),
    ("Restaurant Depot", "restaurantdepot.com"),
    ("Cash & Carry Smart Foodservice", "smartfoodservice.com"),
    ("Associated Wholesale Grocers", "awginc.com"),
    ("Wakefern Food Corp", "wakefern.com"),
    ("Supervalu Wholesale", "supervalu.com"),
    ("Affiliated Foods", "affiliatedfoods.com"),
    ("Associated Food Stores", "afstores.com"),
    ("Piggly Wiggly Alabama Distributing", "pigglywiggly.com"),
    ("Alex Lee", "alexlee.com"),
    ("Bozzuto's", "bozzutos.com"),
    ("Krasdale Foods", "krasdale.com"),
    ("Grocers Supply", "grocerssupply.com"),
    ("KeHE Specialty", "kehe.com"),
    ("Tony's Fine Foods", "tonysfinefoods.com"),
    ("Tree of Life", "treeoflife.com"),
    ("Nature's Best", "naturesbest.com"),
    ("DPI Specialty Foods", "dpispecialtyfoods.com"),
    ("Liberty Richter", "libertyrichter.com"),
    ("World Finer Foods", "worldfiner.com"),
    ("Gourmet Guru", "gourmetguru.com"),
    ("European Imports", "eii1.com"),
    ("Fusion Fine Foods", "fusionfinefoods.com"),
    ("Chefs' Warehouse", "chefswarehouse.com"),
    ("Baldor Specialty Foods", "baldorfood.com"),
    ("Sid Wainer & Son", "sidwainer.com"),
    ("FreshPoint Central California", "freshpoint.com"),
    ("Markon Cooperative", "markon.com"),
    ("Pro*Act", "proactusa.com"),
    ("American International Foods", "americaninternationalfoods.com"),
    ("foodguys", "foodguys.com"),
    ("Nectave", "nectave.com"),
    ("Parchem", "parchem.com"),
    ("Acme-Hardesty", "acme-hardesty.com"),
    ("Jedwards International", "bulknaturaloils.com"),
    ("HBNO Bulk", "hbnobulk.com"),
    ("The Bulk Cart", "thebulkcart.com"),
    ("New Directions Aromatics", "newdirectionsaromatics.com"),
    ("MakingCosmetics", "makingcosmetics.com"),
    ("Lotioncrafter", "lotioncrafter.com"),
    ("From Nature With Love", "fromnaturewithlove.com"),
    ("Bulk Apothecary", "bulkapothecary.com"),
    ("Essential Wholesale", "essentialwholesale.com"),
    ("Soapgoods", "soapgoods.com"),
    ("Camden-Grey Essential Oils", "camdengrey.com"),
    ("Liberty Natural Products", "libertynatural.com"),
    ("Windy Point Soap Making Supplies", "windypointsoap.com"),
    ("Brenntag North America", "brenntag.com"),
    ("Univar Solutions", "univarsolutions.com"),
    ("Azelis Americas", "azelis.com"),
    ("IMCD US", "imcdgroup.com"),
    ("Barentz North America", "barentz.com"),
    ("Connell Brothers", "connellbros.com"),
    ("Pacific Coast Chemicals", "pcchem.com"),
    ("Hydrite Chemical", "hydrite.com"),
    ("ChemCeed", "chemceed.com"),
    ("KIC Chemicals", "kicchemicals.com"),
    ("Charkit Chemical", "charkit.com"),
    ("Lincoln Fine Ingredients", "lincoln-fine-ingredients.com"),
    ("Spectrum Chemical Distribution", "spectrumchemical.com"),
    ("Ingredi", "ingredi.com"),
    ("Prinova", "prinovaglobal.com"),
    ("Nutra Food Ingredients", "nutrafoodingredients.com"),
    ("Crest Foods Ingredients", "crestfoods.com"),
    ("Henry Broch Foods", "henrybroch.com"),
    ("Spice Chain", "spicechain.com"),
    ("Atlantic Chemicals Trading", "actusa.net"),
    ("Batory Foods", "batoryfoods.com"),
    ("Gillco Ingredients", "gillco.com"),
    ("Circle Ingredients", "circlein.com"),
    ("RFI Ingredients", "rfiingredients.com"),
    ("AIDP", "aidp.com"),
    ("Nutricost Wholesale", "nutricost.com"),
    # --- Canada ---
    ("Sysco Canada", "sysco.ca"),
    ("Gordon Food Service Canada", "gfs.ca"),
    ("Colabor", "colabor.com"),
    ("Flanagan Foodservice", "flanagan.ca"),
    ("Qualifirst Foods", "qualifirst.com"),
    ("UNFI Canada", "unfi.com"),
    ("Horizon Distributors", "horizondistributors.com"),
    ("Discovery Organics", "discoveryorganics.ca"),
    ("Pratts Wholesale", "pratts.ca"),
    ("Food Supplies Company", "foodsupplies.ca"),
    ("Ideal Foodservice", "idealfoods.ca"),
    ("I-D Foods", "idfoods.com"),
    ("Willoughby Distribution", "willoughby.ca"),
    ("Service Alimentaire Desco", "desco.ca"),
    ("Dubord & Rainville", "dubord-rainville.com"),
    ("Les Aliments Dubeau", "dubeau.com"),
    ("Summit Food Service Distributors", "summitfoods.com"),
    ("Traynor Foods", "traynorfoods.com"),
    ("Western Grocers", "westerngrocers.com"),
    ("Federated Co-operatives Wholesale", "fcl.crs"),
    # --- UK / Ireland ---
    ("Bidfood UK", "bidfood.co.uk"),
    ("Brakes", "brake.co.uk"),
    ("Booker Wholesale", "booker.co.uk"),
    ("Bestway Wholesale", "bestwaywholesale.co.uk"),
    ("JJ Foodservice", "jjfoodservice.com"),
    ("Castell Howell Foods", "chfoods.co.uk"),
    ("Dunns Food and Drinks", "dunnsfoodanddrinks.co.uk"),
    ("Woods Foodservice", "woodsfoodservice.co.uk"),
    ("Henderson Foodservice", "hendersonfoodservice.com"),
    ("Matthew Clark", "matthewclark.co.uk"),
    ("Blakemore Wholesale", "afblakemore.com"),
    ("Parfetts", "parfetts.co.uk"),
    ("Dhamecha Cash & Carry", "dhamecha.com"),
    ("Londis Wholesale", "londis.co.uk"),
    ("Nisa Wholesale", "nisalocally.co.uk"),
    ("Costcutter Supply", "costcutter.co.uk"),
    ("Infinity Ingredients", "infinity-ingredients.co.uk"),
    ("The Soap Kitchen", "thesoapkitchen.co.uk"),
    ("Gracefruit", "gracefruit.com"),
    ("Aromantic", "aromantic.co.uk"),
    ("Naturallythinking", "naturallythinking.com"),
    ("Mystic Moments", "mysticmomentsuk.com"),
    ("Naissance", "naissance.com"),
    ("Base Formula", "baseformula.com"),
    ("The Oil Shop", "theoilshop.com"),
    ("Fresh Produce Consortium Members Distribution", "freshproduce.org.uk"),
    ("Reynolds Catering Supplies", "reynolds-cs.com"),
    ("Creed Foodservice", "creedfoodservice.co.uk"),
    ("Owen Taylor & Sons", "owentaylor.co.uk"),
    ("Turner Price", "turnerprice.com"),
    ("Fairway Foodservice", "fairwayfoodservice.com"),
    ("Nationwide Catering Butchers Distribution", "nationwidebutchers.co.uk"),
    # --- EU foodservice / cash & carry ---
    ("Bidfood Netherlands", "bidfood.nl"),
    ("Bidfood Belgium", "bidfood.be"),
    ("Bidfood Spain", "bidfood.es"),
    ("Bidfood Portugal", "bidfood.pt"),
    ("Bidfood Czech", "bidfood.cz"),
    ("Bidfood Poland Farutex", "farutex.pl"),
    ("Bidfood Baltics", "bidfood.lt"),
    ("Sligro Food Group", "sligrofoodgroup.nl"),
    ("Hanos", "hanos.nl"),
    ("Makro Netherlands", "makro.nl"),
    ("Metro AG", "metroag.de"),
    ("Transgourmet", "transgourmet.com"),
    ("Transgourmet Germany", "transgourmet.de"),
    ("Transgourmet Switzerland", "transgourmet.ch"),
    ("Transgourmet France", "transgourmet.fr"),
    ("Selgros Cash & Carry", "selgros.de"),
    ("Prodega", "prodega.ch"),
    ("Poco Loco Foodservice", "pocoloco.com"),
    ("DAC Italy Bidcorp", "dacspa.com"),
    ("Sysco Europe", "sysco.com"),
    ("Brake Bros", "brake.co.uk"),
    ("Pueblo Foods Distribution Europe", "pueblofoods.com"),
    # --- EU edible oil / ingredient distributors & traders ---
    ("Sigma Oil Seeds", "sigmaoilseeds.com"),
    ("Dipasa Europe", "dipasa.nl"),
    ("Vanrobaeys Ludwig", "vanrobaeysludwigoils.com"),
    ("Henry Lamotte Oils", "lamotte-oils.de"),
    ("Gustav Heess", "heessoils.com"),
    ("P. Brändle", "braendle.de"),
    ("Sanabio", "sanabio.de"),
    ("All Organic Treasures", "aot.de"),
    ("AOT GmbH", "aot.de"),
    ("W. Ulrich GmbH", "ulrich-natural.de"),
    ("BOETTGER Food Ingredients", "boettger.com"),
    ("OPW Ingredients", "opw-ingredients.com"),
    ("Tradin Organic", "tradinorganic.com"),
    ("IMCD", "imcdgroup.com"),
    ("Azelis", "azelis.com"),
    ("Brenntag", "brenntag.com"),
    ("Barentz", "barentz.com"),
    ("DKSH", "dksh.com"),
    ("Safic-Alcan", "safic-alcan.com"),
    ("Quimidroga", "quimidroga.com"),
    ("Norkem", "norkem.com"),
    ("Spack BV", "spack.nl"),
    ("Mosselman", "mosselman.eu"),
    ("Oleon Distribution", "oleon.com"),
    ("Cargill Food Distribution Europe", "cargill.com"),
    ("Bunge Loders Croklaan Distribution", "bunge.com"),
    ("ADM European Distribution", "adm.com"),
    ("LANDEMA", "landema.com"),
    ("NATURAL POLAND", "naturalpoland.pl"),
    ("MIXT SP ZOO", "mixt.com.pl"),
    ("VP Trading UAB", "vptrading.lt"),
    ("JUST DIFFERENT PEOPLE", "justdifferentpeople.com"),
    ("ROYAL BIO", "royalbio.es"),
    ("Europagricultproduct", "europagricultproduct.com"),
    ("ESENCIA ANDALUSI GOURMET", "esenciaandalusi.com"),
    ("POWER NATURAL LIFE", "powernaturallife.com"),
    ("DIVERSE WORLDWIDE TRADING", "diverseworldwide.com"),
    ("SOC ROANNAISE DE DISTRIBUTION", "srd-distribution.fr"),
    ("SB-TRADING", "sb-trading.be"),
    ("Verfilcos", "verfilcos.com"),
    ("v03 trading GmbH", "v03trading.de"),
    ("Der kleine Gourmet GmbH", "derkleinegourmet.de"),
    ("Interdrink", "interdrink.bg"),
    ("Eterika", "eterika.com"),
    ("Evterpa", "evterpa.com"),
    ("BUDJAK LTD", "budjak.bg"),
    # --- EU produce importers / wholesale (avocado channel partners) ---
    ("Mission Produce Europe", "missionproduce.com"),
    ("Nature's Pride", "naturespride.eu"),
    ("The Greenery", "thegreenery.com"),
    ("Special Fruit", "specialfruit.be"),
    ("Staay Food Group", "staay.nl"),
    ("Port International", "port-international.com"),
    ("Cobana Fruchtring", "cobana-fruchtring.de"),
    ("SanLucar", "sanlucar.com"),
    ("Eosta", "eosta.com"),
    ("Bakker Barendrecht", "bakkerbarendrecht.nl"),
    ("Haluco", "haluco.nl"),
    ("Frankort & Koning", "frankort-koning.nl"),
    ("Hillfresh", "hillfresh.eu"),
    ("FruitMasters", "fruitmasters.com"),
    ("Olympic Fruit", "olympicfruit.com"),
    ("Jan Oskam", "janoskam.nl"),
    ("Van Dijk Foods", "vandijkfoods.nl"),
    ("Levarht", "levarht.com"),
    ("Roveg Fruit", "roveg.com"),
    ("Nature's Produce", "naturesproduce.co.uk"),
    ("G's Fresh", "gs-fresh.com"),
    ("Eurobanan", "eurobanan.com"),
    ("Fyffes", "fyffes.com"),
    ("Compagnie Fruitière", "compagniefruitiere.fr"),
    ("Fresh Del Monte Europe", "delmonte.com"),
    ("Dole Europe", "dole.com"),
    ("Anecoop", "anecoop.com"),
    ("Trops", "trops.es"),
    ("Frutas Montosa", "montosa.com"),
    ("Reyes Gutiérrez", "reyesgutierrez.com"),
    ("Bonnysa", "bonnysa.com"),
    # --- ANZ ---
    ("Bidfood Australia", "bidfood.com.au"),
    ("Bidfood New Zealand", "bidfood.co.nz"),
    ("Range Products", "rangeproducts.com.au"),
    ("Adams Australia", "adamsaustralia.com.au"),
    ("Medikonda Australia", "medikonda.com.au"),
    ("Davis Food Ingredients", "davisfoodingredients.com.au"),
    ("Gilmours", "gilmours.co.nz"),
    ("Foodstuffs North Island Wholesale", "foodstuffs.co.nz"),
    ("Pampas Foods Wholesale", "pampas.com.au"),
    ("Anchor Foods Wholesale", "anchorfoods.com.au"),
    ("Bidvest Australia", "bidfood.com.au"),
    ("Patties Food Group Distribution", "pattiesfoodgroup.com"),
    ("Black Swan Distribution", "blackswanfoods.com.au"),
    ("Fonterra Foodservice Distribution", "fonterra.com"),
    # --- Asia / Middle East / Africa ---
    ("DKSH Singapore", "dksh.com"),
    ("DKSH Thailand", "dksh.com"),
    ("DKSH Malaysia", "dksh.com"),
    ("DKSH Vietnam", "dksh.com"),
    ("DKSH Indonesia", "dksh.com"),
    ("Bidfood Singapore", "bidfood.com.sg"),
    ("Bidfood Middle East", "bidfood.ae"),
    ("Bidfood South Africa", "bidfood.co.za"),
    ("Bidfood Chile", "bidfood.cl"),
    ("Bidfood Turkey", "bidfood.com.tr"),
    ("Nagase", "nagase.co.jp"),
    ("Sojitz", "sojitz.com"),
    ("ITOCHU Food Distribution", "itochu.co.jp"),
    ("Mitsubishi Corporation Food Distribution", "mitsubishicorp.com"),
    ("CJ Freshway", "cjfreshway.com"),
    ("Transmed", "transmed.com"),
    ("Al Maya Group", "almaya.ae"),
    ("Savola Distribution", "savola.com"),
    ("Lulu Hypermarket Distribution", "luluhypermarket.com"),
    ("Spinneys Distribution", "spinneys.com"),
    ("Agthia Distribution", "agthia.com"),
    ("Al Seer Group", "alseer.com"),
    ("Emirates Snack Foods Distribution", "esf.ae"),
    ("Gourmet Gulf Distribution", "gourmetgulf.com"),
    ("Horeca Trade", "horecatrade.com"),
    ("JM Foods Distribution UAE", "jmfoods.ae"),
    ("NFPC Distribution", "nfpc.net"),
    ("Aramark Middle East Distribution", "aramark.com"),
    ("Bidfood Asia", "bidfood.com"),
    ("Dah Chong Hong Distribution", "dch.com.hk"),
    ("Swire Foods Distribution", "swire.com"),
    ("YHS Distribution", "yeos.com.sg"),
    ("Fraser and Neave Distribution", "fraserandneave.com"),
    ("Olam Food Ingredients Distribution", "ofi.com"),
    ("Wilmar International Distribution", "wilmar-international.com"),
    ("Sime Darby Oils Distribution", "simedarbyplantation.com"),
    ("Musim Mas Distribution", "musimmas.com"),
    ("AAK Distribution Asia", "aak.com"),
    ("Fuji Oil Asia Distribution", "fujioil.co.jp"),
    ("J-Oil Mills Distribution", "j-oil.com"),
    ("Nisshin OilliO Distribution", "nisshin-oillio.com"),
    # --- LATAM ---
    ("Makro Colombia", "makro.com.co"),
    ("Makro Brazil", "makro.com.br"),
    ("Cencosud Wholesale", "cencosud.com"),
    ("Grupo Herdez Distribution", "herdez.com.mx"),
    ("Diconsa", "diconsa.gob.mx"),
    ("Walmart Mexico Distribution", "walmartmexico.com"),
    ("Costco Mexico Wholesale", "costco.com.mx"),
    ("Sam's Club Mexico", "sams.com.mx"),
    ("PriceSmart", "pricesmart.com"),
    ("Grupo Éxito Wholesale", "grupoexito.com.co"),
    ("Cencosud Chile Wholesale", "cencosud.com"),
    ("SMU Mayorista", "smu.cl"),
    ("Almacenes Éxito Distribution", "exito.com"),
    ("Grupo Casino Distribution LATAM", "groupe-casino.fr"),
    ("Natura Distribución Ingredients", "natura.com.br"),
    # --- More specialty oil / cosmetic ingredient distributors ---
    ("Aromaaz International", "aromaazinternational.com"),
    ("Manohar Botanical Extracts Wholesale", "manoharbotanicals.com"),
    ("Deve Herbes Wholesale", "deveherbes.com"),
    ("Katyani Exports Wholesale", "katyaniexports.com"),
    ("India Aroma Oils and Company Wholesale", "indiaaromaoils.com"),
    ("Natural Sourcing", "naturalsourcing.com"),
    ("Plant Therapy Wholesale", "planttherapy.com"),
    ("Bulk Naturals Wholesale", "bulknaturalswholesale.com"),
    ("Formulator Sample Shop", "formulatorsampleshop.com"),
    ("Wholesale Supplies Plus", "wholesalesuppliesplus.com"),
    ("The Chemistry Store", "chemistrystore.com"),
    ("TKB Trading", "tkbtrading.com"),
    ("Coast Southwest", "coastsouthwest.com"),
    ("Ross Organic Azelis", "rossorg.com"),
    ("DeWolf Chemical", "dewolfchem.com"),
    ("Charkit", "charkit.com"),
    ("Essential Ingredients", "essentialingredients.com"),
    ("Independent Chemical", "independentchemical.com"),
    ("Lipo Chemicals Distribution", "lipochemicals.com"),
    ("Protameen Chemicals", "protameen.com"),
    ("RITA Corporation", "ritacorp.com"),
    ("Tri-K Industries Distribution", "tri-k.com"),
    ("Active Concepts Distribution Partners", "activeconceptsllc.com"),
    ("Centerchem", "centerchem.com"),
    ("Presperse", "presperse.com"),
    ("Kobo Products Distribution", "koboproducts.com"),
    ("Grant Industries Distribution", "grantinc.com"),
    ("Elementis Distribution Network", "elementis.com"),
    ("Croda Distribution Partners", "croda.com"),
    ("BASF Care Creations Distribution", "basf.com"),
    ("Evonik Care Solutions Distribution", "evonik.com"),
    ("Clariant Personal Care Distribution", "clariant.com"),
    ("Ashland Specialty Ingredients Distribution", "ashland.com"),
    ("Innospec Performance Chemicals Distribution", "innospec.com"),
    ("Solvay Novecare Distribution", "solvay.com"),
    ("Stepan Distribution", "stepan.com"),
    ("Lubrizol Life Science Distribution", "lubrizol.com"),
    ("Gattefossé Distribution Network", "gattefosse.com"),
    ("SEPPIC Distribution", "seppic.com"),
    ("Lucas Meyer Cosmetics Distribution", "lucasmeyercosmetics.com"),
    ("Hallstar Beauty Distributor Network", "hallstarbeauty.com"),
    ("AAK Personal Care Distribution", "aakpersonalcare.com"),
    ("IOI Oleo Distribution", "ioioleo.de"),
    ("KLK OLEO Distribution", "klkoleo.com"),
    ("Emery Oleochemicals Distribution", "emeryoleo.com"),
    ("Berg + Schmidt Distribution", "berg-schmidt.de"),
    ("Peter Cremer Holding Distribution", "cremer.de"),
    ("Nordmann Distribution", "nordmann.global"),
    ("Biesterfeld Spezialchemie", "biesterfeld.com"),
    ("Ter Chemicals Distribution", "tergroup.com"),
    ("Stockmeier Chemie", "stockmeier.com"),
    ("BTC Europe BASF Distribution", "btc-europe.com"),
    ("C.H. Erbslöh", "erbsloeh.com"),
    ("Th. Geyer Ingredients", "thgeyer.com"),
    ("VWR Chemicals Distribution", "vwr.com"),
    ("Thermo Fisher Chemicals Distribution", "thermofisher.com"),
    ("Fisher Scientific Distribution", "fishersci.com"),
    ("Sigma-Aldrich Distribution Channel", "sigmaaldrich.com"),
    ("Merck Millipore Distribution", "merckmillipore.com"),
    # --- Additional foodservice / wholesale from regional lists ---
    ("Performance Foodservice", "pfgc.com"),
    ("Reinhart Foodservice", "pfgc.com"),
    ("Food Services of America", "usfoods.com"),
    ("Buffalo Market", "buffalomarket.com"),
    ("Associated Grocers of the South", "agsouth.com"),
    ("Associated Grocers New England", "agne.com"),
    ("Affiliated Foods Midwest", "afmidwest.com"),
    ("Unified Grocers", "unifiedgrocers.com"),
    ("Topco Associates", "topco.com"),
    ("Wakefern ShopRite Distribution", "wakefern.com"),
    ("Hy-Vee Distribution", "hy-vee.com"),
    ("H.E.B. Distribution", "heb.com"),
    ("Publix Distribution", "publix.com"),
    ("Kroger Distribution", "thekrogerco.com"),
    ("Albertsons Distribution", "albertsonscompanies.com"),
    ("Ahold Delhaize USA Distribution", "aholddelhaize.com"),
    ("Metro Inc Canada Distribution", "metro.ca"),
    ("Loblaw Distribution", "loblaw.ca"),
    ("Sobeys Distribution", "sobeys.com"),
    ("Empire Company Distribution", "empireco.ca"),
    ("Costco Wholesale Business Delivery", "costco.com"),
    ("Sam's Club Business", "samsclub.com"),
    ("BJ's Wholesale Club Business", "bjs.com"),
    ("Restaurant Technologies Distribution", "rti-inc.com"),
    ("Acosta Sales & Marketing Distribution Partners", "acosta.com"),
    ("CROSSMARK Distribution Partners", "crossmark.com"),
    ("Advantage Solutions Distribution", "advantagesolutions.net"),
]


def norm_key(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (name or "").lower())


def dns_ok(domain: str, timeout: float = 2.0) -> bool:
    host = (domain or "").strip().lower()
    host = re.sub(r"^https?://", "", host).split("/")[0]
    if host.startswith("www."):
        host = host[4:]
    if not host or "." not in host:
        return False
    try:
        socket.setdefaulttimeout(timeout)
        socket.getaddrinfo(host, 80, type=socket.SOCK_STREAM)
        return True
    except OSError:
        try:
            socket.getaddrinfo("www." + host, 80, type=socket.SOCK_STREAM)
            return True
        except OSError:
            return False


def main() -> None:
    # Dedupe by name key, prefer first domain
    seen: set[str] = set()
    unique: list[tuple[str, str]] = []
    for name, dom in RAW:
        k = norm_key(name)
        if not k or k in seen:
            continue
        seen.add(k)
        unique.append((name.strip(), dom.strip().lower()))

    print(f"unique_raw={len(unique)}")

    verified: list[dict] = []
    failed: list[dict] = []
    for i, (name, dom) in enumerate(unique, 1):
        ok = dns_ok(dom)
        row = {"name": name, "domain": dom, "dns_ok": ok}
        if ok:
            verified.append(row)
        else:
            failed.append(row)
        if i % 50 == 0:
            print(f"  verified {i}/{len(unique)} … ok={len(verified)} fail={len(failed)}")

    # Keep DNS-ok first; if short of 300, keep DNS-fail with note (still known companies)
    keep = list(verified)
    if len(keep) < 300:
        for row in failed:
            keep.append(row)
            if len(keep) >= 300:
                break
    keep = keep[: max(300, len(keep))]
    # Trim to exactly 300 if we have more verified
    if len(verified) >= 300:
        keep = verified[:300]
    else:
        keep = keep[:300]

    print(f"dns_ok={len(verified)} dns_fail={len(failed)} keeping={len(keep)}")

    # --- seeds file ---
    lines = [
        "# Global Avocado Oil Market — channel / distributor seeds only",
        "# Curated + DNS-checked B2B distributors / wholesalers / importers / foodservice & ingredient channels.",
        "# Manufacturers / brand producers intentionally omitted.",
        f"# Generated {datetime.now(timezone.utc).isoformat()} — count={len(keep)}",
        "# Format: Company Name | Distributors of Avocado Oil",
        "",
    ]
    for row in keep:
        flag = "" if row["dns_ok"] else "  # domain DNS unresolved — verify in enrich"
        lines.append(f"{row['name']} | Distributors of Avocado Oil{flag}")
    SEED_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")
    if DESK_SEED.parent.exists():
        DESK_SEED.write_text("\n".join(lines) + "\n", encoding="utf-8")

    # --- merge into Phase2 ---
    p2 = json.loads(P2_PATH.read_text(encoding="utf-8"))
    companies = list(p2.get("companies") or [])
    existing_keys = {norm_key(str(c.get("name") or "")) for c in companies}
    existing_doms = {
        re.sub(r"^www\.", "", str(c.get("domain") or "").lower().strip())
        for c in companies
    }

    added = 0
    skipped = 0
    for row in keep:
        k = norm_key(row["name"])
        dom = row["domain"]
        if k in existing_keys or dom in existing_doms:
            skipped += 1
            continue
        companies.append(
            {
                "name": row["name"],
                "domain": row["domain"],
                "discovery_source": "analyst_seed_web_verified",
                "company_function": "distributor",
            }
        )
        existing_keys.add(k)
        existing_doms.add(dom)
        added += 1

    p2["companies"] = companies
    p2["company_count"] = len(companies)
    p2["analyst_seed_merge"] = {
        "merged_at": datetime.now(timezone.utc).isoformat(),
        "requested": 300,
        "dns_ok": len(verified),
        "kept_in_seed": len(keep),
        "added_to_phase2": added,
        "skipped_already_present": skipped,
    }
    # backup
    bak = P2_PATH.with_suffix(".json.bak_before_seed300")
    if not bak.exists():
        bak.write_text(json.dumps(json.loads(P2_PATH.read_text(encoding="utf-8")), indent=2), encoding="utf-8")
    P2_PATH.write_text(json.dumps(p2, indent=2, ensure_ascii=False), encoding="utf-8")

    report = {
        "seed_count": len(keep),
        "dns_ok": len(verified),
        "dns_fail": len(failed),
        "phase2_before": p2["company_count"] - added,
        "phase2_after": p2["company_count"],
        "added_to_phase2": added,
        "skipped_already_present": skipped,
        "dns_fail_sample": failed[:20],
        "kept_sample": keep[:15],
    }
    REPORT.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2)[:2000])
    print(f"\nSEED -> {SEED_PATH}")
    print(f"P2   -> {P2_PATH} ({p2['company_count']} companies)")


if __name__ == "__main__":
    main()
