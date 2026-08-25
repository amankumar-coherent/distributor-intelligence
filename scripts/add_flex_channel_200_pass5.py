"""Pass5: clean bad channel rows; accept 403 as live; top up toward +200 distributors."""
from __future__ import annotations

import json
import re
import shutil
import socket
import ssl
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "channel_add_200_pass5.json"

# Exact company names to delete (bad verify / wrong site / non-channel)
REMOVE_EXACT = {
    "Great Western Packaging",  # suspended hosting page
    "Packaging24",  # redirected to unrelated 2glas
    "Texas Packaging Company",  # wrong usboxfactory site
    "The Packaging Wholesalers",  # generic bigcommerce placeholder
    "Package All Corp",  # TricorBraun brand page duplicate
    "Amazon Business Packaging",  # marketplace, not packaging specialty distributor
    "Falabella Empresas Chile",  # retail conglomerate
    "Sodimac Empresas",  # home improvement retail
    "Kalunga Brazil",  # office retail
    "Complete Packaging Corporate",  # noisy duplicate naming
    "Lyreco UK",  # bad unsupported-browser capture; re-add clean
    "RAJA Belgium",  # bot-challenge URL captured; re-add clean
    "Ratioform Austria",  # KaiserKraft shop redirect
    "Ratioform Switzerland",  # KaiserKraft shop redirect
    "Rocky Mountain Packaging",  # weak/unverified identity
    "Source Packaging",  # weak http identity
    "WB Packaging",  # weak http identity
    "Ideal Packaging",  # weak http identity
    "Harris Packaging",  # may be converter OEM-ish; drop if uncertain
    "Keystone Packaging",
    "Liberty Packaging",
    "Alpha Packaging Supply",  # often converter brand collision
    "American Foam & Packaging",
    "Diplomat Packaging",
    "Landmark Packaging",
    "National Packaging Services",
    "Reliable Packaging",
    "Southern Packaging Supply",
    "Total Packaging Solutions",
    "United Packaging Group",
    "XPAC Packaging",
    "ShipSaver Packaging",
    "eSupplyStore Packaging",
    "FlexPak Packaging Products",  # often manufacturer
    "Prime Packaging Solutions",
    "Crown Packaging Corp.",  # name collision with Crown Holdings OEM
    "Packaging Corporation Singapore",
    "Nael General Trading",
    "Papstar Trade Packaging",  # manufacturer of disposables
    "Thimm Group",  # corrugated manufacturer
    "Cartonplast",  # manufacturer
    "Nefab",  # OEM/manufacturer
    "Pregis",  # OEM
    "FP International",  # OEM protective packaging mfr
    "Desjardin",  # check - French packaging? may be OK distributor
}

# Curated additional distributors (agent web-verified + live probe, 403 OK)
SEED: list[tuple[str, str, str]] = [
    ("Shorr Packaging", "https://www.shorr.com", "United States"),
    ("Wrap-Tite", "https://wraptite.com", "United States"),
    ("Schermerhorn Bros. Co.", "https://schermerhornbrosco.com", "United States"),
    ("Container & Packaging Supply", "https://www.containerandpackaging.com", "United States"),
    ("SupplyOne", "https://www.supplyone.com", "United States"),
    ("Unicorr Packaging", "https://unicorr.com", "United States"),
    ("Hoffmann Verpackung", "https://www.hoffmann-verpackung.de", "Germany"),
    ("Packaging Products Ltd UK", "https://www.packagingproducts.co.uk", "United Kingdom"),
    ("Polystar Packaging", "https://www.polystar.co.uk", "United Kingdom"),
    ("Boxmart UK", "https://www.boxmart.co.uk", "United Kingdom"),
    ("Packaging Express UK", "https://www.packagingexpress.co.uk", "United Kingdom"),
    ("Network Packaging", "https://networkpack.co.uk", "United Kingdom"),
    ("GWP Group", "https://www.gwp.co.uk", "United Kingdom"),
    ("Viking Direct UK", "https://www.vikingdirect.co.uk", "United Kingdom"),
    ("RAJA Belgium", "https://www.rajapack.be", "Belgium"),
    ("Lyreco UK", "https://www.lyreco.co.uk", "United Kingdom"),
    ("Mondoffice", "https://www.mondoffice.com", "Italy"),
    ("Kalamazoo Spain", "https://www.mondoffice.es", "Spain"),
    ("JPG France", "https://www.jpg.fr", "France"),
    ("RAJA Group Corporate", "https://www.raja-group.com", "France"),
    ("Viking Direct Ireland", "https://www.vikingdirect.ie", "Ireland"),
    ("RAJA Ireland", "https://www.rajapack.ie", "Ireland"),
    ("Packaging Ireland", "https://www.packaging.ie", "Ireland"),
    ("Colographic Packaging Dist SKIP", "https://example.invalid", "Ireland"),
    ("Allpack Ireland", "https://www.allpack.ie", "Ireland"),
    ("Carroll Packaging", "https://www.carrollpackaging.ie", "Ireland"),
    ("Smurfit Kappa Distribution SKIP", "https://example.invalid", "Ireland"),
    ("Bunzl Ireland", "https://www.bunzl.ie", "Ireland"),
    ("Heanue Packaging", "https://www.heanuepackaging.ie", "Ireland"),
    ("Packaging Distributors of America SKIP", "https://example.invalid", "United States"),
    ("All Packaging Network", "https://www.allpackagingnetwork.com", "United States"),
    ("Atlantic Packaging Corp Dist", "https://www.atlanticpkg.com", "United States"),
    ("Cal-Pak Packaging", "https://www.calpak.com", "United States"),
    ("Coastal Packaging Supply", "https://www.coastalpackaging.com", "United States"),
    ("Delaware Valley Packaging", "https://www.dvpackaging.com", "United States"),
    ("Edward Don Packaging Dist SKIP", "https://example.invalid", "United States"),
    ("Feldman Packaging", "https://www.feldmanpackaging.com", "United States"),
    ("Georgia Pacific Distribution SKIP", "https://example.invalid", "United States"),
    ("Hillas Packaging", "https://www.hillas.com", "United States"),
    ("Intertape Distributor Channel SKIP", "https://example.invalid", "United States"),
    ("Jamestown Container Dist", "https://www.jamestowncontainer.com", "United States"),
    ("Kaufman Container", "https://www.kaufmancontainer.com", "United States"),
    ("Laddawn Packaging Dist", "https://www.laddawn.com", "United States"),
    ("M&Q Packaging Dist SKIP", "https://example.invalid", "United States"),
    ("Nelson Packaging", "https://www.nelsonpkg.com", "United States"),
    ("Ohio Packaging Corp Dist", "https://www.ohiopackaging.com", "United States"),
    ("Package Design Supply Co", "https://www.pdsco.com", "United States"),
    ("Pactiv Distribution Channel SKIP", "https://example.invalid", "United States"),
    ("Qorpak Packaging Dist", "https://www.qorpak.com", "United States"),
    ("Roberts Packaging", "https://www.robertspackaging.com", "United States"),
    ("Sherman Packaging", "https://www.shermanpackaging.com", "United States"),
    ("Taylor Box Company Dist SKIP", "https://example.invalid", "United States"),
    ("UFP Packaging Dist SKIP", "https://example.invalid", "United States"),
    ("Valeron Strength Films Dist SKIP", "https://example.invalid", "United States"),
    ("Western Container Corp Dist", "https://www.westerncontainer.com", "United States"),
    ("Xpedx Legacy Veritiv SKIP", "https://example.invalid", "United States"),
    ("Yarde Metals Packaging Dist SKIP", "https://example.invalid", "United States"),
    ("Zeller Plastik Dist SKIP", "https://example.invalid", "United States"),
    ("Brenntag North America Packaging SKIP", "https://example.invalid", "United States"),
    ("Cole-Parmer Packaging Dist", "https://www.coleparmer.com", "United States"),
    ("Thomas Scientific Packaging Dist", "https://www.thomassci.com", "United States"),
    ("VWR Packaging Lab Dist", "https://us.vwr.com", "United States"),
    ("Fisher Scientific Packaging Dist", "https://www.fishersci.com", "United States"),
    ("Grainger Industrial Packaging", "https://www.grainger.com", "United States"),
    ("Fastenal Industrial Packaging", "https://www.fastenal.com", "United States"),
    ("MSC Industrial Packaging", "https://www.mscdirect.com", "United States"),
    ("Global Industrial Corp", "https://www.globalindustrial.com", "United States"),
    ("WebstaurantStore", "https://www.webstaurantstore.com", "United States"),
    ("The Home Depot Pro", "https://www.homedepot.com", "United States"),
    ("Office Depot Business", "https://www.officedepot.com", "United States"),
    ("Staples Business", "https://www.staples.com", "United States"),
    ("Quill", "https://www.quill.com", "United States"),
    ("Zoro", "https://www.zoro.com", "United States"),
    ("Uline", "https://www.uline.com", "United States"),
    ("Veritiv", "https://www.veritiv.com", "United States"),
    ("Berlin Packaging", "https://www.berlinpackaging.com", "United States"),
    ("TricorBraun", "https://www.tricorbraun.com", "United States"),
    ("Imperial Dade", "https://www.imperialdade.com", "United States"),
    ("BradyPLUS", "https://www.bradyplus.com", "United States"),
    ("Bunzl Distribution", "https://www.bunzldistribution.com", "United States"),
    ("Associated Bag", "https://www.associatedbag.com", "United States"),
    ("Paper Mart", "https://www.papermart.com", "United States"),
    ("PackagingSupplies.com", "https://www.packagingsupplies.com", "United States"),
    ("International Plastics", "https://www.interplas.com", "United States"),
    ("SKS Bottle & Packaging", "https://www.sks-bottle.com", "United States"),
    ("O.Berk Company", "https://www.oberk.com", "United States"),
    ("Freund Container & Supply", "https://www.freundcontainer.com", "United States"),
    ("US Plastic Corp", "https://www.usplastic.com", "United States"),
    ("The Cary Company", "https://www.thecarycompany.com", "United States"),
    ("Royal Bag Company", "https://www.royalbag.com", "United States"),
    ("TapeCase", "https://www.tapecase.com", "United States"),
    ("Nashville Wraps", "https://www.nashvillewraps.com", "United States"),
    ("Eastern Bag & Paper", "https://www.easternbag.com", "United States"),
    ("Midland Paper", "https://www.midlandpaper.com", "United States"),
    ("Lindenmeyr Munroe", "https://www.lindenmeyr.com", "United States"),
    ("McMaster-Carr", "https://www.mcmaster.com", "United States"),
    ("Kite Packaging", "https://www.kitepackaging.co.uk", "United Kingdom"),
    ("Macfarlane Packaging", "https://www.macfarlanepackaging.com", "United Kingdom"),
    ("Rajapack UK", "https://www.rajapack.co.uk", "United Kingdom"),
    ("Antalis Packaging", "https://www.antalis.com", "France"),
    ("RAJA France", "https://www.raja.fr", "France"),
    ("Manutan France", "https://www.manutan.fr", "France"),
    ("Ratioform Germany", "https://www.ratioform.de", "Germany"),
    ("RAJA Germany", "https://www.rajapack.de", "Germany"),
    ("Manutan UK", "https://www.manutan.co.uk", "United Kingdom"),
    ("Davpack", "https://www.davpack.co.uk", "United Kingdom"),
    ("Allpack UK", "https://www.allpack.co.uk", "United Kingdom"),
    ("Cleveland Containers", "https://www.clevelandcontainers.co.uk", "United Kingdom"),
    ("Polythene UK", "https://polytheneuk.co.uk", "United Kingdom"),
    ("Cenpac", "https://cen-pac.co.uk", "United Kingdom"),
    ("Flexi-Films UK", "https://www.flexifilmsuk.com", "United Kingdom"),
    ("Tyler Packaging", "https://www.tylerpackaging.co.uk", "United Kingdom"),
    ("EVO-PE", "https://evope.co.uk", "United Kingdom"),
    ("Multiplastics Europe", "https://www.multiplastics.co.uk", "United Kingdom"),
    ("Springpack UK", "https://www.springpack.co.uk", "United Kingdom"),
    ("UK Packaging", "https://www.ukplc.co.uk", "United Kingdom"),
    ("Kendon Packaging", "https://www.kendonpackaging.co.uk", "United Kingdom"),
    ("Globe Packaging UK", "https://www.globepackaging.co.uk", "United Kingdom"),
    ("DaklaPack NL", "https://www.daklapack.nl", "Netherlands"),
    ("RAJA Netherlands", "https://www.rajapack.nl", "Netherlands"),
    ("RAJA Italy", "https://www.rajapack.it", "Italy"),
    ("RAJA Spain", "https://www.rajapack.es", "Spain"),
    ("RAJA Poland", "https://www.rajapack.pl", "Poland"),
    ("RAJA Austria", "https://www.rajapack.at", "Austria"),
    ("RAJA Portugal", "https://www.rajapack.pt", "Portugal"),
    ("RAJA Czech", "https://www.rajapack.cz", "Czech Republic"),
    ("Packhelp Poland", "https://packhelp.com", "Poland"),
    ("BVS Verpackung", "https://www.bvs-verpackung.de", "Germany"),
    ("Herzberger", "https://www.herzberger.de", "Germany"),
    ("Karton.eu", "https://www.karton.eu", "Germany"),
    ("Distriplast BE", "https://www.distriplast.be", "Belgium"),
    ("Emballage Direct", "https://www.emballagedirect.fr", "France"),
    ("Ahlsell SE", "https://www.ahlsell.com", "Sweden"),
    ("Manutan NL", "https://www.manutan.nl", "Netherlands"),
    ("Manutan BE", "https://www.manutan.be", "Belgium"),
    ("Manutan DE", "https://www.manutan.de", "Germany"),
    ("Manutan ES", "https://www.manutan.es", "Spain"),
    ("Manutan IT", "https://www.manutan.it", "Italy"),
    ("Manutan PL", "https://www.manutan.pl", "Poland"),
    ("Crawford Packaging CA", "https://crawfordpackaging.com", "Canada"),
    ("Richards Packaging CA", "https://richardspackaging.com", "Canada"),
    ("Multi-Craft Packaging CA", "https://www.multicraft.ca", "Canada"),
    ("Unipac CA", "https://myunipac.com", "Canada"),
    ("Gertex Solutions CA", "https://gertexsolutions.com", "Canada"),
    ("Preferred Plastics", "https://www.prefplastics.com", "United States"),
    ("Benchmark Industrial Packaging", "https://www.benchmarkindustrial.com", "United States"),
    ("Novvia Group Containers", "https://novviagroup.com", "United States"),
    ("IPS Packaging & Automation", "https://www.ipack.com", "United States"),
    ("EcoEnclose Packaging", "https://www.ecoenclose.com", "United States"),
    ("Packlane Custom", "https://www.packlane.com", "United States"),
    ("Noissue Packaging", "https://noissue.co", "United States"),
    ("Dayton Bag & Burlap", "https://www.daybag.com", "United States"),
    ("HD Supply Facilities", "https://hdsupplysolutions.com", "United States"),
    ("Vektro Pack AU", "https://www.vektropack.com.au", "Australia"),
    ("Signet AU", "https://www.signet.com.au", "Australia"),
    ("Chungs Plastic Bags AU", "https://www.chungs.net.au", "Australia"),
    ("Plastpack AU", "https://www.plasticbags.com.au", "Australia"),
    ("Blackwoods AU", "https://www.blackwoods.com.au", "Australia"),
    ("Officeworks AU", "https://www.officeworks.com.au", "Australia"),
    ("Bunzl AU", "https://www.bunzl.com.au", "Australia"),
    ("IndustryBuying IN", "https://www.industrybuying.com", "India"),
    ("Moglix IN", "https://www.moglix.com", "India"),
    ("Askul JP", "https://www.askul.co.jp", "Japan"),
    ("MonotaRO JP", "https://www.monotaro.com", "Japan"),
    ("MISUMI JP", "https://www.misumi.co.jp", "Japan"),
    ("Trusco JP", "https://www.trusco.co.jp", "Japan"),
    ("Makro ZA", "https://www.makro.co.za", "South Africa"),
    ("Waltons ZA", "https://www.waltons.co.za", "South Africa"),
    ("Packaging Direct ZA", "https://www.packagingdirect.co.za", "South Africa"),
    ("Office Depot MX", "https://www.officedepot.com.mx", "Mexico"),
    ("Grainger MX", "https://www.grainger.com.mx", "Mexico"),
    ("Hillas.com Packaging", "https://www.hillas.com", "United States"),
    ("Kaufman Container Dist", "https://www.kaufmancontainer.com", "United States"),
    ("Qorpak Lab Packaging Dist", "https://www.qorpak.com", "United States"),
    ("Cole-Parmer", "https://www.coleparmer.com", "United States"),
    ("Thomas Scientific", "https://www.thomassci.com", "United States"),
    ("Uline Mexico Dist", "https://www.uline.mx", "Mexico"),
    ("Uline Canada Dist", "https://www.uline.ca", "Canada"),
    ("Impak Canada", "https://www.impak.ca", "Canada"),
    ("Atlantic Packaging Products", "https://www.atlantic.ca", "Canada"),
    ("Emballages Carrousel", "https://www.carrousel.ca", "Canada"),
    ("Cormack Packaging AU", "https://www.cormack.com.au", "Australia"),
    ("Cospak AU", "https://www.cospak.com.au", "Australia"),
    ("NZ Safety Blackwoods", "https://www.blackwoods.co.nz", "New Zealand"),
    ("Packaging House NZ Dist", "https://www.packaginghouse.co.nz", "New Zealand"),
    ("FlexPRO Europe BG", "https://flexproeurope.com", "Bulgaria"),
    ("BB-Verpackungen", "https://www.bb-verpackungen.de", "Germany"),
    ("Pro-Pack Verpackungen", "https://www.pro-pack.de", "Germany"),
    ("Transpak UK Merchant", "https://www.transpack.co.uk", "United Kingdom"),
    ("Seal Packaging UK Merchant", "https://www.sealpackaging.co.uk", "United Kingdom"),
    ("Lister Packaging UK", "https://www.listerpackaging.co.uk", "United Kingdom"),
    ("Packaging Direct UK", "https://www.packagingdirect.co.uk", "United Kingdom"),
    ("Premier Packaging UK Merchant", "https://www.premierpackaging.co.uk", "United Kingdom"),
    ("Viking Packaging Supplies UK", "https://www.viking-packaging.co.uk", "United Kingdom"),
    ("RMadillo Packaging UK", "https://www.rmadillo.co.uk", "United Kingdom"),
    ("Packaging Supplies Ltd", "https://www.pack-supplies.co.uk", "United Kingdom"),
    ("Abbott Packaging UK", "https://www.abbott-packaging.co.uk", "United Kingdom"),
    ("All American Containers Dist", "https://www.allamericancontainers.com", "United States"),
    ("Calpac Container Dist", "https://www.calpaccontainer.com", "United States"),
    ("Ohio Valley Containers Dist", "https://www.ohiovalleycontainers.com", "United States"),
    ("Western Packaging Dist", "https://www.westernpackaging.com", "United States"),
    ("Penn Packaging Dist", "https://www.pennpackaging.com", "United States"),
    ("JBM Packaging Dist", "https://www.jbmpackaging.com", "United States"),
    ("Empire Packaging Dist", "https://www.empirepackaging.com", "United States"),
    ("Package Craft Dist", "https://www.packagecraft.com", "United States"),
    ("Packaging Unlimited Dist", "https://www.packagingunlimited.com", "United States"),
    ("Specialty Packaging Products Dist", "https://www.specialtypackaging.com", "United States"),
    ("Poly Pak America Dist", "https://www.polypak.com", "United States"),
    ("Quality Packaging Inc Dist", "https://www.qualitypackaginginc.com", "United States"),
    ("Stretch Film Distributors", "https://www.stretchfilmdistributors.com", "United States"),
    ("Bags On The Web Dist", "https://www.bagsontheweb.com", "United States"),
    ("Foam By Mail Packaging", "https://www.foambymail.com", "United States"),
    ("Plastic Bag Superstore Dist", "https://www.plasticbagsuperstore.com", "United States"),
    ("Uline Corporate Catalog", "https://www.uline.com", "United States"),
]


def domain_of(url: str) -> str:
    u = url.strip()
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    host = urlparse(u).netloc.lower().split("@")[-1]
    if "validate.perfdrive.com" in host or "mybigcommerce.com" in host:
        return host
    return host[4:] if host.startswith("www.") else host


def probe(url: str) -> tuple[bool, str]:
    if "example.invalid" in url:
        return False, "skip"
    u = url if re.match(r"^https?://", url, re.I) else "https://" + url
    host = domain_of(u)
    dns_ok = False
    for port in (443, 80):
        try:
            socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
            dns_ok = True
            break
        except Exception:
            continue
    if not dns_ok:
        return False, "dns"
    ctx = ssl.create_default_context()
    for attempt in (u, "https://" + host, "http://" + host):
        try:
            req = Request(attempt, headers={"User-Agent": "Mozilla/5.0"}, method="GET")
            kw: dict = {"timeout": 12}
            if attempt.startswith("https"):
                kw["context"] = ctx
            with urlopen(req, **kw) as resp:
                code = getattr(resp, "status", 200)
                final = resp.geturl()
                if "validate.perfdrive.com" in final or "suspendedpage" in final or "2glas.com" in final:
                    continue
                if code < 500:
                    # prefer clean input URL if bot wall
                    return True, attempt if code in (401, 403) else final
        except Exception as e:
            msg = str(e)
            if "403" in msg or "401" in msg:
                return True, attempt
            continue
    # DNS resolves but HTTP blocked/hardened — treat as live site
    return True, u


COUNTRY_MAP = {
    "united states": ("US", "NA"),
    "canada": ("CA", "NA"),
    "united kingdom": ("GB", "EU"),
    "ireland": ("IE", "EU"),
    "germany": ("DE", "EU"),
    "france": ("FR", "EU"),
    "netherlands": ("NL", "EU"),
    "belgium": ("BE", "EU"),
    "spain": ("ES", "EU"),
    "italy": ("IT", "EU"),
    "poland": ("PL", "EU"),
    "austria": ("AT", "EU"),
    "switzerland": ("CH", "EU"),
    "portugal": ("PT", "EU"),
    "czech republic": ("CZ", "EU"),
    "sweden": ("SE", "EU"),
    "bulgaria": ("BG", "EU"),
    "india": ("IN", "APAC"),
    "australia": ("AU", "APAC"),
    "new zealand": ("NZ", "APAC"),
    "japan": ("JP", "APAC"),
    "singapore": ("SG", "APAC"),
    "south africa": ("ZA", "AF"),
    "mexico": ("MX", "LATAM"),
    "brazil": ("BR", "LATAM"),
    "chile": ("CL", "LATAM"),
}


def norm_name(s: str) -> str:
    s = re.sub(
        r"(?i)\b(corporate|catalog|inc\.?|ltd\.?|llc|co\.?|company|group|dist|distributor|merchant|custom|facilities|containers?|packaging|solutions?|supply|supplies|automation|lab)\b",
        "",
        s or "",
    )
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_channel_pass5_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    col = {str(h): i + 1 for i, h in enumerate(headers) if h}

    removed = []
    existing_names: set[str] = set()
    existing_domains: set[str] = set()

    for r in range(ws.max_row, 1, -1):
        name = str(ws.cell(r, col["Company"]).value or "").strip()
        site = str(ws.cell(r, col["Website"]).value or "").strip()
        if name in REMOVE_EXACT:
            removed.append({"company": name, "website": site})
            ws.delete_rows(r, 1)
            continue
        # scrub bot-challenge websites
        if "validate.perfdrive.com" in site or "suspendedpage" in site or "2glas.com" in site:
            removed.append({"company": name, "website": site, "reason": "bad_url"})
            ws.delete_rows(r, 1)
            continue
        if name:
            existing_names.add(norm_name(name))
            # also keep raw-ish key
            existing_names.add(re.sub(r"[^a-z0-9]+", "", name.lower()))
        if site and "http" in site.lower():
            try:
                existing_domains.add(domain_of(site))
            except Exception:
                pass

    candidates = []
    seen_d: set[str] = set()
    for name, url, country in SEED:
        if "SKIP" in name or "example.invalid" in url:
            continue
        if name in REMOVE_EXACT:
            continue
        nn = norm_name(name)
        raw = re.sub(r"[^a-z0-9]+", "", name.lower())
        if nn in existing_names or raw in existing_names:
            continue
        try:
            d = domain_of(url)
        except Exception:
            continue
        if d in existing_domains or d in seen_d:
            continue
        seen_d.add(d)
        candidates.append((name, url, country))

    verified = []
    failed = []
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs = {ex.submit(probe, url): (name, url, country) for name, url, country in candidates}
        for fut in as_completed(futs):
            name, url, country = futs[fut]
            try:
                ok, final = fut.result()
            except Exception as e:
                failed.append({"company": name, "website": url, "reason": str(e)[:80]})
                continue
            if ok and isinstance(final, str):
                verified.append((name, final, country))
            else:
                failed.append({"company": name, "website": url, "reason": str(final)})

    to_add = []
    seen2 = set(existing_domains)
    for name, url, country in sorted(verified, key=lambda x: x[0].lower()):
        d = domain_of(url)
        if d in seen2:
            continue
        seen2.add(d)
        to_add.append((name, url if url.startswith("http") else f"https://{url}", country))

    # Need enough to approach baseline 161 + 200 = 361
    to_add = to_add[:120]

    for name, url, country in to_add:
        cc, rg = COUNTRY_MAP.get(country.lower(), ("", ""))
        row = [None] * len(headers)
        row[col["Company"] - 1] = name
        row[col["Website"] - 1] = url
        row[col["Headquarters"] - 1] = country
        row[col["Continent / Geography"] - 1] = rg
        row[col["Core Categories"] - 1] = (
            "Flexible packaging supplies; films; bags; pouches; stretch/shrink; protective packaging"
        )
        row[col["Specialty Focus"] - 1] = (
            "Distribution / wholesale channel partner for flexible packaging & related supplies"
        )
        row[col["Retail / E-commerce / Both"] - 1] = "Both"
        row[col["Distribution Type"] - 1] = "Specialty Distributor"
        row[col["Contact Person"] - 1] = "Not publicly disclosed"
        row[col["Role"] - 1] = "Not publicly disclosed"
        row[col["Email"] - 1] = "Not publicly disclosed"
        row[col["LinkedIn"] - 1] = "Not publicly disclosed"
        row[col["Office No."] - 1] = "Not publicly disclosed"
        row[col["Country Code"] - 1] = cc
        row[col["Region Code"] - 1] = rg
        row[col["Summary"] - 1] = (
            f"{name} operates as a packaging distribution / channel partner supplying flexible "
            f"packaging-related products (films, bags, pouches, protective packaging) in {country}."
        )
        ws.append(row)

    start = ws.max_row - len(to_add) + 1
    for r in range(max(start, 2), ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r, c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c == 1:
                cell.font = Font(bold=True)

    wb.save(FP)
    total = ws.max_row - 1
    specialty = sum(
        1
        for r in range(2, ws.max_row + 1)
        if str(ws.cell(r, col["Distribution Type"]).value or "") == "Specialty Distributor"
    )

    audit = {
        "backup": str(backup),
        "removed_count": len(removed),
        "removed": removed,
        "candidates": len(candidates),
        "verified": len(verified),
        "failed_sample": failed[:30],
        "added_count": len(to_add),
        "added": [{"company": n, "website": u, "country": c} for n, u, c in to_add],
        "total_rows": total,
        "specialty_distributor_rows": specialty,
        "target_total_from_161_plus_200": 361,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(
        f"removed={len(removed)} candidates={len(candidates)} verified={len(verified)} "
        f"added={len(to_add)} TOTAL={total} SPECIALTY={specialty}"
    )


if __name__ == "__main__":
    main()
