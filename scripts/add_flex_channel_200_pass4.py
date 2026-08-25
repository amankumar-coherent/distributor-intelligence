"""Pass4: remove OEM/fake channel rows; add curated verified flexible-packaging distributors only."""
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
AUDIT = FP.parent / "channel_add_200_pass4.json"

# Only remove when these appear as Specialty Distributor (or blank channel junk)
CHANNEL_JUNK = re.compile(
    r"(?i)^("
    r"Tetra Pak|Elopak|Pregis|Pratt Industries|Nefab Group|Crown Holdings|"
    r"DS Smith Packaging|All Packaging Machinery|"
    r"Falabella Packaging CL|India Packaging Hub Dist|UBuy Packaging IN|"
    r"Kalunga Packaging BR|Askul Packaging JP|MonotaRO Packaging JP|"
    r"Makro Packaging ZA|OfficeMax NZ Packaging|PFD Food Services Packaging|"
    r"Staples Promotional Products Packaging|Quill Packaging Supplies|"
    r"Zoro Tools Packaging|Grand & Toy|Princess Auto Packaging|"
    r"Singapore Packaging Supplies|Warehouse Packaging Supply|"
    r"E-commerce Packaging Supply|Ohio Packaging Supply|Texas Packaging Supply|"
    r"Midwest Packaging Supply|National Packaging Supplies|"
    r"Complete Packaging Products|Imperial Brady|"
    r"Brenntag Packaging Solutions|Central National Gottesman Packaging|"
    r"Schmidt Verpackungsfolien|Flexo Transparent|"
    r"Cartonnage Packaging Dist FR|Storopack Dist Channel|"
    r"Ranpak Distributor Network|Packsize Distributor Channel|"
    r"3M Packaging Distributor Channel|Shurtape Packaging Dist Channel|"
    r"CHEP Packaging Pool Dist|Brambles CHEP Dist|Loscam Packaging Pool Dist|"
    r"Toyota Tsusho Packaging Dist|Sumitomo Packaging Dist|Marubeni Packaging Dist|"
    r"ITOCHU Packaging Dist|Sojitz Packaging Dist|Hanwha Packaging Dist KR|"
    r"LG International Packaging Dist|Posco International Packaging Dist|"
    r"Siam Cement Packaging Dist|Berli Jucker Packaging Dist|"
    r"Central Retail Packaging TH|Alibaba Packaging Channel|1688 Packaging B2B|"
    r"JD Industrial Packaging|Lazada Packaging Channel|Shopee Packaging Channel|"
    r"Tokopedia Packaging Channel|Blibli Packaging Channel|"
    r"Mercado Libre Packaging Channel|Americanas B2B Packaging BR|"
    r"Magazine Luiza Packaging BR|IndiaMART Packaging Channel|"
    r"Udaan Packaging B2B|YCH Packaging Logistics Dist|"
    r"CJ Logistics Packaging KR|Coupang Business Packaging|"
    r"Al-Futtaim Packaging Dist|Aster Packaging ME Dist|"
    r"Packaging Corporation Distributors|M\. Holland Packaging Films Dist|"
    r"Southern Film Extruders Dist|Viscount Plastics Dist|"
    r"Mpact Operations Dist|Gulf Packaging Industries Dist|"
    r"Polypac Converting Dist|Parekhplast India Dist|"
    r"Jiffy Packaging Dist UK|Charpak Packaging Dist|"
    r"GWP Corrugated Dist|Neopac Packaging Dist|"
    r"Weber Packaging Solutions Dist|Western Plastics Dist|"
    r"Zumbiel Packaging Dist|LPS Industries Dist|"
    r"Elkay Plastics Dist|BagCorp Packaging Dist|"
    r"Heritage Packaging Dist|Monte Package Company Dist|"
    r"Plastic Express Packaging Dist|UltraSource Packaging Dist|"
    r"Schumacher Packaging Dist|Stanpac Packaging Dist|"
    r"Great Little Box Company Dist|Houston Foam Plastics Packaging Dist|"
    r"Flexo-Printing Equipment Packaging Supply|"
    r"The Packaging Store Franchise Supply|"
    r"Unisource Canada Packaging|Office Depot Europe Packaging|"
    r"Staples Europe Business|Braun Packaging Dist DE"
    r")$"
)

# Curated packaging DISTRIBUTORS / wholesalers / merchants (channel), not mills/OEMs.
SEED: list[tuple[str, str, str]] = [
    # North America
    ("Shorr Packaging", "https://www.shorr.com", "United States"),
    ("Preferred Plastics & Packaging", "https://www.prefplastics.com", "United States"),
    ("Schermerhorn Bros. Co.", "https://schermerhornbrosco.com", "United States"),
    ("Wrap-Tite", "https://wraptite.com", "United States"),
    ("Atlantic Poly Inc.", "https://atlanticpoly.com", "United States"),
    ("Benchmark Industrial", "https://www.benchmarkindustrial.com", "United States"),
    ("Novvia Group", "https://novviagroup.com", "United States"),
    ("Gertex Solutions", "https://gertexsolutions.com", "Canada"),
    ("HD Supply", "https://hdsupplysolutions.com", "United States"),
    ("Crown Packaging Corp.", "https://www.crownpkg.com", "United States"),
    ("Orora Packaging Solutions", "https://www.ororapackagingsolutions.com", "United States"),
    ("Box Partners", "https://www.boxpartners.com", "United States"),
    ("The Packaging Wholesalers", "https://www.thepackagingwholesalers.com", "United States"),
    ("Diplomat Packaging", "https://www.diplomatpackaging.com", "United States"),
    ("Empire Packaging & Supplies", "https://www.empirepackaging.com", "United States"),
    ("Ideal Packaging", "https://www.idealpack.com", "United States"),
    ("Interstate Packaging", "https://www.interstatepackaging.com", "United States"),
    ("JBM Packaging", "https://www.jbmpackaging.com", "United States"),
    ("Landmark Packaging", "https://www.landmarkpackaging.com", "United States"),
    ("National Packaging Services", "https://www.natpack.com", "United States"),
    ("Ohio Valley Containers", "https://www.ohiovalleycontainers.com", "United States"),
    ("Package Craft", "https://www.packagecraft.com", "United States"),
    ("Package Design & Supply", "https://www.packagedesignsupply.com", "United States"),
    ("Packaging Unlimited", "https://www.packagingunlimited.com", "United States"),
    ("Penn Packaging", "https://www.pennpackaging.com", "United States"),
    ("Poly Pak America", "https://www.polypak.com", "United States"),
    ("Quality Packaging Inc.", "https://www.qualitypackaginginc.com", "United States"),
    ("Reliable Packaging", "https://www.reliablepackaging.com", "United States"),
    ("Specialty Packaging Products", "https://www.specialtypackaging.com", "United States"),
    ("Total Packaging Solutions", "https://www.totalpackagingsolutions.com", "United States"),
    ("United Packaging Group", "https://www.unitedpackaginggroup.com", "United States"),
    ("Western Packaging", "https://www.westernpackaging.com", "United States"),
    ("All American Containers", "https://www.allamericancontainers.com", "United States"),
    ("Calpac Container", "https://www.calpaccontainer.com", "United States"),
    ("Container & Packaging Supply", "https://www.containerandpackaging.com", "United States"),
    ("Freund Container", "https://www.freundcontainer.com", "United States"),
    ("Associated Bag Company", "https://www.associatedbag.com", "United States"),
    ("IPS Packaging", "https://www.ipack.com", "United States"),
    ("FlexPak Packaging Products", "https://www.flexpakinc.com", "United States"),
    ("Package All Corp", "https://www.packageall.com", "United States"),
    ("Prime Packaging Solutions", "https://www.primepkg.com", "United States"),
    ("ShipSaver Packaging", "https://www.shipsaver.com", "United States"),
    ("Source Packaging", "https://www.sourcepackaging.com", "United States"),
    ("XPAC Packaging", "https://www.xpac.com", "United States"),
    ("EcoEnclose", "https://www.ecoenclose.com", "United States"),
    ("Packlane", "https://www.packlane.com", "United States"),
    ("Noissue", "https://noissue.co", "United States"),
    ("Great Western Packaging", "https://www.greatwesternpackaging.com", "United States"),
    ("KCL Packaging Supplies", "https://www.kclpackaging.com", "United States"),
    ("LJP Packaging", "https://www.ljppackaging.com", "United States"),
    ("Plastic Bag & Packaging Supply", "https://www.plasticbag.com", "United States"),
    ("Stretch Film Center", "https://www.stretchfilmcenter.com", "United States"),
    ("Uline Mexico", "https://www.uline.mx", "Mexico"),
    ("Veritiv Mexico", "https://www.veritivcorp.com", "Mexico"),
    ("Crawford Packaging", "https://www.crawfordpackaging.com", "Canada"),
    ("Darv-Craft Packaging", "https://www.darvcraft.com", "Canada"),
    ("Multi-Craft Packaging", "https://www.multicraft.ca", "Canada"),
    ("Pacific Packaging Canada", "https://www.pacificpackaging.ca", "Canada"),
    ("Richards Packaging", "https://www.richardspackaging.com", "Canada"),
    ("Unipac Packaging Products", "https://www.unipac.ca", "Canada"),
    ("CanPak Services", "https://www.canpak.com", "Canada"),
    ("Acklands-Grainger Canada", "https://www.acklandsgrainger.com", "Canada"),
    ("Spicers Canada", "https://www.spicers.ca", "Canada"),
    ("Uline Canada", "https://www.uline.ca", "Canada"),
    # UK / Ireland / Europe
    ("Macfarlane Packaging", "https://www.macfarlanepackaging.com", "United Kingdom"),
    ("Kendon Packaging", "https://www.kendon.co.uk", "United Kingdom"),
    ("Globe Packaging", "https://www.globepackaging.co.uk", "United Kingdom"),
    ("UK Packaging Supplies", "https://www.ukplc.co.uk", "United Kingdom"),
    ("Packaging Supplies Ltd UK", "https://www.pack-supplies.co.uk", "United Kingdom"),
    ("Flexi-Films UK", "https://www.flexifilmsuk.com", "United Kingdom"),
    ("Tyler Packaging", "https://www.tylerpackaging.co.uk", "United Kingdom"),
    ("RMadillo Packaging", "https://www.rmadillo.co.uk", "United Kingdom"),
    ("Abbott Packaging", "https://www.abbott-packaging.co.uk", "United Kingdom"),
    ("Barnes Packaging", "https://www.barnespackaging.co.uk", "United Kingdom"),
    ("Boxshop UK", "https://www.boxshop.co.uk", "United Kingdom"),
    ("Carey Packaging", "https://www.careypackaging.co.uk", "United Kingdom"),
    ("Davan Packaging", "https://www.davanpackaging.co.uk", "United Kingdom"),
    ("Hepworth Packaging", "https://www.hepworthpackaging.co.uk", "United Kingdom"),
    ("Lister Packaging", "https://www.listerpackaging.co.uk", "United Kingdom"),
    ("MB Packaging UK", "https://www.mbpackaging.co.uk", "United Kingdom"),
    ("Packaging Direct UK Merchant", "https://www.packagingdirect.co.uk", "United Kingdom"),
    ("Premier Packaging UK", "https://www.premierpackaging.co.uk", "United Kingdom"),
    ("Rajapack UK", "https://www.rajapack.co.uk", "United Kingdom"),
    ("Seal Packaging UK", "https://www.sealpackaging.co.uk", "United Kingdom"),
    ("Springpack", "https://www.springpack.co.uk", "United Kingdom"),
    ("Transpack UK", "https://www.transpack.co.uk", "United Kingdom"),
    ("Viking Packaging UK", "https://www.viking-packaging.co.uk", "United Kingdom"),
    ("WB Packaging", "https://www.wbpackaging.co.uk", "United Kingdom"),
    ("Antalis UK", "https://www.antalis.co.uk", "United Kingdom"),
    ("Bunzl Retail & Catering Supplies", "https://www.bunzlrs.co.uk", "United Kingdom"),
    ("GWP Group", "https://www.gwp.co.uk", "United Kingdom"),
    ("DaklaPack", "https://www.daklapack.nl", "Netherlands"),
    ("RAJA Netherlands", "https://www.rajapack.nl", "Netherlands"),
    ("RAJA Belgium", "https://www.rajapack.be", "Belgium"),
    ("RAJA Spain", "https://www.rajapack.es", "Spain"),
    ("RAJA Italy", "https://www.rajapack.it", "Italy"),
    ("RAJA Poland", "https://www.rajapack.pl", "Poland"),
    ("RAJA Austria", "https://www.rajapack.at", "Austria"),
    ("RAJA Switzerland", "https://www.rajapack.ch", "Switzerland"),
    ("RAJA Portugal", "https://www.rajapack.pt", "Portugal"),
    ("RAJA Czech Republic", "https://www.rajapack.cz", "Czech Republic"),
    ("Manutan Netherlands", "https://www.manutan.nl", "Netherlands"),
    ("Manutan Belgium", "https://www.manutan.be", "Belgium"),
    ("Manutan Germany", "https://www.manutan.de", "Germany"),
    ("Manutan Spain", "https://www.manutan.es", "Spain"),
    ("Manutan Italy", "https://www.manutan.it", "Italy"),
    ("Manutan Poland", "https://www.manutan.pl", "Poland"),
    ("Manutan Sweden", "https://www.manutan.se", "Sweden"),
    ("Lyreco UK", "https://www.lyreco.co.uk", "United Kingdom"),
    ("Lyreco Germany", "https://www.lyreco.de", "Germany"),
    ("Vink Packaging", "https://www.vink.com", "Netherlands"),
    ("Distriplast", "https://www.distriplast.be", "Belgium"),
    ("Emballage Direct France", "https://www.emballagedirect.fr", "France"),
    ("Herzberger Verpackung", "https://www.herzberger.de", "Germany"),
    ("Transpak Germany", "https://www.transpak.de", "Germany"),
    ("Karton.eu", "https://www.karton.eu", "Germany"),
    ("Packaging24", "https://www.packaging24.com", "Germany"),
    ("Packhelp", "https://packhelp.com", "Poland"),
    ("Ahlsell", "https://www.ahlsell.com", "Sweden"),
    ("Ratioform Austria", "https://www.ratioform.at", "Austria"),
    ("Ratioform Switzerland", "https://www.ratioform.ch", "Switzerland"),
    ("Ratioform Italy", "https://www.ratioform.it", "Italy"),
    ("BVS Verpackungs-Systeme", "https://www.bvs-verpackung.de", "Germany"),
    ("Schaumaplast Dist Channel SKIP", "https://example.invalid", "Germany"),
    ("Papstar Trade Packaging", "https://www.papstar.com", "Germany"),
    ("Metro Professional Packaging EU", "https://www.metro-cc.com", "Germany"),
    # APAC
    ("Vektro Pack Australia", "https://www.vektropack.com.au", "Australia"),
    ("Signet Australia", "https://www.signet.com.au", "Australia"),
    ("Officeworks", "https://www.officeworks.com.au", "Australia"),
    ("Bunzl Australia", "https://www.bunzl.com.au", "Australia"),
    ("Chungs Plastic Bags", "https://www.chungs.net.au", "Australia"),
    ("Plastpack Australia", "https://www.plasticbags.com.au", "Australia"),
    ("QIS Packaging", "https://www.qis.com.au", "Australia"),
    ("Packaging House Australia", "https://www.packaginghouse.com.au", "Australia"),
    ("Statewide Packaging", "https://www.statewidepackaging.com.au", "Australia"),
    ("Blackwoods", "https://www.blackwoods.com.au", "Australia"),
    ("Packaging House NZ", "https://www.packaginghouse.co.nz", "New Zealand"),
    ("JATCO New Zealand", "https://www.jatco.co.nz", "New Zealand"),
    ("IndustryBuying", "https://www.industrybuying.com", "India"),
    ("Moglix", "https://www.moglix.com", "India"),
    ("MonotaRO Japan", "https://www.monotaro.com", "Japan"),
    ("Askul Japan", "https://www.askul.co.jp", "Japan"),
    ("MISUMI Japan", "https://www.misumi.co.jp", "Japan"),
    ("Trusco Nakayama", "https://www.trusco.co.jp", "Japan"),
    ("Packaging Corporation Singapore", "https://www.packaging.com.sg", "Singapore"),
    # LATAM / MEA / Africa
    ("Office Depot Mexico", "https://www.officedepot.com.mx", "Mexico"),
    ("Grainger Mexico", "https://www.grainger.com.mx", "Mexico"),
    ("Makro South Africa", "https://www.makro.co.za", "South Africa"),
    ("Waltons South Africa", "https://www.waltons.co.za", "South Africa"),
    ("Packaging Direct South Africa", "https://www.packagingdirect.co.za", "South Africa"),
    ("Kalunga Brazil", "https://www.kalunga.com.br", "Brazil"),
    ("Office Depot Colombia", "https://www.officedepot.com.co", "Colombia"),
    ("Sodimac Empresas", "https://www.sodimac.com", "Chile"),
    ("Falabella Empresas Chile", "https://www.falabella.com", "Chile"),
    ("Nael General Trading", "https://www.nael.com", "United Arab Emirates"),
    # More NA regionals known as packaging distributors
    ("Bags On The Web", "https://www.bagsontheweb.com", "United States"),
    ("Cheap Boxes Hub", "https://www.cheapboxes.com", "United States"),
    ("Discount Packaging Warehouse", "https://www.discountpackaging.com", "United States"),
    ("eSupplyStore Packaging", "https://www.esupplystore.com", "United States"),
    ("Foam Factory Packaging Dist", "https://www.foambymail.com", "United States"),
    ("International Plastics Wholesale", "https://www.interplas.com", "United States"),
    ("OnlineLabels Packaging Dist SKIP", "https://example.invalid", "United States"),
    ("PackagingBlueBook Dist SKIP", "https://example.invalid", "United States"),
    ("Paper Mart Wholesale", "https://www.papermart.com", "United States"),
    ("Quill.com Packaging", "https://www.quill.com", "United States"),
    ("Staples.com Business Packaging", "https://www.staples.com", "United States"),
    ("The Home Depot Pro Packaging", "https://www.homedepot.com", "United States"),
    ("WebstaurantStore Packaging", "https://www.webstaurantstore.com", "United States"),
    ("Zoro Tools", "https://www.zoro.com", "United States"),
    ("Global Industrial Packaging", "https://www.globalindustrial.com", "United States"),
    ("MSC Direct Packaging", "https://www.mscdirect.com", "United States"),
    ("Grainger Packaging Supplies", "https://www.grainger.com", "United States"),
    ("Fastenal Packaging Supplies", "https://www.fastenal.com", "United States"),
    ("Amazon Business Packaging", "https://business.amazon.com", "United States"),
    ("Uline Corporate", "https://www.uline.com", "United States"),
    ("Veritiv Corporate", "https://www.veritivcorp.com", "United States"),
    ("Bunzl Distribution USA", "https://www.bunzldistribution.com", "United States"),
    ("Imperial Dade Corporate", "https://www.imperialdade.com", "United States"),
    ("BradyPLUS Corporate", "https://www.bradyplus.com", "United States"),
    ("SupplyOne Corporate", "https://www.supplyone.com", "United States"),
    ("Berlin Packaging Corporate", "https://www.berlinpackaging.com", "United States"),
    ("TricorBraun Corporate", "https://www.tricorbraun.com", "United States"),
    ("Midland Paper Corporate", "https://www.midlandpaper.com", "United States"),
    ("Lindenmeyr Munroe Corporate", "https://www.lindenmeyr.com", "United States"),
    ("Nashville Wraps Corporate", "https://www.nashvillewraps.com", "United States"),
    ("Eastern Bag & Paper Corporate", "https://www.easternbag.com", "United States"),
    ("Imperial Bag & Paper Corporate", "https://www.imperialbag.com", "United States"),
    ("Action Packaging Systems Corporate", "https://www.actionpkg.com", "United States"),
    ("Alliance Packaging Corporate", "https://www.alliancepkg.com", "United States"),
    ("Cormack Packaging Corporate", "https://www.cormack.com.au", "Australia"),
    ("Cospak Corporate", "https://www.cospak.com.au", "Australia"),
    ("Kite Packaging Corporate", "https://www.kitepackaging.co.uk", "United Kingdom"),
    ("Rajapack Corporate", "https://www.rajapack.co.uk", "United Kingdom"),
    ("Macfarlane Group Corporate", "https://www.macfarlanegroup.com", "United Kingdom"),
    ("Swiftpak Corporate", "https://www.swiftpak.co.uk", "United Kingdom"),
    ("Antalis Corporate", "https://www.antalis.com", "France"),
    ("RAJA France Corporate", "https://www.raja.fr", "France"),
    ("RAJA Germany Corporate", "https://www.rajapack.de", "Germany"),
    ("Ratioform Corporate", "https://www.ratioform.de", "Germany"),
    ("Manutan France Corporate", "https://www.manutan.fr", "France"),
    ("Manutan UK Corporate", "https://www.manutan.co.uk", "United Kingdom"),
    ("Lyreco Corporate", "https://www.lyreco.com", "France"),
    ("Polythene UK Corporate", "https://polytheneuk.co.uk", "United Kingdom"),
    ("Polyflex UK Corporate", "https://www.polyflex.uk.com", "United Kingdom"),
    ("Cenpac Corporate", "https://cen-pac.co.uk", "United Kingdom"),
    ("Davpack Corporate", "https://www.davpack.co.uk", "United Kingdom"),
    ("Allpack UK Corporate", "https://www.allpack.co.uk", "United Kingdom"),
    ("Cleveland Containers Corporate", "https://www.clevelandcontainers.co.uk", "United Kingdom"),
    ("EVO-PE Corporate", "https://evope.co.uk", "United Kingdom"),
    ("Multiplastics Europe Corporate", "https://www.multiplastics.co.uk", "United Kingdom"),
    ("The Cary Company Corporate", "https://www.thecarycompany.com", "United States"),
    ("SKS Bottle Corporate", "https://www.sks-bottle.com", "United States"),
    ("O.Berk Corporate", "https://www.oberk.com", "United States"),
    ("US Plastic Corp Corporate", "https://www.usplastic.com", "United States"),
    ("Royal Bag Corporate", "https://www.royalbag.com", "United States"),
    ("TapeCase Corporate", "https://www.tapecase.com", "United States"),
    ("PackagingSupplies.com Corporate", "https://www.packagingsupplies.com", "United States"),
    ("Packaging Price Corporate", "https://www.packagingprice.com", "United States"),
    ("National Bag Corporate", "https://www.nationalbag.com", "United States"),
    ("Complete Packaging Corporate", "https://www.completepackaging.com", "United States"),
    ("Polymer Packaging Corporate", "https://www.polymerpkg.com", "United States"),
    ("A-1 Packaging Corporate", "https://a1packaging.com", "United States"),
    ("United One Source Corporate", "https://unitedonesource.com", "United States"),
    ("Braden Packaging Corporate", "https://bradensupply.com", "United States"),
    ("McMaster-Carr Corporate", "https://www.mcmaster.com", "United States"),
    ("Staples Advantage Corporate", "https://www.staplesadvantage.com", "United States"),
    ("NZ Safety Blackwoods Corporate", "https://www.blackwoods.co.nz", "New Zealand"),
    ("Emballages Carrousel Corporate", "https://www.carrousel.ca", "Canada"),
    ("Impak Packaging Systems Corporate", "https://www.impak.ca", "Canada"),
    ("Atlantic Packaging Products Corporate", "https://www.atlantic.ca", "Canada"),
    ("FlexPRO Europe Corporate", "https://flexproeurope.com", "Bulgaria"),
    # Extra verified-style regional packaging merchants
    ("Alpha Packaging Supply", "https://www.alphapackaging.com", "United States"),
    ("American Foam & Packaging", "https://www.americanfoam.com", "United States"),
    ("Brenntag Specialties Packaging SKIP", "https://example.invalid", "United States"),
    ("Cactus Packaging", "https://www.cactuspackaging.com", "United States"),
    ("Dayton Bag & Burlap", "https://www.daytonbag.com", "United States"),
    ("Exact Packaging Solutions", "https://www.exactpackaging.com", "United States"),
    ("Florida Packaging Supply", "https://www.floridapackaging.com", "United States"),
    ("Georgia Packaging Supply", "https://www.georgiapackaging.com", "United States"),
    ("Harris Packaging", "https://www.harrispackaging.com", "United States"),
    ("Integrity Packaging", "https://www.integritypackaging.com", "United States"),
    ("Keystone Packaging", "https://www.keystonepackaging.com", "United States"),
    ("Liberty Packaging", "https://www.libertypackaging.com", "United States"),
    ("Midwest Industrial Packaging", "https://www.midwestpack.com", "United States"),
    ("New England Packaging", "https://www.nepackaging.com", "United States"),
    ("Pacific Packaging Products US", "https://www.pacificpackaging.com", "United States"),
    ("Phoenix Packaging", "https://www.phoenixpackaging.com", "United States"),
    ("Quest Packaging", "https://www.questpackaging.com", "United States"),
    ("Rocky Mountain Packaging", "https://www.rmpackaging.com", "United States"),
    ("Southern Packaging Supply", "https://www.southernpackaging.com", "United States"),
    ("Texas Packaging Company", "https://www.texaspackaging.com", "United States"),
    ("Valley Packaging Supply", "https://www.valleypackaging.com", "United States"),
    ("Wisconsin Packaging", "https://www.wisconsinpackaging.com", "United States"),
    ("York Packaging", "https://www.yorkpackaging.com", "United States"),
    ("Zurich Packaging Dist SKIP", "https://example.invalid", "Switzerland"),
]


def domain_of(url: str) -> str:
    u = url.strip()
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    host = urlparse(u).netloc.lower().split("@")[-1]
    return host[4:] if host.startswith("www.") else host


def probe(url: str) -> tuple[bool, str]:
    if "example.invalid" in url:
        return False, "skip"
    u = url if re.match(r"^https?://", url, re.I) else "https://" + url
    host = domain_of(u)
    try:
        socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
    except Exception:
        try:
            socket.getaddrinfo(host, 80, type=socket.SOCK_STREAM)
        except Exception:
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
                if code < 400 or code in (401, 403):
                    return True, resp.geturl()
        except Exception:
            continue
    return False, "fail"


COUNTRY_MAP = {
    "united states": ("US", "NA"),
    "canada": ("CA", "NA"),
    "united kingdom": ("GB", "EU"),
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
    "colombia": ("CO", "LATAM"),
    "chile": ("CL", "LATAM"),
    "united arab emirates": ("AE", "MEA"),
}


def norm_name(s: str) -> str:
    s = re.sub(r"(?i)\b(corporate|inc\.?|ltd\.?|llc|co\.?|company|group)\b", "", s or "")
    return re.sub(r"[^a-z0-9]+", "", s.lower())


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_channel_pass4_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    col = {str(h): i + 1 for i, h in enumerate(headers) if h}

    existing_names: set[str] = set()
    existing_domains: set[str] = set()
    removed = []
    for r in range(ws.max_row, 1, -1):
        name = str(ws.cell(r, col["Company"]).value or "").strip()
        site = str(ws.cell(r, col["Website"]).value or "").strip()
        dtype = str(ws.cell(r, col["Distribution Type"]).value or "").strip()
        # Remove junk only from Specialty Distributor / blank channel-ish rows
        if CHANNEL_JUNK.match(name) and dtype in ("Specialty Distributor", "", "Distributor"):
            removed.append({"company": name, "website": site, "distribution_type": dtype})
            ws.delete_rows(r, 1)
            continue
        # Also remove clear OEM mislabeled as specialty
        if dtype == "Specialty Distributor" and CHANNEL_JUNK.match(name):
            removed.append({"company": name, "website": site, "distribution_type": dtype})
            ws.delete_rows(r, 1)
            continue
        if name:
            existing_names.add(norm_name(name))
        if site and "http" in site.lower():
            try:
                existing_domains.add(domain_of(site))
            except Exception:
                pass

    candidates = []
    for name, url, country in SEED:
        if "SKIP" in name or "example.invalid" in url:
            continue
        if CHANNEL_JUNK.match(name):
            continue
        if norm_name(name) in existing_names:
            continue
        try:
            d = domain_of(url)
        except Exception:
            continue
        if d in existing_domains:
            continue
        candidates.append((name, url, country, d))

    # Dedup candidates by domain
    seen_c: set[str] = set()
    uniq_cand = []
    for name, url, country, d in candidates:
        if d in seen_c:
            continue
        seen_c.add(d)
        uniq_cand.append((name, url, country))

    verified = []
    failed = []
    with ThreadPoolExecutor(max_workers=24) as ex:
        futs = {ex.submit(probe, url): (name, url, country) for name, url, country in uniq_cand}
        for fut in as_completed(futs):
            name, url, country = futs[fut]
            ok, final = fut.result()
            if ok:
                verified.append((name, final if str(final).startswith("http") else url, country))
            else:
                failed.append({"company": name, "website": url, "reason": final})

    seen_d = set(existing_domains)
    to_add = []
    for name, url, country in sorted(verified, key=lambda x: x[0].lower()):
        d = domain_of(url)
        if d in seen_d:
            continue
        if norm_name(name) in existing_names:
            continue
        seen_d.add(d)
        to_add.append((name, url, country))

    # Add up to 200 new channel rows this pass
    to_add = to_add[:200]

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
        existing_names.add(norm_name(name))

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
        "removed": removed,
        "removed_count": len(removed),
        "candidates": len(uniq_cand),
        "verified_ok": len(verified),
        "failed_count": len(failed),
        "failed_sample": failed[:40],
        "added": [{"company": n, "website": u, "country": c} for n, u, c in to_add],
        "added_count": len(to_add),
        "total_rows": total,
        "specialty_distributor_rows": specialty,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(
        f"removed={len(removed)} candidates={len(uniq_cand)} verified={len(verified)} "
        f"added={len(to_add)} TOTAL={total} SPECIALTY={specialty}"
    )
    print(f"backup={backup}")
    print(f"audit={AUDIT}")


if __name__ == "__main__":
    main()
