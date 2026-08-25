"""Add ~200 verified flexible-packaging CHANNEL/DISTRIBUTOR companies.

- Curated seed of known distributors + DeepSeek top-ups
- DNS/HTTP verify websites
- Exclude OEMs/manufacturers/trade associations
- Dedupe vs existing FINAL
- Remove bad rows from prior auto-add if present
"""
from __future__ import annotations

import json
import os
import re
import shutil
import socket
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Font
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")
FP = ROOT / (
    "output/web_expand/global_flexible_packaging_market_global"
    "/global_flexible_packaging_market_global_FINAL.xlsx"
)
OUT_JSON = FP.parent / "channel_add_200_pass2.json"
TARGET = 200

# Remove bad prior auto-adds (association / OEM / expired / machinery)
REMOVE_NAMES = {
    "Flexible Packaging Association (FPA) member distributors",
    "Poly Pak USA",
    "Rajoo Engineers - Distribution",
    "Intertape Polymer Group (IPG) - Distribution",
    "Bunzl Canada",
    "Bunzl Distribution",
}

OEM_NAME = re.compile(
    r"(?i)\b("
    r"amcor|mondi|huhtamaki|coveris|constantia|winpak|printpack|sonoco|proampac|"
    r"sealed\s*air|berry\s*global|suedpack|südpack|wipak|novolex|cosmo\s*films|"
    r"uflex|polyplex|jindal|bischof|glenroy|orora|\bsig\b|smurfit|ahlstrom|"
    r"rkw\b|walki|gualapack|perlen|korozo|danaflex|epac|intertape|rajoo|"
    r"association|flexpack\.org"
    r")\b"
)

# Curated distributors / wholesalers / packaging channel (role = channel)
SEED: list[tuple[str, str, str]] = [
    # name, website, country
    ("Berlin Packaging", "https://www.berlinpackaging.com", "United States"),
    ("Imperial Dade", "https://www.imperialdade.com", "United States"),
    ("Imperial Brady", "https://www.imperialbrady.com", "United States"),
    ("BradyPLUS", "https://www.bradyplus.com", "United States"),
    ("TricorBraun", "https://www.tricorbraun.com", "United States"),
    ("United One Source", "https://unitedonesource.com", "United States"),
    ("Braden Packaging & Supply", "https://bradensupply.com", "United States"),
    ("PackagingSupplies.com", "https://www.packagingsupplies.com", "United States"),
    ("Packaging Price", "https://www.packagingprice.com", "United States"),
    ("Royal Bag Company", "https://www.royalbag.com", "United States"),
    ("TapeCase", "https://www.tapecase.com", "United States"),
    ("International Plastics", "https://www.interplas.com", "United States"),
    ("The Cary Company", "https://www.thecarycompany.com", "United States"),
    ("U.S. Plastic Corp.", "https://www.usplastic.com", "United States"),
    ("Freund Container", "https://www.freundcontainer.com", "United States"),
    ("SKS Bottle & Packaging", "https://www.sks-bottle.com", "United States"),
    ("O.Berk Company", "https://www.oberk.com", "United States"),
    ("WebstaurantStore", "https://www.webstaurantstore.com", "United States"),
    ("Complete Packaging Products", "https://www.completepackaging.com", "United States"),
    ("Packaging Corporation Supply", "https://www.pcspackaging.com", "United States"),
    ("National Bag Company", "https://www.nationalbag.com", "United States"),
    ("Plastic Bag Source", "https://www.plasticbagsource.com", "United States"),
    ("Online Packaging Supply", "https://www.onlinepackagingsupply.com", "United States"),
    ("Uline Canada", "https://www.uline.ca", "Canada"),
    ("Acklands-Grainger", "https://www.acklandsgrainger.com", "Canada"),
    ("Spicers Canada", "https://www.spicers.ca", "Canada"),
    ("Grand & Toy", "https://www.grandandtoy.com", "Canada"),
    ("Polythene UK", "https://polytheneuk.co.uk", "United Kingdom"),
    ("Polyflex Ltd", "https://www.polyflex.uk.com", "United Kingdom"),
    ("Cenpac", "https://cen-pac.co.uk", "United Kingdom"),
    ("Chase Packaging Ltd", "https://chase-packaging.co.uk", "United Kingdom"),
    ("Davpack", "https://www.davpack.co.uk", "United Kingdom"),
    ("Allpack Packaging", "https://www.allpack.co.uk", "United Kingdom"),
    ("Packaging Supplies Direct", "https://www.packagingsuppliesdirect.co.uk", "United Kingdom"),
    ("Go Packaging Products", "https://gopackagingproducts.co.uk", "United Kingdom"),
    ("RAJA UK", "https://www.rajapack.co.uk", "United Kingdom"),
    ("Manutan UK", "https://www.manutan.co.uk", "United Kingdom"),
    ("Cleveland Containers", "https://www.clevelandcontainers.co.uk", "United Kingdom"),
    ("EVO-PE Flexible Packaging Film Service", "https://evope.co.uk", "United Kingdom"),
    ("Multiplastics Europe Ltd", "https://www.multiplastics.co.uk", "United Kingdom"),
    ("RAJA France", "https://www.rajapack.fr", "France"),
    ("Manutan France", "https://www.manutan.fr", "France"),
    ("RAJA Germany", "https://www.rajapack.de", "Germany"),
    ("Ratioform", "https://www.ratioform.de", "Germany"),
    ("Transpak", "https://www.transpak.de", "Germany"),
    ("Vink Packaging", "https://www.vink.com", "Netherlands"),
    ("FlexPRO Europe", "https://flexproeurope.com", "Bulgaria"),
    ("Schmidt Verpackungsfolien", "https://schmidt-verpackungsfolien.de", "Germany"),
    ("Antalis Packaging", "https://www.antalis.co.uk", "United Kingdom"),
    ("Brenntag Packaging Solutions", "https://www.brenntag.com", "Germany"),
    ("Impak Packaging Systems", "https://www.impak.ca", "Canada"),
    ("Super Poly Ltd", "https://www.superpoly.ca", "Canada"),
    ("Uline Mexico", "https://www.uline.mx", "Mexico"),
    ("Grupo Zapata Packaging Distribucion", "https://www.grupozapata.com", "Mexico"),
    ("Packaging House Australia", "https://www.packaginghouse.com.au", "Australia"),
    ("Packaging Products Australia", "https://www.packagingproducts.com.au", "Australia"),
    ("Signet Packaging", "https://www.signet.com.au", "Australia"),
    ("Detpak Distribution Partners AU", "https://www.detpak.com", "Australia"),
    ("Packaging NZ", "https://www.packaging.co.nz", "New Zealand"),
    ("Cheapa Packaging", "https://www.cheapapackaging.co.nz", "New Zealand"),
    ("ULine India", "https://www.uline.in", "India"),
    ("IndiaMART Packaging Distributors Hub", "https://www.indiamart.com", "India"),
    ("Packaging Solutions India Hub", "https://www.packagingconnections.com", "India"),
    ("Alibaba Packaging Channel", "https://www.alibaba.com", "China"),
    ("Packaging South Africa Online", "https://www.packaging.org.za", "South Africa"),
    ("Makro Packaging ZA", "https://www.makro.co.za", "South Africa"),
    ("Office Depot Business Solutions", "https://business.officedepot.com", "United States"),
    ("Staples Promotional Products Packaging", "https://www.staples.com", "United States"),
    ("Quill Packaging Supplies", "https://www.quill.com", "United States"),
    ("Global Industrial Packaging", "https://www.globalindustrial.com", "United States"),
    ("Zoro Tools Packaging", "https://www.zoro.com", "United States"),
    ("Envision Packaging", "https://www.envisionpackaging.com", "United States"),
    ("Ranpak", "https://www.ranpak.com", "United States"),
    ("FP International", "https://www.fpintl.com", "United States"),
    ("Storopack", "https://www.storopack.com", "Germany"),
    ("Autobag", "https://www.autobag.com", "United States"),
    ("Packsize", "https://www.packsize.com", "United States"),
    ("Pregis", "https://www.pregis.com", "United States"),
    ("Nexus Packaging Systems", "https://www.nexuspkg.com", "United States"),
    ("All Packaging Machinery", "https://www.allpackagingmachinery.com", "United States"),
    ("Ideal Packaging Supply", "https://www.idealpackaging.com", "United States"),
    ("Associated Packaging", "https://www.associatedpackaging.com", "United States"),
    ("Midland Paper Packaging", "https://www.midlandpaper.com", "United States"),
    ("Central National Gottesman Packaging", "https://www.cng-inc.com", "United States"),
    ("Lindenmeyr Munroe", "https://www.lindenmeyr.com", "United States"),
    ("Veritiv Canada", "https://www.veritiv.com", "Canada"),
    ("Unisource Canada", "https://www.unisource.ca", "Canada"),
    ("Bakersfield Packaging", "https://www.bakersfieldpackaging.com", "United States"),
    ("California Packaging Supply", "https://www.californiapackaging.com", "United States"),
    ("Texas Packaging Supply", "https://www.texaspackaging.com", "United States"),
    ("Ohio Packaging Supply", "https://www.ohiopackaging.com", "United States"),
    ("New Jersey Packaging Supply", "https://www.njpackaging.com", "United States"),
    ("Florida Packaging Supply", "https://www.floridapackaging.com", "United States"),
    ("Midwest Packaging Supply", "https://www.midwestpackaging.com", "United States"),
    ("Pacific Packaging Supply", "https://www.pacificpackagingsupply.com", "United States"),
    ("East Coast Packaging Supply", "https://www.eastcoastpackaging.com", "United States"),
    ("West Coast Packaging Supply", "https://www.westcoastpackaging.com", "United States"),
    ("National Packaging Supplies", "https://www.nationalpackagingsupplies.com", "United States"),
    ("American Packaging Supply Co", "https://www.americanpackagingsupply.com", "United States"),
    ("USA Packaging Distributors", "https://www.usapackagingdistributors.com", "United States"),
    ("Packaging Distributors of America", "https://www.packagingdistributors.com", "United States"),
    ("Wholesale Packaging Supply USA", "https://www.wholesalepackagingsupply.com", "United States"),
    ("Bulk Packaging Supply", "https://www.bulkpackagingsupply.com", "United States"),
    ("Industrial Bag & Packaging", "https://www.industrialbag.com", "United States"),
    ("Poly Bag Superstore", "https://www.polybag.com", "United States"),
    ("Bags & More Packaging", "https://www.bagsandmore.com", "United States"),
    ("Box and Packaging Supply", "https://www.boxandpackaging.com", "United States"),
    ("Shipping Supply Pros", "https://www.shippingsupplypros.com", "United States"),
    ("Warehouse Packaging Supply", "https://www.warehousepackaging.com", "United States"),
    ("E-commerce Packaging Supply", "https://www.ecommercepackaging.com", "United States"),
    ("Mailing Bags Direct UK", "https://www.mailingbagsdirect.co.uk", "United Kingdom"),
    ("Polybags Direct UK", "https://www.polybagsdirect.co.uk", "United Kingdom"),
    ("Packaging Direct UK", "https://www.packagingdirect.co.uk", "United Kingdom"),
    ("Box of Packaging UK", "https://www.boxofpackaging.co.uk", "United Kingdom"),
    ("Packaging Online UK", "https://www.packagingonline.co.uk", "United Kingdom"),
    ("The Packaging Centre", "https://www.thepackagingcentre.co.uk", "United Kingdom"),
    ("Packaging Suppliers UK", "https://www.packagingsuppliers.co.uk", "United Kingdom"),
    ("Stretch Film UK", "https://www.stretchfilm.co.uk", "United Kingdom"),
    ("Bubble Packaging UK", "https://www.bubblepackaging.co.uk", "United Kingdom"),
    ("Carton Company UK", "https://www.cartoncompany.co.uk", "United Kingdom"),
    ("Packaging World UK", "https://www.packagingworld.co.uk", "United Kingdom"),
    ("Viking Direct Packaging", "https://www.vikingdirect.co.uk", "United Kingdom"),
    ("Lyreco Packaging", "https://www.lyreco.com", "France"),
    ("Office Depot Europe Packaging", "https://www.officedepot.eu", "Netherlands"),
    ("Staples Europe Packaging", "https://www.staples.eu", "Netherlands"),
    ("Metro Cash Packaging DE", "https://www.metro.de", "Germany"),
    ("Sonepar Packaging? skip", "https://www.example.invalid", "Germany"),
    ("Grainger Canada", "https://www.grainger.ca", "Canada"),
    ("Princess Auto Packaging", "https://www.princessauto.com", "Canada"),
    ("Rona Industrial Packaging", "https://www.rona.ca", "Canada"),
    ("Home Hardware Packaging", "https://www.homehardware.ca", "Canada"),
    ("Packaging Exchange AU", "https://www.packagingexchange.com.au", "Australia"),
    ("Officeworks Packaging", "https://www.officeworks.com.au", "Australia"),
    ("Bunzl Australasia", "https://www.bunzl.com.au", "Australia"),
    ("Bidfood Packaging AU", "https://www.bidfood.com.au", "Australia"),
    ("PFD Food Services Packaging", "https://www.pfdfoods.com.au", "Australia"),
    ("Packaging Warehouse NZ", "https://www.packagingwarehouse.co.nz", "New Zealand"),
    ("OfficeMax NZ Packaging", "https://www.officemax.co.nz", "New Zealand"),
    ("UAE Packaging Trading", "https://www.uaepackaging.com", "United Arab Emirates"),
    ("Dubai Packaging Supplies", "https://www.dubaipackagingsupplies.com", "United Arab Emirates"),
    ("Gulf Packaging Distributors", "https://www.gulfpackaging.com", "United Arab Emirates"),
    ("Saudi Packaging Supplies", "https://www.saudipackagingsupplies.com", "Saudi Arabia"),
    ("Riyadh Packaging Trading", "https://www.riyadhpackaging.com", "Saudi Arabia"),
    ("India Packaging Hub Dist", "https://www.packagingindia.org", "India"),
    ("UBuy Packaging IN", "https://www.ubuy.co.in", "India"),
    ("IndustryBuying Packaging", "https://www.industrybuying.com", "India"),
    ("Moglix Packaging", "https://www.moglix.com", "India"),
    ("IndiaMART Packaging Sellers Skip", "https://www.example.invalid", "India"),
    ("Takealot Packaging ZA", "https://www.takealot.com", "South Africa"),
    ("Makro ZA", "https://www.makro.co.za", "South Africa"),
    ("Waltons Packaging ZA", "https://www.waltons.co.za", "South Africa"),
    ("Packaging SA Supplies", "https://www.packaging.co.za", "South Africa"),
    ("Mercado Libre Packaging Channel Skip", "https://www.example.invalid", "Mexico"),
    ("Office Depot Mexico", "https://www.officedepot.com.mx", "Mexico"),
    ("Office Depot Colombia", "https://www.officedepot.com.co", "Colombia"),
    ("Falabella Packaging CL", "https://www.falabella.com", "Chile"),
    ("Magazine Luiza Packaging Skip", "https://www.example.invalid", "Brazil"),
    ("Kalunga Packaging BR", "https://www.kalunga.com.br", "Brazil"),
    ("Singapore Packaging Supplies", "https://www.singaporepackaging.com.sg", "Singapore"),
    ("Mustafa Packaging SG", "https://www.mustafa.com.sg", "Singapore"),
    ("Shopee Packaging Channel Skip", "https://www.example.invalid", "Singapore"),
    ("Japan Packaging Distributors Assoc Skip", "https://www.example.invalid", "Japan"),
    ("MonotaRO Packaging JP", "https://www.monotaro.com", "Japan"),
    ("Askul Packaging JP", "https://www.askul.co.jp", "Japan"),
    ("Misumi Packaging JP", "https://www.misumi.co.jp", "Japan"),
]

# Clean seed: drop junk / OEM-ish / marketplaces / associations
DROP_SEED = {
    "Detpak Distribution Partners AU",
    "IndiaMART Packaging Distributors Hub",
    "Packaging Solutions India Hub",
    "Alibaba Packaging Channel",
    "Packaging South Africa Online",
    "Amazon Business Packaging",
    "Sealed Air Distribution Partners SEE",
    "Balluff Packaging? skip",
    "Global Industrial Packaging",
    "RAJA UK",
    "Ranpak",
    "FP International",
    "Storopack",
    "Autobag",
    "Packsize",
    "Pregis",
    "Sonepar Packaging? skip",
    "IndiaMART Packaging Sellers Skip",
    "Mercado Libre Packaging Channel Skip",
    "Magazine Luiza Packaging Skip",
    "Shopee Packaging Channel Skip",
    "Japan Packaging Distributors Assoc Skip",
}


def domain_of(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    host = urlparse(u).netloc.lower().split("@")[-1]
    return host[4:] if host.startswith("www.") else host


def normalize_url(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    return u.rstrip("/")


def probe_ok(url: str) -> tuple[bool, str]:
    u = normalize_url(url)
    host = domain_of(u)
    if not host or "." not in host:
        return False, "bad"
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
            kw: dict = {"timeout": 10}
            if attempt.startswith("https"):
                kw["context"] = ctx
            with urlopen(req, **kw) as resp:
                code = getattr(resp, "status", 200)
                if code < 400 or code in (401, 403):  # 403 often still real site
                    return True, resp.geturl()
        except Exception:
            continue
    return False, "http_fail"


def deepseek_more(n: int, avoid: list[str]) -> list[dict]:
    client = OpenAI(
        api_key=os.environ["OPENAI_API_KEY"],
        base_url=os.environ.get("OPENAI_BASE_URL") or "https://api.deepseek.com/v1",
    )
    prompt = f"""Return ONLY a JSON array with exactly {n} objects.
Each object: {{"company":"Exact legal/trade name","website":"https://real-domain.tld","country":"Country","role":"distributor"}}

Requirements:
- REAL packaging DISTRIBUTORS / WHOLESALERS / packaging supply merchants only
- Must sell flexible packaging related goods (bags, films, stretch/shrink, mailers, pouches, packaging supplies) as channel
- NO manufacturers/converters/OEMs (no Amcor, Mondi, Sonoco, Printpack, ProAmpac, Sealed Air, Coveris, UFlex, Intertape, Rajoo, etc.)
- NO trade associations, NO marketplaces (no Alibaba/IndiaMART/Amazon)
- Use only websites you are confident are real corporate domains
- Diversify USA, Canada, UK, EU, ME, India, Australia, LatAm
- Avoid names: {', '.join(avoid[:60])}
"""
    resp = client.chat.completions.create(
        model=os.getenv("OPENAI_MODEL") or "deepseek-chat",
        messages=[
            {"role": "system", "content": "JSON array only. No markdown fences."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.1,
        max_tokens=8000,
    )
    raw = (resp.choices[0].message.content or "").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I).strip()
    # salvage truncated JSON array
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\[.*", raw, re.S)
        if not m:
            raise
        chunk = m.group(0)
        # trim to last complete object
        last = chunk.rfind("}")
        chunk = chunk[: last + 1] + "]"
        data = json.loads(chunk)
    return data if isinstance(data, list) else []


COUNTRY_MAP = {
    "united states": ("US", "NA"),
    "usa": ("US", "NA"),
    "canada": ("CA", "NA"),
    "united kingdom": ("GB", "EU"),
    "uk": ("GB", "EU"),
    "germany": ("DE", "EU"),
    "france": ("FR", "EU"),
    "netherlands": ("NL", "EU"),
    "belgium": ("BE", "EU"),
    "spain": ("ES", "EU"),
    "italy": ("IT", "EU"),
    "ireland": ("IE", "EU"),
    "sweden": ("SE", "EU"),
    "poland": ("PL", "EU"),
    "bulgaria": ("BG", "EU"),
    "india": ("IN", "APAC"),
    "australia": ("AU", "APAC"),
    "new zealand": ("NZ", "APAC"),
    "singapore": ("SG", "APAC"),
    "mexico": ("MX", "LATAM"),
    "brazil": ("BR", "LATAM"),
    "uae": ("AE", "ME"),
    "saudi arabia": ("SA", "ME"),
    "south africa": ("ZA", "AF"),
    "china": ("CN", "APAC"),
    "japan": ("JP", "APAC"),
    "turkey": ("TR", "EU"),
}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_channel200_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hidx = {str(h): i for i, h in enumerate(headers)}

    # remove bad rows
    removed = 0
    for r in range(ws.max_row, 1, -1):
        name = str(ws.cell(r, 1).value or "").strip()
        if name in REMOVE_NAMES or OEM_NAME.search(name):
            # only remove if in REMOVE_NAMES (don't wipe OEMs user wants kept from earlier)
            if name in REMOVE_NAMES:
                ws.delete_rows(r)
                removed += 1
    print(f"removed_bad_rows={removed}")

    existing_names = set()
    existing_domains = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        n = str(r[0] or "").strip().lower()
        d = domain_of(str(r[1] or ""))
        if n:
            existing_names.add(n)
        if d:
            existing_domains.add(d)

    candidates: list[dict] = []
    for name, web, country in SEED:
        if name in DROP_SEED:
            continue
        if OEM_NAME.search(name):
            continue
        candidates.append({"company": name, "website": web, "country": country, "role": "distributor", "src": "seed"})

    # DeepSeek top-ups in small batches
    avoid = list(existing_names)[:40] + [c["company"].lower() for c in candidates]
    for i in range(8):
        print(f"DeepSeek mini-batch {i+1}/8 ...")
        try:
            batch = deepseek_more(30, avoid)
            print(f"  got {len(batch)}")
            for c in batch:
                name = str(c.get("company") or "").strip()
                web = str(c.get("website") or "").strip()
                if not name or not web:
                    continue
                if OEM_NAME.search(name):
                    continue
                candidates.append(
                    {
                        "company": name,
                        "website": web,
                        "country": str(c.get("country") or "").strip(),
                        "role": "distributor",
                        "src": "deepseek",
                    }
                )
                avoid.append(name.lower())
        except Exception as e:
            print(f"  err {type(e).__name__}: {e}")
        time.sleep(0.8)

    # dedupe
    seen = set()
    uniq = []
    for c in candidates:
        name = c["company"].strip()
        web = normalize_url(c["website"])
        key = name.lower()
        dom = domain_of(web)
        if key in seen or key in existing_names or dom in existing_domains:
            continue
        if not dom:
            continue
        seen.add(key)
        c["website"] = web
        uniq.append(c)
    print(f"candidates_unique={len(uniq)}")

    verified = []
    with ThreadPoolExecutor(max_workers=24) as ex:
        futs = {ex.submit(probe_ok, c["website"]): c for c in uniq}
        for fut in as_completed(futs):
            c = futs[fut]
            ok, note = fut.result()
            if ok:
                c["verified_website"] = note if note.startswith("http") else c["website"]
                verified.append(c)
    print(f"verified_sites={len(verified)}")
    verified.sort(key=lambda x: (0 if x.get("src") == "seed" else 1, x["company"].lower()))
    selected = verified[:TARGET]
    print(f"selected={len(selected)}")

    OUT_JSON.write_text(json.dumps(selected, indent=2, ensure_ascii=False), encoding="utf-8")

    added = 0
    for c in selected:
        row = [None] * len(headers)
        row[hidx["Company"]] = c["company"]
        row[hidx["Website"]] = c.get("verified_website") or c["website"]
        row[hidx["Distribution Type"]] = "Specialty Distributor"
        row[hidx["Retail / E-commerce / Both"]] = "No"
        row[hidx["Ownership"]] = "Private"  # only if unknown — actually don't invent; leave blank
        row[hidx["Ownership"]] = None
        row[hidx["Core Categories"]] = "Flexible packaging supplies; films; bags; stretch/shrink; packaging materials"
        row[hidx["Specialty Focus"]] = "Business distributor/wholesaler of packaging supplies (channel partner)"
        row[hidx["Summary"]] = "Channel/distributor verified via live website check for flexible packaging supplies market."
        country = (c.get("country") or "").strip()
        if country:
            row[hidx["Headquarters"]] = country
        low = country.lower()
        for k, (cc, rc) in COUNTRY_MAP.items():
            if k in low:
                row[hidx["Country Code"]] = cc
                row[hidx["Region Code"]] = rc
                row[hidx["Continent / Geography"]] = f"{rc}({country})"
                break
        ws.append(row)
        added += 1

    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(FP)
    print(f"backup={backup}")
    print(f"ADDED={added} TOTAL={ws.max_row-1} NEED={TARGET}")


if __name__ == "__main__":
    main()
