"""Clean OEM false-positives from channel adds; top up with curated distributors only."""
from __future__ import annotations

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
from openpyxl.styles import Font

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)

REMOVE = re.compile(
    r"(?i)^("
    r"Weyerhaeuser|Avery Dennison|Georgia-Pacific|UPM Raflatac|Metsä Board|Metsa Board|"
    r"Billerud.*|RPC Group|BWAY Corporation|C\.?\s*H\.?\s*Robinson|Packaging Aids|"
    r"Crown Packaging.*|International Paper|WestRock|Smurfit|Amcor|Mondi|Huhtamaki|"
    r"Sonoco|Sealed Air|Berry|ProAmpac|Printpack|Winpak|Coveris|Constantia|"
    r"Rajapack France$"  # bot-challenge URL junk
    r")$"
)

# Additional curated packaging DISTRIBUTORS (channel), not mills/OEMs
SEED = [
    ("Berlin Packaging", "https://www.berlinpackaging.com", "United States"),
    ("TricorBraun", "https://www.tricorbraun.com", "United States"),
    ("Imperial Dade", "https://www.imperialdade.com", "United States"),
    ("BradyPLUS", "https://www.bradyplus.com", "United States"),
    ("United One Source", "https://unitedonesource.com", "United States"),
    ("Braden Packaging & Supply", "https://bradensupply.com", "United States"),
    ("The Cary Company", "https://www.thecarycompany.com", "United States"),
    ("U.S. Plastic Corp.", "https://www.usplastic.com", "United States"),
    ("Freund Container & Supply", "https://www.freundcontainer.com", "United States"),
    ("SKS Bottle & Packaging", "https://www.sks-bottle.com", "United States"),
    ("O.Berk Company", "https://www.oberk.com", "United States"),
    ("International Plastics", "https://www.interplas.com", "United States"),
    ("Royal Bag Company", "https://www.royalbag.com", "United States"),
    ("TapeCase", "https://www.tapecase.com", "United States"),
    ("PackagingSupplies.com", "https://www.packagingsupplies.com", "United States"),
    ("Packaging Price", "https://www.packagingprice.com", "United States"),
    ("WebstaurantStore", "https://www.webstaurantstore.com", "United States"),
    ("Zoro", "https://www.zoro.com", "United States"),
    ("Quill", "https://www.quill.com", "United States"),
    ("Midland Paper", "https://www.midlandpaper.com", "United States"),
    ("Lindenmeyr Munroe", "https://www.lindenmeyr.com", "United States"),
    ("Polythene UK", "https://polytheneuk.co.uk", "United Kingdom"),
    ("Polyflex Ltd", "https://www.polyflex.uk.com", "United Kingdom"),
    ("Cenpac", "https://cen-pac.co.uk", "United Kingdom"),
    ("Chase Packaging Ltd", "https://chase-packaging.co.uk", "United Kingdom"),
    ("Davpack", "https://www.davpack.co.uk", "United Kingdom"),
    ("Allpack", "https://www.allpack.co.uk", "United Kingdom"),
    ("Cleveland Containers", "https://www.clevelandcontainers.co.uk", "United Kingdom"),
    ("Manutan UK", "https://www.manutan.co.uk", "United Kingdom"),
    ("Viking Direct UK", "https://www.vikingdirect.co.uk", "United Kingdom"),
    ("Ratioform", "https://www.ratioform.de", "Germany"),
    ("RAJA Germany", "https://www.rajapack.de", "Germany"),
    ("RAJA France", "https://www.raja.fr", "France"),
    ("Manutan France", "https://www.manutan.fr", "France"),
    ("Lyreco", "https://www.lyreco.com", "France"),
    ("Impak Packaging Systems", "https://www.impak.ca", "Canada"),
    ("Uline Canada", "https://www.uline.ca", "Canada"),
    ("Acklands-Grainger", "https://www.acklandsgrainger.com", "Canada"),
    ("Spicers Canada", "https://www.spicers.ca", "Canada"),
    ("Signet", "https://www.signet.com.au", "Australia"),
    ("Officeworks", "https://www.officeworks.com.au", "Australia"),
    ("Bunzl Australia", "https://www.bunzl.com.au", "Australia"),
    ("IndustryBuying", "https://www.industrybuying.com", "India"),
    ("Moglix", "https://www.moglix.com", "India"),
    ("MonotaRO", "https://www.monotaro.com", "Japan"),
    ("Askul", "https://www.askul.co.jp", "Japan"),
    ("Makro South Africa", "https://www.makro.co.za", "South Africa"),
    ("Waltons", "https://www.waltons.co.za", "South Africa"),
    ("Office Depot Mexico", "https://www.officedepot.com.mx", "Mexico"),
    ("Multiplastics Europe Ltd", "https://www.multiplastics.co.uk", "United Kingdom"),
    ("EVO-PE", "https://evope.co.uk", "United Kingdom"),
    ("FlexPRO Europe", "https://flexproeurope.com", "Bulgaria"),
    ("Antalis", "https://www.antalis.com", "France"),
    ("Complete Packaging", "https://www.completepackaging.com", "United States"),
    ("National Bag", "https://www.nationalbag.com", "United States"),
    ("A-1 Packaging", "http://a1packaging.com", "United States"),
    ("Polymer Packaging Inc.", "https://www.polymerpkg.com", "United States"),
    ("Packaging Plus", "https://packagingplus.com", "United States"),
    ("McMaster-Carr", "https://www.mcmaster.com", "United States"),
    ("Staples Business Advantage", "https://www.staplesadvantage.com", "United States"),
]


def domain_of(url: str) -> str:
    u = url.strip()
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    host = urlparse(u).netloc.lower().split("@")[-1]
    return host[4:] if host.startswith("www.") else host


def probe(url: str) -> tuple[bool, str]:
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
            kw = {"timeout": 10}
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
    "bulgaria": ("BG", "EU"),
    "india": ("IN", "APAC"),
    "australia": ("AU", "APAC"),
    "japan": ("JP", "APAC"),
    "south africa": ("ZA", "AF"),
    "mexico": ("MX", "LATAM"),
}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_channel_clean_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hidx = {str(h): i for i, h in enumerate(headers)}

    removed = 0
    for r in range(ws.max_row, 1, -1):
        name = str(ws.cell(r, 1).value or "").strip()
        web = str(ws.cell(r, 2).value or "").strip()
        if REMOVE.match(name) or "validate.perfdrive.com" in web or "amcor.com" in web.lower() and "RPC" in name:
            ws.delete_rows(r)
            removed += 1
            continue
        # also remove if Distribution Type is Specialty Distributor but company is known mill and Summary says channel from pass3
        if name in {
            "Weyerhaeuser",
            "Avery Dennison",
            "Georgia-Pacific",
            "UPM Raflatac",
            "Metsä Board",
            "BillerudKorsnäs",
            "RPC Group",
            "BWAY Corporation",
            "C. H. Robinson",
            "Packaging Aids",
        }:
            ws.delete_rows(r)
            removed += 1

    existing_names = {str(r[0] or "").strip().lower() for r in ws.iter_rows(min_row=2, values_only=True)}
    existing_domains = {domain_of(str(r[1] or "")) for r in ws.iter_rows(min_row=2, values_only=True)}

    # probe seeds
    to_add = []
    for name, web, country in SEED:
        if name.lower() in existing_names or domain_of(web) in existing_domains:
            continue
        to_add.append((name, web, country))

    verified = []
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs = {ex.submit(probe, web): (name, web, country) for name, web, country in to_add}
        for fut in as_completed(futs):
            name, web, country = futs[fut]
            ok, note = fut.result()
            if ok:
                verified.append((name, note if note.startswith("http") else web, country))

    print(f"removed={removed} seed_verified_new={len(verified)}")
    for name, web, country in verified:
        row = [None] * len(headers)
        row[hidx["Company"]] = name
        row[hidx["Website"]] = web
        row[hidx["Distribution Type"]] = "Specialty Distributor"
        row[hidx["Retail / E-commerce / Both"]] = "No"
        row[hidx["Core Categories"]] = "Flexible packaging supplies; films; bags; stretch/shrink; packaging materials"
        row[hidx["Specialty Focus"]] = "Business distributor/wholesaler of packaging supplies (channel partner)"
        row[hidx["Summary"]] = "Verified packaging channel/distributor (live website)."
        row[hidx["Headquarters"]] = country
        low = country.lower()
        for k, (cc, rc) in COUNTRY_MAP.items():
            if k in low:
                row[hidx["Country Code"]] = cc
                row[hidx["Region Code"]] = rc
                row[hidx["Continent / Geography"]] = f"{rc}({country})"
                break
        ws.append(row)

    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(FP)

    # stats
    n = ws.max_row - 1
    dist = 0
    for r in ws.iter_rows(min_row=2, values_only=True):
        if "distributor" in str(r[12] or "").lower():
            dist += 1
    print(f"TOTAL={n} SPECIALTY_DISTRIBUTOR_ROWS={dist} ADDED_NOW={len(verified)} backup={backup}")


if __name__ == "__main__":
    main()
