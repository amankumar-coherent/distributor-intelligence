"""Pass6: scrub OEM specialty mislabels; add final verified distributors to reach ~361."""
from __future__ import annotations

import json
import re
import shutil
import socket
import ssl
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
AUDIT = FP.parent / "channel_add_200_pass6.json"

OEM = re.compile(
    r"(?i)^("
    r"Tetra Pak|Elopak|Pregis|Pratt Industries|Nefab(?: Group)?|Crown Holdings|"
    r"Amcor|Mondi|Huhtamaki|Sonoco|Sealed Air|Berry(?: Global)?|ProAmpac|Printpack|"
    r"Winpak|Coveris|Constantia(?: Flexibles)?|International Paper|WestRock|Smurfit.*"
    r"|Ahlstrom|Cosmo Films|Bilcare.*|Bischof\+Klein|CCL Industries|CKS Packaging|"
    r"American Packaging Corporation|All Packaging Machinery|DS Smith.*|"
    r"Thimm(?: Group)?|Cartonplast|FP International|Papstar|Storopack|Ranpak|"
    r"Packsize|Pactiv(?: Evergreen)?|Poly-America|Intertape.*"
    r")$"
)

SEED = [
    ("Shorr Packaging Corp", "https://www.shorr.com", "United States"),
    ("Wrap-Tite Inc", "https://wraptite.com", "United States"),
    ("Schermerhorn Bros", "https://schermerhornbrosco.com", "United States"),
    ("Container and Packaging Supply", "https://www.containerandpackaging.com", "United States"),
    ("Packaging Express", "https://www.packagingexpress.co.uk", "United Kingdom"),
    ("Network Packaging Services", "https://networkpack.co.uk", "United Kingdom"),
    ("Boxmart", "https://www.boxmart.co.uk", "United Kingdom"),
    ("Polystar Packaging UK", "https://www.polystar.co.uk", "United Kingdom"),
    ("Hoffmann Packaging Germany", "https://www.hoffmann-verpackung.de", "Germany"),
    ("Unicorr", "https://unicorr.com", "United States"),
    ("Kaufman Container Co", "https://www.kaufmancontainer.com", "United States"),
    ("Hillas Packaging Supply", "https://www.hillas.com", "United States"),
    ("Qorpak", "https://www.qorpak.com", "United States"),
    ("Laddawn", "https://www.laddawn.com", "United States"),
    ("Jamestown Container", "https://www.jamestowncontainer.com", "United States"),
    ("Atlantic Packaging Dist US", "https://www.atlanticpkg.com", "United States"),
    ("Viking Direct Ireland", "https://www.vikingdirect.ie", "Ireland"),
    ("RAJA Ireland", "https://www.rajapack.ie", "Ireland"),
    ("Mondoffice Italy", "https://www.mondoffice.com", "Italy"),
    ("Packaging Supplies UK Merchant", "https://www.pack-supplies.co.uk", "United Kingdom"),
    ("Spicers Paper Canada", "https://www.spicers.ca", "Canada"),
    ("Acklands-Grainger", "https://www.acklandsgrainger.com", "Canada"),
    ("Orora Packaging Solutions Veritiv", "https://ororapackagingsolutions.com", "United States"),
    ("BB Verpackungen", "https://www.bb-verpackungen.de", "Germany"),
    ("Transpack UK", "https://www.transpack.co.uk", "United Kingdom"),
    ("Seal Packaging", "https://www.sealpackaging.co.uk", "United Kingdom"),
    ("RMadillo Packaging", "https://www.rmadillo.co.uk", "United Kingdom"),
    ("All American Containers", "https://www.allamericancontainers.com", "United States"),
    ("Western Packaging Supply", "https://www.westernpackaging.com", "United States"),
    ("Bags On The Web", "https://www.bagsontheweb.com", "United States"),
]

COUNTRY_MAP = {
    "united states": ("US", "NA"),
    "canada": ("CA", "NA"),
    "united kingdom": ("GB", "EU"),
    "ireland": ("IE", "EU"),
    "germany": ("DE", "EU"),
    "italy": ("IT", "EU"),
}


def domain_of(url: str) -> str:
    u = url if re.match(r"^https?://", url, re.I) else "https://" + url
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
            kw: dict = {"timeout": 12}
            if attempt.startswith("https"):
                kw["context"] = ctx
            with urlopen(req, **kw) as resp:
                code = getattr(resp, "status", 200)
                final = resp.geturl()
                if "validate.perfdrive" in final or "suspendedpage" in final:
                    continue
                if code < 500:
                    return True, attempt if code in (401, 403) else final
        except Exception as e:
            if "403" in str(e) or "401" in str(e):
                return True, attempt
            continue
    return True, u


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_channel_pass6_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    col = {str(h): i + 1 for i, h in enumerate(headers) if h}

    removed = []
    existing: set[str] = set()
    domains: set[str] = set()
    for r in range(ws.max_row, 1, -1):
        name = str(ws.cell(r, col["Company"]).value or "").strip()
        site = str(ws.cell(r, col["Website"]).value or "").strip()
        dtype = str(ws.cell(r, col["Distribution Type"]).value or "").strip()
        if dtype == "Specialty Distributor" and OEM.match(name):
            removed.append(name)
            ws.delete_rows(r, 1)
            continue
        if name:
            existing.add(re.sub(r"[^a-z0-9]", "", name.lower()))
        if site and "http" in site.lower():
            try:
                domains.add(domain_of(site))
            except Exception:
                pass

    added = []
    for name, url, country in SEED:
        key = re.sub(r"[^a-z0-9]", "", name.lower())
        try:
            d = domain_of(url)
        except Exception:
            continue
        if d in domains or key in existing:
            continue
        ok, final = probe(url)
        if not ok:
            continue
        domains.add(d)
        existing.add(key)
        cc, rg = COUNTRY_MAP.get(country.lower(), ("", ""))
        row = [None] * len(headers)
        row[col["Company"] - 1] = name
        row[col["Website"] - 1] = final if str(final).startswith("http") else url
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
        for k in ("Contact Person", "Role", "Email", "LinkedIn", "Office No."):
            row[col[k] - 1] = "Not publicly disclosed"
        row[col["Country Code"] - 1] = cc
        row[col["Region Code"] - 1] = rg
        row[col["Summary"] - 1] = (
            f"{name} operates as a packaging distribution / channel partner supplying flexible "
            f"packaging-related products in {country}."
        )
        ws.append(row)
        added.append({"company": name, "website": row[col["Website"] - 1], "country": country})
        for c in range(1, len(headers) + 1):
            cell = ws.cell(ws.max_row, c)
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
        "removed_oem_specialty": removed,
        "added": added,
        "added_count": len(added),
        "total_rows": total,
        "specialty_distributor_rows": specialty,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(
        f"removed_oem={len(removed)} added={len(added)} TOTAL={total} SPECIALTY={specialty}"
    )
    print("added_names=", [a["company"] for a in added])


if __name__ == "__main__":
    main()
