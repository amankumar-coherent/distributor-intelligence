"""Pass2b: apply researched website fixes in parallel; DNS-live counts as OK."""
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

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "website_fix_audit_pass2.json"

FIXES: dict[str, str | None] = {
    "Emirates Packaging Industries": "https://empack.ae",
    "Eastern Bag & Paper": "https://www.imperialdade.com",
    "Imperial Bag & Paper Co.": "https://www.imperialdade.com",
    "Dayton Bag & Burlap": "https://www.daybag.com",
    "Flexipol Packaging": "https://www.flexipol.co.uk",
    "Stretch Film UK": None,
    "ULine India": None,
    "Empire Packaging Dist": None,
    "Flextrus": "https://www.flextrus.com",
    "Sanko Co. Ltd.": "https://www.sanko-p.com",
    "Mega P & L": "https://mega-pak.com",
    "Polyair Inter Pack": "https://www.itape.com",
    "The Packaging Centre": "https://www.thepackagingcentre.co.uk",
    "National Bag Company": "https://www.nationalbag.com",
    "National Bag": "https://www.nationalbag.com",
    "Imperial Dade": "https://www.imperialdade.com",
    "Packlane": "https://www.packlane.com",
    "Qorpak": "https://www.qorpak.com",
    "Qorpak Packaging Dist": "https://www.qorpak.com",
    "Orora": "https://www.ororagroup.com",
    "Macfarlane Group": "https://www.macfarlanegroup.com",
    "Veritiv Corporation": "https://www.veritiv.com",
    "Veritiv Canada": "https://www.veritiv.com",
    "Packaging Solutions": "https://www.veritiv.com",
    "Antalis": "https://www.antalis.com",
    "Antalis Packaging": "https://www.antalis.com",
    "Emballages Carrousel": "https://www.carrousel.ca",
    "Atlantic Packaging Products Ltd.": "https://www.atlantic.ca",
    "Midland Paper": "https://www.midlandpaper.com",
    "Midland Paper Packaging": "https://www.midlandpaper.com",
    "Allpack": "https://www.allpack.co.uk",
    "Allpack Packaging": "https://www.allpack.co.uk",
    "Polymer Packaging": "https://www.polymerpkg.com",
    "Polymer Packaging Inc.": "https://www.polymerpkg.com",
    "Polymer Packaging & Supply": "https://www.polymerpkg.com",
    "Poly Bag Superstore": "https://www.plasticbag.com",
    "Bags and Bows": "https://www.bagsandbowsonline.com",
    "Bags & Bows": "https://www.bagsandbowsonline.com",
    "Manutan France": "https://www.manutan.fr",
    "Manutan UK": "https://www.manutan.co.uk",
    "Manutan Belgium": "https://www.manutan.be",
    "Manutan Germany": "https://www.manutan.de",
    "Manutan Italy": "https://www.manutan.it",
    "Manutan Netherlands": "https://www.manutan.nl",
    "Manutan Poland": "https://www.manutan.pl",
    "Manutan Spain": "https://www.manutan.es",
    "Uline Mexico": "https://www.uline.mx",
    "Ahlsell": "https://www.ahlsell.com",
    "Kalamazoo Spain": "https://www.kalamazoo.es",
    "Network Packaging": "https://www.networkpack.co.uk",
    "Noissue": "https://www.noissue.co",
    "Packhelp": "https://www.packhelp.com",
    "A-1 Packaging": "https://a1packaging.com",
    "Packaging Supplies UK": "https://www.pack-supplies.co.uk",
    "Packaging Supplies UK Merchant": "https://www.pack-supplies.co.uk",
    "Industrial Packaging Supplies": "https://www.industrialpackaging.com",
    "Box and Packaging Supply": "https://www.boxandpackaging.com",
    "Bubble Packaging UK": "https://www.bubblepackaging.co.uk",
    "Plastic Bag Source": "https://www.plasticbagsource.com",
    "The Packaging Source": "https://www.thepackagingsource.com",
    "Packaging Supplies": "https://www.packagingsupplies.co.uk",
    "Pro-Pack Verpackungen": "https://www.pro-pack.de",
    "Roberts Packaging": "https://www.robertspackaging.com",
    "Cal-Pak Packaging": "https://www.calpak.com",
    "Cenpac": "https://www.cen-pac.co.uk",
    "Chase Packaging Ltd": "https://www.chase-packaging.co.uk",
    "EVO-PE Flexible Packaging Film Service": "https://www.evope.co.uk",
    "FlexPRO Europe": "https://www.flexproeurope.com",
    "TapeCase": "https://www.tapecase.com",
    "Packaging Plus": "https://www.packagingplus.com",
    "Brown Packaging": "https://www.brownpackaging.com",
    "A. R. Arena Products": "https://www.arenaproducts.com",
    "Bunzl Distribution": "https://www.bunzldistribution.com",
    "United One Source": "https://www.unitedonesource.com",
    "Braden Packaging & Supply": "https://www.bradensupply.com",
    "Royal Bag": "https://www.royalbag.com",
    "Polythene UK": "https://www.polytheneuk.co.uk",
    "Spicers Paper Canada": "https://www.spicers.ca",
    "Unipac Packaging Products": "https://www.myunipac.com",
    "Multi-Craft Packaging": "https://www.multicraft.ca",
    "Kendon Packaging": "https://www.kendonpackaging.co.uk",
    "Crawford Packaging": "https://www.crawfordpackaging.com",
    "Gertex Solutions": "https://www.gertexsolutions.com",
    "HD Supply": "https://www.hdsupplysolutions.com",
    "Preferred Plastics & Packaging": "https://www.prefplastics.com",
    "Richards Packaging": "https://www.richardspackaging.com",
    "Novvia Group": "https://www.novviagroup.com",
    "Shorr Packaging": "https://www.shorr.com",
    "Wrap-Tite": "https://www.wraptite.com",
    "Schermerhorn Bros. Co.": "https://www.schermerhornbrosco.com",
    "Container and Packaging Supply": "https://www.containerandpackaging.com",
    "Unicorr Packaging": "https://www.unicorr.com",
    "SupplyOne": "https://www.supplyone.com",
    "RAJA Group Corporate": "https://www.raja-group.com",
    "Cole-Parmer Packaging Dist": "https://www.coleparmer.com",
    "Thomas Scientific Packaging Dist": "https://www.thomassci.com",
    "VWR Packaging Lab Dist": "https://www.vwr.com",
    "Fisher Scientific Packaging Dist": "https://www.fishersci.com",
    "Action Packaging Systems": "https://www.actionpkg.com",
    "ePac Flexible Packaging": "https://www.epacflexibles.com",
    "Glenroy": "https://www.glenroy.com",
    "American Packaging Corporation": "https://www.americanpackaging.com",
    "Ahlstrom": "https://www.ahlstrom.com",
    "Bischof+Klein": "https://www.bischof-klein.com",
    "CCL Industries": "https://www.cclind.com",
    "Constantia Flexibles": "https://www.cflex.com",
    "Cosmo Films": "https://www.cosmofilms.com",
    "Danaflex": "https://www.danaflex.com",
    "Flair Flexible Packaging": "https://www.flairpackaging.com",
    "Gualapack": "https://www.gualapack.com",
    "Jindal Poly Films": "https://www.jindalpoly.com",
    "Korozo": "https://www.korozogroup.com",
    "Novolex": "https://www.novolex.com",
    "Perlen Packaging": "https://www.perlenpackaging.com",
    "RKW Group": "https://www.rkw-group.com",
    "SIG": "https://www.sig.biz",
    "Smurfit Westrock": "https://www.smurfitwestrock.com",
    "Suedpack": "https://www.suedpack.com",
    "Walki": "https://www.walki.com",
    "Wipak": "https://www.wipak.com",
    "Arabian Packaging Co.": "https://www.arabianpackaging.com",
    "Delta Packaging Systems": "https://www.deltapack.com",
    "Blackwoods AU": "https://www.blackwoods.com.au",
    "Officeworks AU": "https://www.officeworks.com.au",
    "Signet AU": "https://www.signet.com.au",
    "Office Depot": "https://www.officedepot.com",
    "Staples": "https://www.staples.com",
    "US Plastic Corp": "https://www.usplastic.com",
    "Freund Container & Supply": "https://www.freundcontainer.com",
    "Impak Canada": "https://www.impak.ca",
    "JPG France": "https://www.jpg.fr",
    "Laddawn Packaging Dist": "https://www.laddawn.com",
    "SRF": "https://www.srf.com",
    "Saudi Pack": "https://www.saudipack.com",
    "Polytex Packaging": "https://www.polytex.com.au",
    "Packaging House": "https://www.packaginghouse.co.nz",
    "Bulk Packaging": "https://www.bulkpackaging.com",
    "Global Packaging": "https://www.globalpackaging.ca",
    "H&H Packaging": "https://www.hhpackaging.com",
    "Jordan Packaging Co.": "https://www.jordanpackaging.com",
    "Karan Packaging": "https://www.karanpackaging.com",
    "Packaging Products Corporation": "https://www.packagingproducts.com",
    "Packaging Resources": "https://www.packagingresources.com",
    "Poly Pak America": "https://www.polypakamerica.com",
    "Pouch Partners": "https://www.pouchpartners.com",
    "Unique Packaging": "https://www.uniquepackaging.in",
    "Coastal Packaging Supply": "https://www.coastalpackaging.com",
    "Delaware Valley Packaging": "https://www.dvpackaging.com",
    "Package Design Supply Co": "https://www.pdsco.com",
    "Packaging Direct UK": "https://www.packagingdirect.co.uk",
    "Packaging Unlimited Dist": "https://www.packagingunlimited.com",
    "Premier Packaging UK Merchant": "https://www.premierpackaging.co.uk",
    "Seal Packaging UK Merchant": "https://www.sealpackaging.co.uk",
    "Western Container Corp Dist": "https://www.westerncontainer.com",
    "All American Containers Dist": "https://www.allamericancontainers.com",
    "Packaging Online UK": "https://www.challengepackaging.co.uk",
}

BAD_MARKERS = (
    "hugedomains",
    "domainmanage",
    "sedo.com",
    "tangyshop",
    "suspendedpage",
    "buy-a-domain",
    "validate.perfdrive",
    "2glas.com",
    "godaddy.com/forsale",
)


def normalize(url: str | None) -> str:
    if not url:
        return ""
    u = str(url).strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    u = re.sub(r":443(?=/|$)", "", u)
    return u.rstrip("/")


def domain_of(url: str) -> str:
    u = normalize(url)
    if not u:
        return ""
    host = urlparse(u).netloc.lower().split("@")[-1]
    return host[4:] if host.startswith("www.") else host


def dns_ok(url: str) -> bool:
    host = domain_of(url)
    if not host:
        return False
    for port in (443, 80):
        try:
            socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
            return True
        except Exception:
            continue
    return False


def probe(url: str) -> dict:
    u = normalize(url)
    if not u:
        return {"ok": False, "reason": "empty", "final": ""}
    if any(m in u.lower() for m in BAD_MARKERS):
        return {"ok": False, "reason": "bad_marker", "final": u}
    if not dns_ok(u):
        return {"ok": False, "reason": "dns", "final": u}
    host = domain_of(u)
    ctx = ssl.create_default_context()
    for attempt in (u, "https://" + host, "http://" + host):
        try:
            req = Request(
                attempt,
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36"
                },
            )
            kw: dict = {"timeout": 10}
            if attempt.startswith("https"):
                kw["context"] = ctx
            with urlopen(req, **kw) as resp:
                final = normalize(resp.geturl())
                if any(m in final.lower() for m in BAD_MARKERS):
                    return {"ok": False, "reason": "parked", "final": final}
                return {"ok": True, "reason": "http", "final": normalize(attempt) if getattr(resp, "status", 200) in (401, 403) else final}
        except Exception as e:
            msg = str(e)
            if any(x in msg for x in ("403", "401", "406", "429")):
                return {"ok": True, "reason": "http_blocked", "final": normalize(attempt)}
            continue
    # DNS live but HTTP hardened — keep URL
    return {"ok": True, "reason": "dns_only", "final": u}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_website_fix2b_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    col = {str(h): i + 1 for i, h in enumerate(headers) if h}

    # Flextrus: only keep if DNS works
    if not dns_ok("https://www.flextrus.com"):
        FIXES["Flextrus"] = None

    planned: dict[int, str | None] = {}
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, col["Company"]).value or "").strip()
        old = str(ws.cell(r, col["Website"]).value or "").strip()
        if name in FIXES:
            planned[r] = FIXES[name]
        else:
            planned[r] = old if old else None

    # Probe all planned URLs in parallel
    probe_targets = {}
    for r, url in planned.items():
        n = normalize(url) if url else ""
        probe_targets[r] = n

    results: dict[int, dict] = {}
    with ThreadPoolExecutor(max_workers=32) as ex:
        futs = {ex.submit(probe, url): r for r, url in probe_targets.items() if url}
        for r, url in probe_targets.items():
            if not url:
                results[r] = {"ok": False, "reason": "empty", "final": ""}
        for fut in as_completed(futs):
            r = futs[fut]
            try:
                results[r] = fut.result()
            except Exception as e:
                results[r] = {"ok": False, "reason": f"exc:{e}", "final": probe_targets[r]}

    changed = []
    empty = []
    bad = []
    ok_n = 0
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, col["Company"]).value or "").strip()
        old = str(ws.cell(r, col["Website"]).value or "").strip()
        res = results.get(r, {"ok": False, "reason": "missing", "final": ""})
        if res.get("ok"):
            # Prefer clean FIXES homepage over deep redirects
            prefer = normalize(FIXES[name]) if name in FIXES and FIXES[name] else ""
            final = prefer if prefer else normalize(res.get("final") or "")
            if not final:
                final = normalize(probe_targets.get(r, ""))
            if normalize(old) != final:
                changed.append({"company": name, "old": old, "new": final, "reason": res.get("reason")})
            ws.cell(r, col["Website"]).value = final
            ok_n += 1
        else:
            # clear bad/empty
            new = None
            if name in FIXES and FIXES[name] is None:
                new = None
            elif probe_targets.get(r) and dns_ok(probe_targets[r]):
                new = normalize(probe_targets[r])
                ok_n += 1
            else:
                new = None
            if normalize(old) != normalize(new or ""):
                changed.append(
                    {
                        "company": name,
                        "old": old,
                        "new": new or "(cleared)",
                        "reason": res.get("reason"),
                    }
                )
            ws.cell(r, col["Website"]).value = new
            if not new:
                empty.append(name)
                bad.append({"company": name, "old": old, "reason": res.get("reason")})

    wb.save(FP)
    audit = {
        "backup": str(backup),
        "total": ws.max_row - 1,
        "ok_or_dns": ok_n,
        "empty_count": len(empty),
        "empty": empty,
        "changed_count": len(changed),
        "changed": changed,
        "bad": bad,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"TOTAL={ws.max_row-1} OK={ok_n} EMPTY={len(empty)} CHANGED={len(changed)}")
    print("EMPTY:", ", ".join(empty) if empty else "(none)")
    print(f"audit={AUDIT}")


if __name__ == "__main__":
    main()
