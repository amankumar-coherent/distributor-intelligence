"""Verify and fix websites for all Flexible Packaging FINAL rows."""
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
AUDIT = FP.parent / "website_fix_audit.json"

# Manual corrections from agent review / known official sites
MANUAL_FIX: dict[str, str] = {
    "Stretch Film UK": "",  # stretchfilm.co.uk parked — no verified official site
    "ULine India": "",  # tangyshop.com wrong — no verified Uline India site
    "Empire Packaging Dist": "",  # hugedomains parked
    "Imperial Dade": "https://www.imperialdade.com",
    "National Bag Company": "https://www.nationalbag.com",
    "National Bag": "https://www.nationalbag.com",
    "Polymer Packaging": "https://www.polymerpkg.com",
    "Polymer Packaging & Supply": "https://www.polymerpkg.com",
    "Polymer Packaging Inc.": "https://www.polymerpkg.com",
    "The Packaging Centre": "https://www.thepackagingcentre.co.uk",
    "Poly Bag Superstore": "https://www.plasticbag.com",
    "Packaging Solutions": "https://www.veritiv.com",
    "Polyair Inter Pack": "https://www.polyair.com",
    "Packlane": "https://www.packlane.com",
    "Midland Paper Packaging": "https://www.midlandpaper.com",
    "Midland Paper": "https://www.midlandpaper.com",
    "Allpack Packaging": "https://www.allpack.co.uk",
    "Allpack": "https://www.allpack.co.uk",
    "A-1 Packaging": "https://www.a1packaging.com",
    "Packaging Supplies UK": "https://www.pack-supplies.co.uk",
    "Packaging Online UK": "https://www.packagingonline.co.uk",
    "Packaging Suppliers UK": "https://www.wholesalepackaging.co.uk",
    "Industrial Packaging Supplies": "https://www.industrialpackaging.com",
    "Box and Packaging Supply": "https://www.boxandpackaging.com",
    "Bubble Packaging UK": "https://www.bubblepackaging.co.uk",
    "Plastic Bag Source": "https://www.plasticbagsource.com",
    "The Packaging Source": "https://www.thepackagingsource.com",
    "Packaging Supplies": "https://www.packagingsupplies.co.uk",
    "Canadian Packaging": "https://www.canadianpackaging.com",  # trade mag — may clear later
    "Veritiv Canada": "https://www.veritiv.com",
    "Veritiv Corporation": "https://www.veritiv.com",
    "Antalis": "https://www.antalis.com",
    "Antalis Packaging": "https://www.antalis.com",
    "Emballages Carrousel": "https://www.carrousel.ca",
    "Atlantic Packaging Products Ltd.": "https://www.atlantic.ca",
    "Macfarlane Group": "https://www.macfarlanegroup.com",
    "Orora": "https://www.ororagroup.com",
    "Quill.com Packaging": "https://www.quill.com",
    "Zoro Tools": "https://www.zoro.com",
    "IndustryBuying Packaging": "https://www.industrybuying.com",
    "Moglix Packaging": "https://www.moglix.com",
    "Office Depot Business Solutions": "https://www.odpbusiness.com",
    "Staples Business Advantage": "https://www.staplesadvantage.com",
    "RAJA Group Corporate": "https://www.raja-group.com",
    "Kalamazoo Spain": "https://www.kalamazoo.es",
    "Cole-Parmer Packaging Dist": "https://www.coleparmer.com",
    "Fisher Scientific Packaging Dist": "https://www.fishersci.com",
    "Thomas Scientific Packaging Dist": "https://www.thomassci.com",
    "VWR Packaging Lab Dist": "https://www.vwr.com",
    "Qorpak Packaging Dist": "https://www.qorpak.com",
    "Qorpak": "https://www.qorpak.com",
    "Laddawn Packaging Dist": "https://www.laddawn.com",
    "Jamestown Container Dist": "https://www.jamestowncontainer.com",
    "Atlantic Packaging Corp Dist": "https://www.atlanticpkg.com",
    "All American Containers Dist": "https://www.allamericancontainers.com",
    "Empire Packaging Dist": "",
    "JBM Packaging Dist": "https://www.jbmpackaging.com",
    "Ohio Packaging Corp Dist": "https://www.ohiopackaging.com",
    "Package Craft Dist": "https://www.packagecraft.com",
    "Packaging Unlimited Dist": "https://www.packagingunlimited.com",
    "Western Container Corp Dist": "https://www.westerncontainer.com",
    "Western Packaging Dist": "https://www.westernpackaging.com",
    "Premier Packaging UK Merchant": "https://www.premierpackaging.co.uk",
    "Seal Packaging UK Merchant": "https://www.sealpackaging.co.uk",
    "Transpak UK Merchant": "https://www.transpack.co.uk",
    "Packaging Supplies UK Merchant": "https://www.pack-supplies.co.uk",
    "Pro-Pack Verpackungen": "https://www.pro-pack.de",
    "Roberts Packaging": "https://www.robertspackaging.com",
    "Cal-Pak Packaging": "https://www.calpak.com",
    "Coastal Packaging Supply": "https://www.coastalpackaging.com",
    "Delaware Valley Packaging": "https://www.dvpackaging.com",
    "Package Design Supply Co": "https://www.pdsco.com",
    "Packaging Ireland": "https://www.packaging.ie",
    "Ahlsell": "https://www.ahlsell.com",
    "Manutan France": "https://www.manutan.fr",
    "Manutan UK": "https://www.manutan.co.uk",
    "Manutan Belgium": "https://www.manutan.be",
    "Manutan Germany": "https://www.manutan.de",
    "Manutan Italy": "https://www.manutan.it",
    "Manutan Netherlands": "https://www.manutan.nl",
    "Manutan Poland": "https://www.manutan.pl",
    "Manutan Spain": "https://www.manutan.es",
    "Uline Mexico": "https://www.uline.mx",
    "Uline Canada": "https://www.uline.ca",
    "Uline": "https://www.uline.com",
    "Bunzl Distribution": "https://www.bunzldistribution.com",
    "SupplyOne": "https://www.supplyone.com",
    "Shorr Packaging": "https://www.shorr.com",
    "Wrap-Tite": "https://www.wraptite.com",
    "Schermerhorn Bros. Co.": "https://www.schermerhornbrosco.com",
    "Container and Packaging Supply": "https://www.containerandpackaging.com",
    "A. R. Arena Products": "https://www.arenaproducts.com",
    "Bags & Bows": "https://www.bagsandbowsonline.com",
    "Bags and Bows": "https://www.bagsandbowsonline.com",
    "Brown Packaging": "https://www.brownpackaging.com",
    "TapeCase": "https://www.tapecase.com",
    "Packaging Plus": "https://www.packagingplus.com",
    "Braden Packaging & Supply": "https://www.bradensupply.com",
    "Cenpac": "https://www.cen-pac.co.uk",
    "Chase Packaging Ltd": "https://www.chase-packaging.co.uk",
    "EVO-PE Flexible Packaging Film Service": "https://www.evope.co.uk",
    "FlexPRO Europe": "https://www.flexproeurope.com",
    "United One Source": "https://www.unitedonesource.com",
    "Crawford Packaging": "https://www.crawfordpackaging.com",
    "Dayton Bag & Burlap": "https://www.daytonbag.com",
    "Gertex Solutions": "https://www.gertexsolutions.com",
    "HD Supply": "https://www.hdsupplysolutions.com",
    "Kendon Packaging": "https://www.kendonpackaging.co.uk",
    "Multi-Craft Packaging": "https://www.multicraft.ca",
    "Noissue": "https://www.noissue.co",
    "Novvia Group": "https://www.novviagroup.com",
    "Packhelp": "https://www.packhelp.com",
    "Preferred Plastics & Packaging": "https://www.prefplastics.com",
    "Richards Packaging": "https://www.richardspackaging.com",
    "Unipac Packaging Products": "https://www.myunipac.com",
    "Network Packaging": "https://www.networkpack.co.uk",
    "Unicorr Packaging": "https://www.unicorr.com",
    "Spicers Paper Canada": "https://www.spicers.ca",
    "Polythene UK": "https://www.polytheneuk.co.uk",
    "Royal Bag": "https://www.royalbag.com",
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
    "Emirates Packaging Industries": "https://www.emiratespack.com",
}

BAD_HOST_MARKERS = (
    "hugedomains.com",
    "domainmanage.com",
    "sedo.com",
    "godaddy.com",
    "afternic.com",
    "dan.com",
    "parkingcrew",
    "sedoparking",
    "tangyshop.com",
    "suspendedpage",
    "validate.perfdrive.com",
    "2glas.com",
    "buy-a-domain",
)


def normalize_url(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    # strip :443 and trailing junk
    u = re.sub(r":443(?=/|$)", "", u)
    u = u.rstrip("/")
    # collapse blog/acquisition deep links to homepage where known
    return u


def domain_of(url: str) -> str:
    u = normalize_url(url)
    if not u:
        return ""
    host = urlparse(u).netloc.lower().split("@")[-1]
    return host[4:] if host.startswith("www.") else host


def is_bad_final(url: str) -> bool:
    low = (url or "").lower()
    return any(m in low for m in BAD_HOST_MARKERS)


def probe(url: str) -> dict:
    u = normalize_url(url)
    if not u:
        return {"ok": False, "reason": "empty", "final": "", "code": None}
    host = domain_of(u)
    try:
        socket.getaddrinfo(host, 443, type=socket.SOCK_STREAM)
    except Exception:
        try:
            socket.getaddrinfo(host, 80, type=socket.SOCK_STREAM)
        except Exception:
            return {"ok": False, "reason": "dns", "final": u, "code": None}

    ctx = ssl.create_default_context()
    last_err = ""
    for attempt in (u, "https://" + host, "http://" + host, "https://www." + host):
        try:
            req = Request(attempt, headers={"User-Agent": "Mozilla/5.0 (compatible; VendorIntelBot/1.0)"}, method="GET")
            kw: dict = {"timeout": 14}
            if attempt.startswith("https"):
                kw["context"] = ctx
            with urlopen(req, **kw) as resp:
                code = getattr(resp, "status", 200)
                final = resp.geturl()
                if is_bad_final(final):
                    return {"ok": False, "reason": "parked_or_wrong", "final": final, "code": code}
                if code < 500:
                    # prefer clean canonical homepage-ish url
                    clean = normalize_url(attempt if code in (401, 403) else final)
                    if is_bad_final(clean):
                        return {"ok": False, "reason": "parked_or_wrong", "final": clean, "code": code}
                    return {"ok": True, "reason": "ok", "final": clean, "code": code}
        except Exception as e:
            last_err = str(e)
            if "403" in last_err or "401" in last_err:
                return {"ok": True, "reason": "ok_forbidden", "final": normalize_url(attempt), "code": 403}
            continue
    return {"ok": False, "reason": "fail", "final": u, "code": None, "error": last_err[:120]}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = FP.with_name(FP.stem + f"_backup_before_website_fix_{stamp}.xlsx")
    shutil.copy2(FP, backup)

    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    col = {str(h): i + 1 for i, h in enumerate(headers) if h}

    rows = []
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, col["Company"]).value or "").strip()
        site = str(ws.cell(r, col["Website"]).value or "").strip()
        rows.append({"row": r, "company": name, "website": site})

    # Apply manual fixes first
    applied_manual = []
    for item in rows:
        name = item["company"]
        if name in MANUAL_FIX:
            new = MANUAL_FIX[name]
            if new != item["website"]:
                applied_manual.append(
                    {"company": name, "old": item["website"], "new": new or "(cleared)"}
                )
                item["website"] = new
                item["manual"] = True

    # Probe all current websites
    results = {}
    with ThreadPoolExecutor(max_workers=28) as ex:
        futs = {}
        for item in rows:
            url = item["website"]
            if not url:
                results[item["row"]] = {"ok": False, "reason": "empty", "final": "", "code": None}
                continue
            futs[ex.submit(probe, url)] = item["row"]
        for fut in as_completed(futs):
            rid = futs[fut]
            try:
                results[rid] = fut.result()
            except Exception as e:
                results[rid] = {"ok": False, "reason": "exception", "final": "", "error": str(e)[:120]}

    fixed = []
    still_bad = []
    ok_count = 0
    for item in rows:
        r = item["row"]
        name = item["company"]
        old = str(ws.cell(r, col["Website"]).value or "").strip()
        candidate = item["website"]
        res = results.get(r, {"ok": False, "reason": "missing"})

        new_url = candidate
        status = "ok"
        if not candidate:
            status = "empty"
            new_url = ""
        elif res.get("ok"):
            final = normalize_url(res.get("final") or candidate)
            # Prefer https homepage root when final is a deep path for known companies
            if final.count("/") > 3 and name in MANUAL_FIX and MANUAL_FIX[name]:
                final = normalize_url(MANUAL_FIX[name])
            new_url = final
            ok_count += 1
            status = res.get("reason", "ok")
        else:
            # keep manual URL if set and looks sane, else mark bad
            if item.get("manual") and candidate and not is_bad_final(candidate):
                # re-probe already failed; keep manual for agent second pass
                new_url = normalize_url(candidate)
                status = f"kept_manual_unverified:{res.get('reason')}"
                still_bad.append({"company": name, "website": new_url, "reason": status})
            else:
                new_url = normalize_url(candidate) if candidate and not is_bad_final(candidate) else ""
                status = res.get("reason", "fail")
                still_bad.append(
                    {
                        "company": name,
                        "old": old,
                        "website": new_url or old,
                        "reason": status,
                        "probe_final": res.get("final"),
                    }
                )

        if normalize_url(old) != normalize_url(new_url):
            fixed.append({"company": name, "old": old, "new": new_url or "(cleared)", "status": status})
            ws.cell(r, col["Website"]).value = new_url or None

        # also normalize unchanged ok urls to https clean form
        elif new_url and normalize_url(old) != new_url and res.get("ok"):
            ws.cell(r, col["Website"]).value = new_url
            fixed.append({"company": name, "old": old, "new": new_url, "status": "normalized"})

    # Second pass: for remaining empty/bad, try company-name derived guesses already in MANUAL only
    # Write audit of still_bad for agent web search follow-up
    wb.save(FP)

    audit = {
        "backup": str(backup),
        "total_rows": len(rows),
        "ok_count": ok_count,
        "manual_applied": applied_manual,
        "changed": fixed,
        "changed_count": len(fixed),
        "still_bad": still_bad,
        "still_bad_count": len(still_bad),
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"TOTAL={len(rows)} OK={ok_count} CHANGED={len(fixed)} STILL_BAD={len(still_bad)}")
    print(f"backup={backup}")
    print(f"audit={AUDIT}")
    if still_bad:
        print("--- STILL_BAD ---")
        for x in still_bad[:80]:
            print(f"{x['company']} | {x.get('website') or x.get('old')} | {x['reason']}")


if __name__ == "__main__":
    main()
