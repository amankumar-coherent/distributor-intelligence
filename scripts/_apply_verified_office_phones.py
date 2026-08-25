# -*- coding: utf-8 -*-
"""Overlay verified HQ/switchboard phones (SEC / official legal pages) onto empty Office No."""
from __future__ import annotations

from pathlib import Path

from _harvest_office_phones import PATHS, apply_excel, has_val, load_cache, save_cache, abs_url
from openpyxl import load_workbook

# Verified from SEC 10-K / official legal/contact pages. Never guessed.
BY_COMPANY = {
    "Amazon": "+1 206-266-1000",
    "Walmart": "+1 479-273-4000",
    "Target": "+1 612-304-6073",
    "Best Buy": "+1 612-291-1000",
    "CDW": "+1 847-465-6000",
    "TD SYNNEX": "+1 510-668-3400",
    "Ingram Micro": "+1 714-566-1000",
    "Arrow Electronics": "+1 303-824-4000",
    "Arrow Electronics (Asia Pacific)": "+1 303-824-4000",
    "Avnet": "+1 480-643-2000",
    "Avnet (Asia Pacific)": "+1 480-643-2000",
    "Anixter (now WESCO)": "+1 412-454-2200",
    "ALSO Holding": "+41 41 266 11 11",
    "MediaMarktSaturn": "+49 211 5408 7000",
    "Sysco": "+1 281-584-1390",
    "Fnac Darty": "+33 1 55 21 57 00",
    "Dicker Data": "1800 688 586",
    "ScanSource": "+1 800-944-2432",
    "SHI International": "+1 888-764-8888",
    "Insight Enterprises": "+1 800-467-4448",
    "Softcat": "+44 1628 403403",
    "Azelis": "+32 3 613 01 20",
    "Computacenter": "+44 1707 631000",
    "ELKO Group": "+371 6709 3230",
    "ASBIS": "+357 25 857 000",
    "Asbis": "+357 25 857 000",
    "Bechtle": "+49 7132 981-20",
    "Border States Electric": "+1 701-293-5834",
    "Harvey Norman": "+61 2 9201 6111",
}

# Domain overlay when company name variants differ across files
BY_DOMAIN = {
    "amazon.com": "+1 206-266-1000",
    "walmart.com": "+1 479-273-4000",
    "target.com": "+1 612-304-6073",
    "bestbuy.com": "+1 612-291-1000",
    "cdw.com": "+1 847-465-6000",
    "tdsynnex.com": "+1 510-668-3400",
    "ingrammicro.com": "+1 714-566-1000",
    "arrow.com": "+1 303-824-4000",
    "avnet.com": "+1 480-643-2000",
    "wesco.com": "+1 412-454-2200",
    "anixter.com": "+1 412-454-2200",
    "also.com": "+41 41 266 11 11",
    "ceconomy.de": "+49 211 5408 7000",
    "ceconomy-mms.com": "+49 211 5408 7000",
    "mediamarktsaturn.com": "+49 211 5408 7000",
    "rexel.com": "+33 1 42 85 85 00",
    "sonepar.com": "+33 1 58 44 13 13",
    "grainger.com": "+1 847-535-1000",
    "mckesson.com": "+1 972-446-4800",
    "sysco.com": "+1 281-584-1390",
    "fnacdarty.com": "+33 1 55 21 57 00",
    "dickerdata.com.au": "1800 688 586",
    "scansource.com": "+1 800-944-2432",
    "shi.com": "+1 888-764-8888",
    "insight.com": "+1 800-467-4448",
    "softcat.com": "+44 1628 403403",
    "azelis.com": "+32 3 613 01 20",
    "brenntag.com": "+49 201 6496-0",
    "imcdgroup.com": "+31 10 290 8684",
    "computacenter.com": "+44 1707 631000",
    "elkogroup.com": "+371 6709 3230",
    "asbis.com": "+357 25 857 000",
    "bechtle.com": "+49 7132 981-20",
    "borderstates.com": "+1 701-293-5834",
    "harveynorman.com.au": "+61 2 9201 6111",
    "harveynormanholdings.com.au": "+61 2 9201 6111",
    "cargill.com": "+1 800-227-4455",
}


def domain_of(website: str) -> str:
    u = abs_url(website).lower()
    if not u:
        return ""
    host = u.split("://", 1)[-1].split("/", 1)[0]
    if host.startswith("www."):
        host = host[4:]
    return host


def main() -> None:
    cache = load_cache()
    for host, phone in BY_DOMAIN.items():
        # seed cache keys that contain this host
        for k in list(cache.keys()):
            if host in k and not cache.get(k):
                cache[k] = phone
        # also store canonical
        cache[f"https://{host}"] = phone
        cache[f"https://www.{host}"] = phone
    save_cache(cache)

    for p in PATHS:
        if not p.exists():
            continue
        wb = load_workbook(p)
        ws = wb["Companies"] if "Companies" in wb.sheetnames else wb.active
        headers = None
        hr = 1
        for r in range(1, 6):
            vals = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=r, max_row=r))]
            if "Company" in vals:
                headers = vals
                hr = r
                break
        co_i = headers.index("Company") + 1
        web_i = headers.index("Website") + 1 if "Website" in headers else None
        off_i = headers.index("Office No.") + 1
        n = 0
        for r in range(hr + 1, ws.max_row + 1):
            co = str(ws.cell(r, co_i).value or "").strip()
            if not co:
                continue
            cur = str(ws.cell(r, off_i).value or "").strip()
            if has_val(cur):
                continue
            phone = BY_COMPANY.get(co) or ""
            if not phone and web_i:
                host = domain_of(str(ws.cell(r, web_i).value or ""))
                if host:
                    phone = BY_DOMAIN.get(host) or ""
                    if not phone:
                        for d, ph in BY_DOMAIN.items():
                            if host == d or host.endswith("." + d):
                                phone = ph
                                break
            if phone:
                ws.cell(r, off_i).value = phone
                n += 1
        tmp = p.with_suffix(".xlsx.tmp")
        wb.save(tmp)
        wb.close()
        tmp.replace(p)
        print(f"SEED {p.name} filled={n}")


if __name__ == "__main__":
    main()
