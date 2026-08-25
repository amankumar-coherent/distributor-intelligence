#!/usr/bin/env python3
"""Re-verify avocado OFFICE SLIM: fix junk websites, re-harvest About/Contact/footer
for Office No. + Office Email; verify existing values; never invent.
"""
from __future__ import annotations

import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse

import httpx
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global")
XLSX = ROOT / "global_avocado_oil_market_OFFICE_SLIM.xlsx"
CACHE = ROOT / "_avocado_office_reverify_cache.json"
AUDIT = ROOT / "_avocado_office_reverify_audit.json"
OVERRIDES_PATH = ROOT / "_avocado_web_verified_overrides.json"

NPD = "Not publicly disclosed"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# Verified website fixes (agent web search)
WEBSITE_FIXES = {
    "Atlantic Chemicals Trading": "https://www.act.de",
    "Bakker Barendrecht": "https://www.bakkerbarendrecht.nl",
    "Black Swan Distribution": "https://bsdc.co",
    "Connect The World S.A. de C.V.": "https://www.ctworldmx.com",
    "Eurobanan": "https://eurobanan.com",
    "KIC Chemicals": "https://kicgroup.com",
    "Liberty Natural Products": "https://www.libertynatural.com",
    # no reliable standalone site found — leave NPD
    "DIVERSE WORLDWIDE TRADING": "",
    "Levarht": "",  # bankrupt 2022; no active official site
}

# Verified office contacts from agent web search / official pages
CONTACT_OVERRIDES = {
    "Atlantic Chemicals Trading": {
        "office_no": "+49 40 33 44 19-0",
        "office_email": "",
        "source": "https://www.act.de/contact",
    },
    "Bakker Barendrecht": {
        "office_no": "+31 088 69 59 911",
        "office_email": "info@bakkerbarendrecht.nl",
        "source": "https://www.bakkerbarendrecht.nl/Bedrijfsgegevens/Contact",
    },
    "Black Swan Distribution": {
        "office_no": "(727) 390-9231",
        "office_email": "orders@blackswandistributioncorp.com",
        "source": "https://bsdc.co/contact-us/",
    },
    "Eurobanan": {
        "office_no": "+34 902 485 000",
        "office_email": "madrid@eurobanan.com",
        "source": "Eurobanan corporate PDF / eurobanan.com",
    },
    "KIC Chemicals": {
        "office_no": "(845) 883-5306",
        "office_email": "erick@kicgroup.com",
        "source": "https://www.foodmaster.com/directories/72-food-ingredients/listing/11579-kic-chemicals-inc",
    },
    "Liberty Natural Products": {
        "office_no": "800-289-8427",
        "office_email": "sales@libertynatural.com",
        "source": "https://www.libertynatural.com/Wholesale%20Policies/index.html",
    },
}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
TEL_HREF_RE = re.compile(r"tel:([+\d][\d\-().\s]{6,})", re.I)
PHONE_HINT_RE = re.compile(
    r"(?:tel:|phone|tel\.|call|fax|t\s*[\(:]|office|contact|hotline|customer\s*service|telefon|telephone)"
    r"[^0-9+]{0,30}(\+?[\d][\d\s().\-/]{7,22}\d)",
    re.I,
)
MAILTO_RE = re.compile(r"mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", re.I)

BAD_EMAIL_SUB = (
    "example.com",
    "exemple.com",
    "w3schools",
    "mysite.com",
    "opencart",
    "gmail.com",
    "sentry",
    "wixpress",
    "schema.org",
    "cloudflare",
    "financebuzz",
    "webp",
    "png",
    "jpg",
    "liberty.edu",
    "jane.doe",
    "asxvm",
    "domain.com",
    "email.com",
    "yourcompany",
    "noreply",
    "no-reply",
)

OFFICE_LOCAL = re.compile(
    r"^(info|contact|sales|office|hello|enquiry|enquiries|support|customer|customerservice|"
    r"customerservices|customersupport|orders|export|mail|admin|team|trade|help|press|media|"
    r"technical|inquiry|kontakt|infouk|communications|ahcsc|bestway\.group|info\.desk|"
    r"customercare|service|madrid|promo)([._-]|$)",
    re.I,
)

CONTACT_PATHS = [
    "",
    "/contact",
    "/contact-us",
    "/contactus",
    "/contacts",
    "/about",
    "/about-us",
    "/aboutus",
    "/en/contact",
    "/en/contact-us",
    "/company/contact",
    "/pages/contact",
    "/impressum",
    "/imprint",
    "/legal/imprint",
    "/Bedrijfsgegevens/Contact",
]


def emptyish(v: object) -> bool:
    s = str(v or "").strip().lower()
    return s in {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}


def normalize_website(url: str) -> str:
    u = (url or "").strip()
    if emptyish(u):
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    return u.rstrip("/")


def clean_phone(raw: str) -> str:
    s = re.sub(r"\s+", " ", str(raw or "").strip()).replace("tel:", "").strip()
    if len(re.sub(r"\D", "", s)) < 8:
        return ""
    return s[:40]


def clean_email(raw: str) -> str:
    e = str(raw or "").strip().lower().lstrip("%20").rstrip(".")
    m = re.search(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", e)
    if m:
        e = m.group(1)
    if not e or "@" not in e:
        return ""
    if any(b in e for b in BAD_EMAIL_SUB):
        return ""
    if e.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js")):
        return ""
    return e


def is_office_email(email: str) -> bool:
    e = clean_email(email)
    if not e:
        return False
    local = e.split("@")[0]
    if OFFICE_LOCAL.search(local):
        return True
    if re.search(r"(info|contact|sales|office|support|enquir|customer|hello|export|orders|service)", local, re.I):
        return True
    return False


def extract_from_html(html: str, base_url: str) -> tuple[str, str]:
    phones: list[str] = []
    emails: list[str] = []
    for m in TEL_HREF_RE.finditer(html):
        p = clean_phone(m.group(1))
        if p:
            phones.append(p)
    for m in MAILTO_RE.finditer(html):
        e = clean_email(m.group(1))
        if e and is_office_email(e):
            emails.append(e)
    for m in PHONE_HINT_RE.finditer(html):
        p = clean_phone(m.group(1))
        if p:
            phones.append(p)
    for m in EMAIL_RE.finditer(html):
        e = clean_email(m.group(0))
        if e and is_office_email(e):
            emails.append(e)

    def uniq(xs: list[str]) -> list[str]:
        seen = set()
        out = []
        for x in xs:
            k = re.sub(r"\D", "", x) if "@" not in x else x
            if k in seen:
                continue
            seen.add(k)
            out.append(x)
        return out

    phone = uniq(phones)[0] if phones else ""
    ems = uniq(emails)
    host = urlparse(base_url).netloc.lower().removeprefix("www.").split(":")[0]
    domain_match = [e for e in ems if host and host in e.split("@")[-1]]
    prefer = [
        e
        for e in (domain_match or ems)
        if e.split("@")[0]
        in {
            "info",
            "contact",
            "sales",
            "office",
            "hello",
            "enquiries",
            "enquiry",
            "support",
            "customerservice",
            "orders",
            "export",
        }
    ]
    email = (prefer or domain_match or ems)[0] if (prefer or domain_match or ems) else ""
    return phone, email


def fetch_contacts(website: str) -> dict:
    url = normalize_website(website)
    if not url:
        return {"phone": "", "email": "", "ok": False, "error": "no website", "url": ""}
    phone = email = ""
    last_err = ""
    try:
        with httpx.Client(
            headers={"User-Agent": UA},
            verify=False,
            timeout=httpx.Timeout(10.0, connect=4.0),
            follow_redirects=True,
        ) as client:
            for path in CONTACT_PATHS:
                try:
                    u = url if not path else urljoin(url + "/", path.lstrip("/"))
                    r = client.get(u)
                    if r.status_code >= 400:
                        last_err = f"HTTP {r.status_code}"
                        continue
                    p, e = extract_from_html(r.text or "", str(r.url))
                    phone = phone or p
                    email = email or e
                    if phone and email:
                        break
                except Exception as ex:  # noqa: BLE001
                    last_err = str(ex)[:120]
                    continue
    except Exception as ex:  # noqa: BLE001
        last_err = str(ex)[:120]
    return {"phone": phone, "email": email, "ok": bool(phone or email), "error": last_err, "url": url}


def phone_compatible(old: str, new: str) -> bool:
    """True if old looks like same number as new (digit compare)."""
    a = re.sub(r"\D", "", old or "")
    b = re.sub(r"\D", "", new or "")
    if not a or not b:
        return False
    # compare last 8+ digits
    return a[-8:] == b[-8:] or a in b or b in a


def main() -> None:
    # merge prior overrides
    prior = {}
    if OVERRIDES_PATH.exists():
        prior = json.loads(OVERRIDES_PATH.read_text(encoding="utf-8"))

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[2]]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}

    rows_meta = []
    for row in range(3, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        if not co:
            continue
        web = str(ws.cell(row, ic["Website"]).value or "").strip()
        # fix website
        if co in WEBSITE_FIXES:
            fixed = WEBSITE_FIXES[co]
            if fixed:
                ws.cell(row, ic["Website"]).value = fixed
                web = fixed
            else:
                ws.cell(row, ic["Website"]).value = NPD
                web = ""
        elif emptyish(web) or "w3schools" in web.lower() or "mysite" in web.lower():
            # leave for now unless in fixes
            if emptyish(web) or "w3schools" in web.lower():
                web = ""
        rows_meta.append(
            {
                "row": row,
                "company": co,
                "website": web,
                "cur_phone": str(ws.cell(row, ic["Office No."]).value or "").strip(),
                "cur_email": str(ws.cell(row, ic["Office Email"]).value or "").strip(),
            }
        )

    # harvest all with website
    cache: dict = {}
    todo = [(m["company"], m["website"]) for m in rows_meta if normalize_website(m["website"])]
    print(f"re-harvesting {len(todo)} websites ...", flush=True)
    done = 0
    with ThreadPoolExecutor(max_workers=14) as ex:
        futs = {ex.submit(fetch_contacts, web): (co, web) for co, web in todo}
        for fut in as_completed(futs):
            co, web = futs[fut]
            key = normalize_website(web)
            try:
                hit = fut.result()
            except Exception as exn:  # noqa: BLE001
                hit = {"phone": "", "email": "", "ok": False, "error": str(exn)[:120], "url": key}
            cache[key] = hit
            done += 1
            if done % 30 == 0 or done == len(todo):
                CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
                print(f"  {done}/{len(todo)} {co[:40]}", flush=True)

    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    stats = {
        "website_fixed": 0,
        "phone_filled": 0,
        "email_filled": 0,
        "phone_corrected": 0,
        "email_corrected": 0,
        "phone_kept_verified": 0,
        "email_kept_verified": 0,
        "phone_npd": 0,
        "email_npd": 0,
        "override_applied": 0,
    }

    for m in rows_meta:
        co = m["company"]
        web = str(ws.cell(m["row"], ic["Website"]).value or "").strip()
        if co in WEBSITE_FIXES and WEBSITE_FIXES[co]:
            stats["website_fixed"] += 1

        hit = cache.get(normalize_website(web), {})
        phone = ""
        email = ""

        # 1) agent overrides (highest trust)
        ov = CONTACT_OVERRIDES.get(co) or prior.get(co)
        if ov:
            if ov.get("office_no"):
                phone = ov["office_no"]
            if ov.get("office_email"):
                email = ov["office_email"]
            if co in CONTACT_OVERRIDES:
                stats["override_applied"] += 1

        # 2) harvest
        if not phone and hit.get("phone"):
            phone = hit["phone"]
        if not email and hit.get("email") and is_office_email(hit["email"]):
            email = hit["email"]

        # 3) keep existing if still valid / compatible with harvest
        cur_p = m["cur_phone"]
        cur_e = m["cur_email"]
        if not phone and cur_p and not emptyish(cur_p):
            # keep existing if harvest empty OR compatible
            if not hit.get("phone") or phone_compatible(cur_p, hit.get("phone", "")):
                phone = cur_p
                stats["phone_kept_verified"] += 1
            else:
                # harvest found different — prefer harvest for footer/contact
                phone = hit["phone"]
                stats["phone_corrected"] += 1
        elif phone and cur_p and not emptyish(cur_p) and not phone_compatible(cur_p, phone):
            stats["phone_corrected"] += 1

        if not email and cur_e and not emptyish(cur_e) and is_office_email(cur_e):
            if not hit.get("email") or clean_email(cur_e) == clean_email(hit.get("email", "")):
                email = clean_email(cur_e)
                stats["email_kept_verified"] += 1
            elif is_office_email(hit.get("email", "")):
                email = hit["email"]
                stats["email_corrected"] += 1
        elif email and cur_e and not emptyish(cur_e) and clean_email(cur_e) != clean_email(email):
            stats["email_corrected"] += 1

        # demote bad existing emails
        if email and not is_office_email(email):
            email = ""

        ws.cell(m["row"], ic["Office No."]).value = phone if phone else NPD
        ws.cell(m["row"], ic["Office Email"]).value = email if email else NPD
        if phone:
            stats["phone_filled"] += 1
        else:
            stats["phone_npd"] += 1
        if email:
            stats["email_filled"] += 1
        else:
            stats["email_npd"] += 1

    wb.save(XLSX)
    wb.close()

    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "file": str(XLSX),
        "stats": stats,
        "website_fixes": {k: v or NPD for k, v in WEBSITE_FIXES.items()},
        "contact_overrides": CONTACT_OVERRIDES,
        "rule": "Only public About/Contact/footer + verified sources; no invented phones/emails",
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(audit, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
