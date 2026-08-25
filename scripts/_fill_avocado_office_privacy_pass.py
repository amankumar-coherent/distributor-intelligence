#!/usr/bin/env python3
"""Second pass: privacy-policy + contact discovery for remaining avocado NPD emails/phones."""
from __future__ import annotations

import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse

import httpx
import openpyxl
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl.styles import Alignment

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")
OUT = ROOT / "output/web_expand/global_avocado_oil_market_global"
XLSX = OUT / "global_avocado_oil_market_OFFICE_SLIM.xlsx"
CACHE = OUT / "_avocado_privacy_contact_cache.json"
AUDIT = OUT / "_avocado_privacy_contact_audit.json"
NPD = "Not publicly disclosed"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

PATHS = [
    "",
    "/contact",
    "/contact-us",
    "/contactus",
    "/en/contact",
    "/about-us/contact",
    "/pages/contact",
    "/privacy",
    "/privacy-policy",
    "/policies/privacy-policy",
    "/legal/privacy",
    "/en/privacy-policy",
    "/impressum",
    "/imprint",
    "/customer-service",
    "/support",
    "/pages/support",
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
TEL_HREF_RE = re.compile(r"tel:([+\d][\d\-().\s]{6,})", re.I)
MAILTO_RE = re.compile(r"mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", re.I)
HREF_RE = re.compile(r"""href=["']([^"']+)["']""", re.I)

BAD_EMAIL_SUB = (
    "example.com",
    "sentry",
    "wixpress",
    "schema.org",
    "cloudflare",
    "noreply",
    "no-reply",
    "donotreply",
    "w3schools",
    "googleapis",
    "github.com",
)

OFFICE_OK = re.compile(
    r"(info|contact|sales|office|hello|enquiry|enquiries|support|customer|orders|"
    r"export|wholesale|customercare|service|reception|general|privacy|legal|"
    r"communications|media|press|help)([._+-]|$)",
    re.I,
)


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in {
        "",
        "n/a",
        "na",
        "none",
        "-",
        "—",
        "not publicly disclosed",
        "unknown",
        "tbd",
        "null",
    }


def normalize_website(url: str) -> str:
    u = (url or "").strip()
    if emptyish(u):
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    return u.rstrip("/")


def clean_email(raw: str) -> str:
    e = str(raw or "").strip().lower().rstrip(".,;)>")
    m = EMAIL_RE.search(e)
    if not m:
        return ""
    e = m.group(0)
    if any(b in e for b in BAD_EMAIL_SUB):
        return ""
    if e.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js")):
        return ""
    return e


def is_office_email(email: str) -> bool:
    e = clean_email(email)
    if not e:
        return False
    return bool(OFFICE_OK.search(e.split("@")[0]))


def clean_phone(raw: str) -> str:
    s = re.sub(r"\s+", " ", str(raw or "").strip()).replace("tel:", "").strip()
    digits = re.sub(r"\D", "", s)
    if len(digits) < 8 or len(digits) > 15:
        return ""
    if len(set(digits)) <= 2:
        return ""
    if re.search(r"20\d{2}\.\d{2}\.\d{2}", s):
        return ""
    return s[:40]


def html_to_text(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    t = re.sub(r"&\w+;", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def scrape(website: str) -> dict:
    url = normalize_website(website)
    if not url:
        return {"phone": "", "email": "", "text": "", "ok": False}
    phones, emails, texts = [], [], []
    try:
        with httpx.Client(
            headers={"User-Agent": UA},
            verify=False,
            timeout=httpx.Timeout(12.0, connect=5.0),
            follow_redirects=True,
        ) as client:
            home_html = ""
            try:
                r = client.get(url)
                if r.status_code < 400:
                    home_html = r.text or ""
                    base = str(r.url).rstrip("/")
                else:
                    base = url
            except Exception:
                base = url
                home_html = ""

            extra = []
            for href in HREF_RE.findall(home_html):
                low = href.lower()
                if any(k in low for k in ("contact", "privacy", "impressum", "imprint", "support", "customer")):
                    full = urljoin(base + "/", href)
                    if urlparse(full).scheme in {"http", "https"}:
                        extra.append(full.split("#")[0])

            urls = [urljoin(url + "/", p.lstrip("/")) for p in PATHS]
            urls = [url] + urls[1:] + extra[:10]
            seen = set()
            for u in urls:
                key = u.rstrip("/")
                if key in seen:
                    continue
                seen.add(key)
                try:
                    r = client.get(u)
                    if r.status_code >= 400:
                        continue
                    html = r.text or ""
                    for m in TEL_HREF_RE.finditer(html):
                        p = clean_phone(m.group(1))
                        if p:
                            phones.append(p)
                    for m in MAILTO_RE.finditer(html):
                        e = clean_email(m.group(1))
                        if e and is_office_email(e):
                            emails.append(e)
                    for m in EMAIL_RE.finditer(html):
                        e = clean_email(m.group(0))
                        if e and is_office_email(e):
                            emails.append(e)
                    txt = html_to_text(html)
                    if txt:
                        texts.append(txt[:4500])
                    if phones and emails and len(texts) >= 2:
                        break
                except Exception:
                    continue
    except Exception as ex:  # noqa: BLE001
        return {"phone": "", "email": "", "text": "", "ok": False, "error": str(ex)[:100]}

    def uniq(xs, phone=False):
        seen, out = set(), []
        for x in xs:
            k = re.sub(r"\D", "", x) if phone else x.lower()
            if k in seen:
                continue
            seen.add(k)
            out.append(x)
        return out

    phones, emails = uniq(phones, True), uniq(emails)
    host = urlparse(url).netloc.lower().removeprefix("www.")
    domain = [e for e in emails if host.split(":")[0] in e.split("@")[-1]]
    prefer = [
        e
        for e in (domain or emails)
        if e.split("@")[0].split("+")[0]
        in {
            "info",
            "contact",
            "sales",
            "hello",
            "support",
            "customerservice",
            "enquiries",
            "office",
            "orders",
            "export",
            "wholesale",
        }
    ]
    # allow privacy@ only if nothing better
    email = (prefer or domain or emails)[0] if (prefer or domain or emails) else ""
    if email.startswith("privacy@") and prefer:
        email = prefer[0]
    return {
        "phone": phones[0] if phones else "",
        "email": email,
        "text": "\n---\n".join(texts)[:12000],
        "ok": bool(phones or emails or texts),
        "emails": emails[:8],
        "phones": phones[:8],
    }


def ds_client():
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    return OpenAI(api_key=key, base_url=base, timeout=60.0), model


def ds_extract(client, model, company, website, text):
    if not text.strip():
        return {"phone": "", "email": ""}
    prompt = (
        f"Company: {company}\nWebsite: {website}\n\nTEXT:\n{text[:8500]}\n\n"
        "Extract public office phone and office email ONLY if present in TEXT. "
        "Prefer info@/sales@/contact@/hello@/support@. Never invent. JSON only:\n"
        '{"found_phone":true|false,"office_phone":"","found_email":true|false,'
        '"office_email":"","confidence":"high|medium|low"}'
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=250,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "Extract contacts from text only. No invention. JSON only."},
                {"role": "user", "content": prompt},
            ],
        )
        raw = resp.choices[0].message.content or "{}"
        data = json.loads(raw)
    except Exception as ex:  # noqa: BLE001
        return {"phone": "", "email": "", "error": str(ex)[:100]}
    conf = str(data.get("confidence") or "").lower()
    out = {"phone": "", "email": "", "confidence": conf}
    if conf not in {"high", "medium"}:
        return out
    digits = re.sub(r"\D", "", text)
    if data.get("found_phone"):
        p = clean_phone(str(data.get("office_phone") or ""))
        if p and re.sub(r"\D", "", p)[-7:] in digits:
            out["phone"] = p
    if data.get("found_email"):
        e = clean_email(str(data.get("office_email") or ""))
        if e and e in text.lower() and is_office_email(e):
            # skip privacy@ if better options likely; still allow if only one
            out["email"] = e
    return out


def main():
    client, model = ds_client()
    cache = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}

    gaps = []
    for row in range(3, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        if not co:
            continue
        web = str(ws.cell(row, ic["Website"]).value or "").strip()
        ph = str(ws.cell(row, ic["Office No."]).value or "").strip()
        em = str(ws.cell(row, ic["Office Email"]).value or "").strip()
        if emptyish(ph) or emptyish(em):
            gaps.append(
                {
                    "row": row,
                    "company": co,
                    "website": web,
                    "need_p": emptyish(ph),
                    "need_e": emptyish(em),
                }
            )

    print(f"gaps={len(gaps)}", flush=True)
    scrape_map = {}
    todo = [(g["company"], normalize_website(g["website"])) for g in gaps if normalize_website(g["website"])]
    with ThreadPoolExecutor(max_workers=10) as ex:
        futs = {ex.submit(scrape, web): (co, web) for co, web in todo}
        done = 0
        for fut in as_completed(futs):
            co, web = futs[fut]
            try:
                scrape_map[web] = fut.result()
            except Exception as e:  # noqa: BLE001
                scrape_map[web] = {"phone": "", "email": "", "text": "", "ok": False, "error": str(e)[:80]}
            done += 1
            if done % 25 == 0 or done == len(todo):
                print(f" scraped {done}/{len(todo)}", flush=True)

    filled_p = filled_e = ds_calls = 0
    details = []
    for i, g in enumerate(gaps, 1):
        web = normalize_website(g["website"])
        scrap = scrape_map.get(web, {})
        phone = scrap.get("phone") or ""
        email = scrap.get("email") or ""
        # prefer non-privacy when multiple
        for e in scrap.get("emails") or []:
            if e.split("@")[0] in {"info", "sales", "contact", "hello", "support", "customerservice"}:
                email = e
                break

        still = (g["need_p"] and not phone) or (g["need_e"] and not email)
        key = f"ds2::{g['company']}::{web}"
        ds_hit = cache.get(key) or {}
        if still and scrap.get("text") and not (ds_hit.get("phone") or ds_hit.get("email")):
            print(f"[{i}/{len(gaps)}] DS {g['company'][:42]}", flush=True)
            ds_hit = ds_extract(client, model, g["company"], web, scrap.get("text") or "")
            cache[key] = ds_hit
            ds_calls += 1
            if ds_calls % 20 == 0:
                CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
            time.sleep(0.08)

        if g["need_p"] and not phone and ds_hit.get("phone"):
            phone = ds_hit["phone"]
        if g["need_e"] and not email and ds_hit.get("email") and is_office_email(ds_hit["email"]):
            email = ds_hit["email"]

        changed = {}
        if g["need_p"] and phone:
            ws.cell(g["row"], ic["Office No."]).value = phone
            filled_p += 1
            changed["phone"] = phone
        elif g["need_p"]:
            ws.cell(g["row"], ic["Office No."]).value = NPD
        if g["need_e"] and email and is_office_email(email):
            # demote privacy@ if that's all we got — still a public contact; keep
            ws.cell(g["row"], ic["Office Email"]).value = email
            filled_e += 1
            changed["email"] = email
        elif g["need_e"]:
            ws.cell(g["row"], ic["Office Email"]).value = NPD
        if changed:
            details.append({"company": g["company"], **changed})
            for col in ("Office No.", "Office Email"):
                ws.cell(g["row"], ic[col]).alignment = Alignment(wrap_text=True, vertical="top")

    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    wb.save(XLSX)
    wb.close()

    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    ic2 = {str(h): i for i, h in enumerate(hdr) if h}
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if str(r[ic2["Company"]] or "").strip()]
    pn = sum(1 for r in rows if not emptyish(r[ic2["Office No."]]))
    en = sum(1 for r in rows if not emptyish(r[ic2["Office Email"]]))
    wb.close()

    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "gaps": len(gaps),
        "ds_calls": ds_calls,
        "newly_filled_phone": filled_p,
        "newly_filled_email": filled_e,
        "total_phone_filled": pn,
        "total_email_filled": en,
        "remaining_phone_npd": len(rows) - pn,
        "remaining_email_npd": len(rows) - en,
        "samples": details[:40],
        "model": model,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: audit[k] for k in audit if k != "samples"}, indent=2))
    for d in details[:30]:
        print("+", d)


if __name__ == "__main__":
    main()
