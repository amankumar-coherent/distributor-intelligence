# -*- coding: utf-8 -*-
"""Fill Office No. + Office Email for wearable medical devices FINAL.

1) Insert Office Email column after Office No. if missing
2) Scrape homepage + contact pages (tel:/mailto:/hint regex)
3) DeepSeek extracts ONLY values present in scraped text (no invent)
4) Clean junk phones (years, placeholders) and prefer office-style emails
5) Migrate verified office-style emails from Email → Office Email when needed
"""
from __future__ import annotations

import argparse
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
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

OUT = ROOT / "output/web_expand/global_wearable_medical_devices_market_global"
XLSX = OUT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
CACHE = OUT / "wearable_office_contact_cache.json"
AUDIT = OUT / "wearable_office_contact_audit.json"
SUMMARY = OUT / "wearable_office_contact_summary.json"

NPD = "Not publicly disclosed"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

CONTACT_PATHS = [
    "",
    "/contact",
    "/contact-us",
    "/contactus",
    "/contacts",
    "/contact-us/",
    "/en/contact",
    "/en/contact-us",
    "/us/contact",
    "/about/contact",
    "/company/contact",
    "/pages/contact",
    "/pages/contact-us",
    "/about",
    "/about-us",
    "/impressum",
    "/imprint",
    "/get-in-touch",
    "/customer-service",
    "/customer-care",
    "/support/contact",
    "/locations",
    "/our-locations",
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
TEL_HREF_RE = re.compile(r"tel:([+\d][\d\-().\s]{6,})", re.I)
MAILTO_RE = re.compile(r"mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", re.I)
HREF_RE = re.compile(r"""href=["']([^"']+)["']""", re.I)
JSON_TEL = re.compile(r"""["']telephone["']\s*:\s*["']([^"']+)["']""", re.I)
PHONE_HINT_RE = re.compile(
    r"(?:tel:|phone|tel\.|call|fax|t\s*[\(:]|office|contact|hotline|customer\s*service|"
    r"telefon|telephone|teléfono|telefone|contacto|freephone|toll[\s-]?free|"
    r"switchboard|hq|head\s*office)"
    r"[^0-9+]{0,40}(\+?[\d][\d\s().\-/]{7,22}\d)",
    re.I,
)

BAD_EMAIL_SUB = (
    "example.com",
    "exemple.com",
    "w3schools",
    "mysite.com",
    "sentry",
    "wixpress",
    "schema.org",
    "cloudflare",
    "noreply",
    "no-reply",
    "donotreply",
    "yourcompany",
    "domain.com",
    "email.com",
    "sentry.io",
    "github.com",
    "googleapis",
    "godaddy",
    "squarespace",
    "shopify.com",
)

OFFICE_LOCAL = re.compile(
    r"^(info|contact|sales|office|hello|enquiry|enquiries|support|customer|customerservice|"
    r"customerservices|customersupport|orders|export|mail|admin|team|trade|help|press|media|"
    r"technical|inquiry|kontakt|infouk|communications|customercare|service|promo|"
    r"reception|general|wholesale|distributor|distributors|procurement|pharmacy|medical|"
    r"healthcare|care|hospital|pharma|dataprotection|privacy|"
    r"info\.|sales\.|contact\.)([._-]|$)",
    re.I,
)

JUNK_PHONE_DIGITS = {
    "6666666667",
    "6666666666",
    "1234567890",
    "0000000000",
    "1111111111",
    "0123456789",
    "123456789",
    "12345678",
}


def emptyish(v: object) -> bool:
    s = str(v or "").strip().lower()
    return s in {
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


def split_phone_parts(raw: str) -> list[str]:
    s = str(raw or "").strip()
    if not s:
        return []
    parts = re.split(r"\s*[;|]\s*|\s+/\s+|\s+or\s+", s, flags=re.I)
    return [p.strip() for p in parts if p and p.strip()]


def clean_phone_one(raw: str) -> str:
    s = re.sub(r"\s+", " ", str(raw or "").strip())
    s = s.replace("tel:", "").replace("TEL:", "").replace("%20", " ")
    s = s.strip(" \t.;,|/")
    # drop trailing extension labels kept short
    s = re.sub(r"\s*(ext\.?|extension|x)\s*\d+$", "", s, flags=re.I).strip()
    digits = re.sub(r"\D", "", s)
    if len(digits) < 8 or len(digits) > 15:
        return ""
    if digits in JUNK_PHONE_DIGITS:
        return ""
    if len(set(digits)) <= 2:
        return ""
    # filter years / dates scraped as phones
    if re.fullmatch(r"(19|20)\d{2}", digits):
        return ""
    if re.search(r"(19|20)\d{2}", s) and len(digits) <= 8:
        return ""
    # reject obvious date-like fragments
    if re.search(r"\b(19|20)\d{2}[-/.]\d{1,2}[-/.]\d{1,2}\b", s):
        return ""
    # bare 4-digit years disguised inside longer scrap
    if re.fullmatch(r"(19|20)\d{2}\d{0,4}", digits) and len(digits) <= 8:
        return ""
    return s[:48]


def clean_phone(raw: str) -> str:
    """Return first good phone, or semicolon-joined list when multiple valid parts."""
    parts = split_phone_parts(raw)
    if not parts:
        one = clean_phone_one(raw)
        return one
    good: list[str] = []
    seen: set[str] = set()
    for part in parts:
        c = clean_phone_one(part)
        if not c:
            continue
        key = re.sub(r"\D", "", c)
        if key in seen:
            continue
        seen.add(key)
        good.append(c)
    if not good:
        return clean_phone_one(raw)
    return "; ".join(good[:3])


def clean_email(raw: str) -> str:
    e = str(raw or "").strip().lower().lstrip("%20").rstrip(".,;)>")
    m = re.search(r"([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})", e)
    if m:
        e = m.group(1)
    if not e or "@" not in e:
        return ""
    if any(b in e for b in BAD_EMAIL_SUB):
        return ""
    if e.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js")):
        return ""
    local, _, dom = e.partition("@")
    if not local or not dom or "." not in dom:
        return ""
    if len(local) > 64 or len(dom) > 120:
        return ""
    return e


def is_office_email(email: str) -> bool:
    e = clean_email(email)
    if not e:
        return False
    local = e.split("@")[0]
    if OFFICE_LOCAL.search(local):
        return True
    return bool(
        re.search(
            r"(info|contact|sales|office|support|enquir|customer|hello|export|orders|"
            r"service|wholesale|pharmacy|medical|healthcare|pharma|care|trade|data.?protection)",
            local,
            re.I,
        )
    )


def is_junk_phone(phone: str) -> bool:
    if emptyish(phone):
        return True
    if "@" in phone:
        return True
    cleaned = clean_phone(phone)
    if not cleaned:
        return True
    # multi: junk only if no valid part survived
    for part in split_phone_parts(cleaned) or [cleaned]:
        digits = re.sub(r"\D", "", part)
        if digits in JUNK_PHONE_DIGITS:
            continue
        if len(set(digits)) <= 2:
            continue
        if re.search(r"(19|20)\d{2}", part) and len(digits) <= 10:
            continue
        if 8 <= len(digits) <= 15:
            return False
    return True


def format_phone_one(phone: str) -> str:
    """Light normalize one phone for display; keep original country style when possible."""
    p = clean_phone_one(phone)
    if not p:
        return ""
    p = re.sub(r"\s+", " ", p).strip()
    # US 11-digit starting with 1 → 1-XXX-XXX-XXXX if bare digits
    d = re.sub(r"\D", "", p)
    bare = p.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    if re.fullmatch(r"\d+", bare):
        if len(d) == 11 and d.startswith("1"):
            return f"1-{d[1:4]}-{d[4:7]}-{d[7:]}"
        if len(d) == 10:
            return f"({d[0:3]}) {d[3:6]}-{d[6:]}"
    return p


def format_phone_display(phone: str) -> str:
    parts = split_phone_parts(phone)
    if len(parts) <= 1:
        return format_phone_one(phone) or clean_phone(phone)
    out: list[str] = []
    seen: set[str] = set()
    for part in parts:
        fmt = format_phone_one(part)
        if not fmt:
            continue
        key = re.sub(r"\D", "", fmt)
        if key in seen:
            continue
        seen.add(key)
        out.append(fmt)
    return "; ".join(out[:3])


def html_to_text(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<noscript.*?>.*?</noscript>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    t = re.sub(r"&\w+;", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def discover_contact_urls(base: str, html: str) -> list[str]:
    out: list[str] = []
    seen = set()
    host = urlparse(base).netloc.lower().replace("www.", "")
    keys = (
        "contact",
        "contacto",
        "kontakt",
        "impressum",
        "imprint",
        "get-in-touch",
        "customer-service",
        "customer-care",
        "locations",
        "find-us",
        "about-us",
    )
    for href in HREF_RE.findall(html or ""):
        h = href.strip()
        if not h or h.startswith(("#", "javascript:", "mailto:", "tel:")):
            continue
        low = h.lower()
        if not any(k in low for k in keys):
            continue
        full = urljoin(base + "/", h)
        p = urlparse(full)
        if p.scheme not in {"http", "https"}:
            continue
        nh = p.netloc.lower().replace("www.", "")
        if host and nh and host not in nh and nh not in host:
            continue
        key = full.split("#")[0].rstrip("/")
        if key in seen:
            continue
        seen.add(key)
        out.append(key)
        if len(out) >= 12:
            break
    return out


def uniq_list(xs: list[str], phone: bool = False) -> list[str]:
    seen = set()
    out = []
    for x in xs:
        k = re.sub(r"\D", "", x) if phone else x.lower()
        if not k or k in seen:
            continue
        seen.add(k)
        out.append(x)
    return out


def extract_from_html(html: str) -> tuple[list[str], list[str], str]:
    phones: list[str] = []
    emails: list[str] = []
    for m in TEL_HREF_RE.finditer(html):
        p = clean_phone(m.group(1))
        if p:
            phones.append(p)
    for m in JSON_TEL.finditer(html):
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
    for m in PHONE_HINT_RE.finditer(html):
        p = clean_phone(m.group(1))
        if p:
            phones.append(p)
    for m in PHONE_HINT_RE.finditer(txt):
        p = clean_phone(m.group(1))
        if p:
            phones.append(p)
    return uniq_list(phones, phone=True), uniq_list(emails), txt


def pick_best(phones: list[str], emails: list[str], base_url: str) -> tuple[str, str]:
    phone = phones[0] if phones else ""
    host = urlparse(base_url).netloc.lower().removeprefix("www.").split(":")[0]
    domain = [
        e
        for e in emails
        if host and (host in e.split("@")[-1] or e.split("@")[-1] in host)
    ]
    prefer_locals = {
        "info",
        "contact",
        "sales",
        "office",
        "hello",
        "enquiries",
        "enquiry",
        "support",
        "customerservice",
        "customerservices",
        "orders",
        "export",
        "wholesale",
        "customercare",
        "pharmacy",
        "medical",
        "healthcare",
        "pharma",
        "care",
        "dataprotection",
    }
    pool = domain or emails
    prefer = [e for e in pool if e.split("@")[0].split("+")[0] in prefer_locals]
    email = (prefer or pool)[0] if pool else ""
    return phone, email


def scrape_site(website: str) -> dict:
    url = normalize_website(website)
    if not url:
        return {"phone": "", "email": "", "text": "", "url": "", "ok": False, "pages": 0}
    phones: list[str] = []
    emails: list[str] = []
    texts: list[str] = []
    pages = 0
    try:
        with httpx.Client(
            headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"},
            verify=False,
            timeout=httpx.Timeout(12.0, connect=5.0),
            follow_redirects=True,
        ) as client:
            try:
                home = client.get(url)
                if home.status_code < 400:
                    pages += 1
                    p, e, t = extract_from_html(home.text or "")
                    phones.extend(p)
                    emails.extend(e)
                    if t:
                        texts.append(t[:5000])
                    extra = discover_contact_urls(str(home.url).rstrip("/"), home.text or "")
                else:
                    extra = []
            except Exception:
                extra = []

            urls = [urljoin(url + "/", path.lstrip("/")) for path in CONTACT_PATHS[1:]]
            urls.extend(extra)
            seen_u = {url.rstrip("/")}
            for u in urls:
                key = u.split("#")[0].rstrip("/")
                if key in seen_u:
                    continue
                seen_u.add(key)
                try:
                    r = client.get(u)
                    if r.status_code >= 400:
                        continue
                    pages += 1
                    p, e, t = extract_from_html(r.text or "")
                    phones.extend(p)
                    emails.extend(e)
                    if t:
                        texts.append(t[:5000])
                    if phones and emails and pages >= 3:
                        break
                except Exception:
                    continue
    except Exception as ex:  # noqa: BLE001
        return {
            "phone": "",
            "email": "",
            "text": "",
            "url": url,
            "ok": False,
            "pages": 0,
            "error": str(ex)[:120],
        }

    phones = uniq_list(phones, phone=True)
    emails = uniq_list(emails)
    phone, email = pick_best(phones, emails, url)
    return {
        "phone": phone,
        "email": email,
        "phones": phones[:8],
        "emails": emails[:8],
        "text": "\n---\n".join(texts)[:14000],
        "url": url,
        "ok": bool(phone or email or texts),
        "pages": pages,
    }


def ds_client() -> tuple[OpenAI | None, str]:
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key:
        return None, model
    return OpenAI(api_key=key, base_url=base, timeout=60.0), model


def _parse_json(text: str) -> dict:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I | re.S).strip()
    try:
        return json.loads(raw)
    except Exception:
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return {}
        return {}


def deepseek_extract(
    client: OpenAI, model: str, company: str, website: str, page_text: str
) -> dict:
    if not (page_text or "").strip():
        return {"phone": "", "email": ""}
    system = (
        "Extract PUBLIC office switchboard phone and office email from website text only. "
        "Prefer info@/sales@/contact@/support@/customerservice@. "
        "Reject personal emails unless clearly listed as company contact. "
        "Reject years/dates as phones. Do NOT invent. If not clearly in the text, found_*=false. "
        "Return ONLY JSON."
    )
    user = (
        f"Company: {company}\nWebsite: {website}\n\n"
        f"TEXT:\n{page_text[:9000]}\n\n"
        '{"found_phone":true|false,"office_phone":"","found_email":true|false,'
        '"office_email":"","confidence":"high|medium|low"}'
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=300,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        data = _parse_json(resp.choices[0].message.content or "{}")
    except Exception as ex:  # noqa: BLE001
        return {"phone": "", "email": "", "error": str(ex)[:120]}

    conf = str(data.get("confidence") or "").lower()
    out = {"phone": "", "email": "", "confidence": conf, "raw": data}
    if conf not in {"high", "medium"}:
        return out

    digits_page = re.sub(r"\D", "", page_text)
    text_l = page_text.lower()

    if data.get("found_phone"):
        phone = clean_phone(str(data.get("office_phone") or ""))
        if phone:
            d = re.sub(r"\D", "", phone)
            if len(d) >= 8 and d[-7:] in digits_page:
                out["phone"] = phone

    if data.get("found_email"):
        email = clean_email(str(data.get("office_email") or ""))
        if email and email in text_l and is_office_email(email):
            out["email"] = email
    return out


def ensure_office_email_column(ws) -> dict[str, int]:
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    if "Office Email" not in headers:
        oi = headers.index("Office No.") + 1  # 1-based
        ws.insert_cols(oi + 1)
        cell = ws.cell(1, oi + 1)
        cell.value = "Office Email"
        cell.font = Font(bold=True)
        headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        print("inserted Office Email after Office No.")
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}
    return ic


def style_wrap(ws, ic: dict[str, int]) -> None:
    for col_name in ("Office No.", "Office Email"):
        if col_name not in ic:
            continue
        col = ic[col_name]
        ws.column_dimensions[get_column_letter(col)].width = 28
        for r in range(2, ws.max_row + 1):
            ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical="top")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="Max gap rows to process (0=all)")
    ap.add_argument("--workers", type=int, default=10)
    ap.add_argument("--no-deepseek", action="store_true")
    ap.add_argument("--resume", action="store_true", help="Reuse cache")
    args = ap.parse_args()

    client, model = (None, "")
    if not args.no_deepseek:
        client, model = ds_client()
        if client:
            print(f"DeepSeek model={model}", flush=True)
        else:
            print("DeepSeek key missing — scrape-only mode", flush=True)

    cache: dict = {}
    if args.resume and CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    # backup
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(XLSX.stem + f"_BEFORE_OFFICE_{stamp}.xlsx")
    wb0 = openpyxl.load_workbook(XLSX)
    wb0.save(backup)
    wb0.close()
    print(f"backup {backup.name}", flush=True)

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    ic = ensure_office_email_column(ws)

    # migrate office-style emails from personal Email column → Office Email
    migrated = 0
    if "Email" in ic:
        for r in range(2, ws.max_row + 1):
            oe = str(ws.cell(r, ic["Office Email"]).value or "").strip()
            if not emptyish(oe):
                continue
            raw_em = str(ws.cell(r, ic["Email"]).value or "").strip()
            if emptyish(raw_em):
                continue
            cands = [clean_email(x) for x in re.split(r"[;,\s]+", raw_em) if "@" in x]
            office = next((e for e in cands if e and is_office_email(e)), "")
            if office:
                ws.cell(r, ic["Office Email"]).value = office
                migrated += 1
    print(f"migrated_email_to_office_email={migrated}", flush=True)

    # clean existing junk phones
    cleaned = 0
    reformatted = 0
    for r in range(2, ws.max_row + 1):
        ph = str(ws.cell(r, ic["Office No."]).value or "").strip()
        if emptyish(ph):
            continue
        if is_junk_phone(ph):
            ws.cell(r, ic["Office No."]).value = NPD
            cleaned += 1
            continue
        fmt = format_phone_display(ph)
        if fmt and fmt != ph:
            ws.cell(r, ic["Office No."]).value = fmt
            reformatted += 1
        elif fmt:
            ws.cell(r, ic["Office No."]).value = fmt
    print(f"cleaned_junk_phones={cleaned} reformatted={reformatted}", flush=True)

    gaps = []
    for row in range(2, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        if not co:
            continue
        web = str(ws.cell(row, ic["Website"]).value or "").strip()
        ph = str(ws.cell(row, ic["Office No."]).value or "").strip()
        em = str(ws.cell(row, ic["Office Email"]).value or "").strip()
        need_p = emptyish(ph) or is_junk_phone(ph)
        need_e = emptyish(em)
        if need_p or need_e:
            gaps.append(
                {
                    "row": row,
                    "company": co,
                    "website": web,
                    "need_phone": need_p,
                    "need_email": need_e,
                }
            )

    if args.limit and args.limit > 0:
        gaps = gaps[: args.limit]
    print(f"gap rows={len(gaps)}", flush=True)

    todo_webs = []
    seen_w = set()
    for g in gaps:
        web = normalize_website(g["website"])
        if web and web not in seen_w:
            seen_w.add(web)
            if args.resume and web in cache and cache[web].get("ok") is not None:
                continue
            todo_webs.append(web)

    print(f"scraping {len(todo_webs)} unique sites (workers={args.workers}) ...", flush=True)
    scrape_map: dict[str, dict] = {
        k: v for k, v in cache.items() if isinstance(v, dict)
    }
    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        futs = {ex.submit(scrape_site, web): web for web in todo_webs}
        done = 0
        for fut in as_completed(futs):
            web = futs[fut]
            try:
                scrape_map[web] = fut.result()
            except Exception as exn:  # noqa: BLE001
                scrape_map[web] = {
                    "phone": "",
                    "email": "",
                    "text": "",
                    "url": web,
                    "ok": False,
                    "error": str(exn)[:120],
                }
            done += 1
            if done % 25 == 0 or done == len(todo_webs):
                print(f"  scrape {done}/{len(todo_webs)}", flush=True)
                CACHE.write_text(
                    json.dumps(scrape_map, indent=2, ensure_ascii=False), encoding="utf-8"
                )

    CACHE.write_text(json.dumps(scrape_map, indent=2, ensure_ascii=False), encoding="utf-8")

    audit = []
    filled_phone = 0
    filled_email = 0
    ds_used = 0

    for i, g in enumerate(gaps, 1):
        web = normalize_website(g["website"])
        sc = scrape_map.get(web) or {}
        phone = clean_phone(sc.get("phone") or "")
        email = clean_email(sc.get("email") or "")
        if email and not is_office_email(email):
            email = ""

        # DeepSeek only if still missing and we have page text
        if client and ((g["need_phone"] and not phone) or (g["need_email"] and not email)):
            text = sc.get("text") or ""
            if text.strip():
                ds = deepseek_extract(client, model, g["company"], web, text)
                ds_used += 1
                if g["need_phone"] and not phone and ds.get("phone"):
                    phone = clean_phone(ds["phone"])
                if g["need_email"] and not email and ds.get("email"):
                    e2 = clean_email(ds["email"])
                    if e2 and is_office_email(e2):
                        email = e2
                # light throttle
                if ds_used % 20 == 0:
                    time.sleep(0.4)

        row = g["row"]
        note = {"company": g["company"], "website": web, "row": row}

        if g["need_phone"]:
            if phone and not is_junk_phone(phone):
                ws.cell(row, ic["Office No."]).value = format_phone_display(phone)
                filled_phone += 1
                note["phone"] = format_phone_display(phone)
            else:
                ws.cell(row, ic["Office No."]).value = NPD
                note["phone"] = NPD

        if g["need_email"]:
            if email:
                ws.cell(row, ic["Office Email"]).value = email
                filled_email += 1
                note["email"] = email
            else:
                ws.cell(row, ic["Office Email"]).value = NPD
                note["email"] = NPD

        # ensure non-gap cells that were empty Office Email get NPD
        audit.append(note)
        if i % 50 == 0:
            print(f"  apply {i}/{len(gaps)} phone+={filled_phone} email+={filled_email}", flush=True)

    # fill Office Email NPD for any remaining blanks
    for r in range(2, ws.max_row + 1):
        if emptyish(ws.cell(r, ic["Office Email"]).value):
            ws.cell(r, ic["Office Email"]).value = NPD
        if emptyish(ws.cell(r, ic["Office No."]).value):
            ws.cell(r, ic["Office No."]).value = NPD

    style_wrap(ws, ic)
    wb.save(XLSX)
    wb.close()

    # recount
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    oi = headers.index("Office No.")
    ei = headers.index("Office Email")
    phone_n = email_n = 0
    total = ws.max_row - 1
    for row in ws.iter_rows(min_row=2, values_only=True):
        ph = str(row[oi] or "").strip()
        em = str(row[ei] or "").strip()
        if ph and ph.lower() != NPD.lower() and not is_junk_phone(ph):
            phone_n += 1
        if em and em.lower() != NPD.lower() and "@" in em:
            email_n += 1
    wb.close()

    summary = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "file": str(XLSX),
        "backup": str(backup),
        "total_rows": total,
        "office_no_filled": phone_n,
        "office_email_filled": email_n,
        "filled_phone_this_run": filled_phone,
        "filled_email_this_run": filled_email,
        "deepseek_calls": ds_used,
        "gaps_processed": len(gaps),
    }
    SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    AUDIT.write_text(json.dumps(audit[:500], indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=2), flush=True)


if __name__ == "__main__":
    # silence insecure request warnings from verify=False
    try:
        import urllib3

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    except Exception:
        pass
    main()
