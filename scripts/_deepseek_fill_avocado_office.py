#!/usr/bin/env python3
"""Fill more avocado Office No. / Office Email gaps using DeepSeek + website scrape.

Rules:
- Prefer numbers/emails that appear in scraped About/Contact/footer HTML.
- DeepSeek may suggest a corrected official website for junk URLs; we re-fetch.
- Never invent: if not found / low confidence → leave Not publicly disclosed.
"""
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

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

OUT_DIR = ROOT / "output/web_expand/global_avocado_oil_market_global"
XLSX = OUT_DIR / "global_avocado_oil_market_OFFICE_SLIM.xlsx"
CACHE = OUT_DIR / "_avocado_deepseek_office_cache.json"
AUDIT = OUT_DIR / "_avocado_deepseek_office_audit.json"

NPD = "Not publicly disclosed"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

JUNK_HOST_PARTS = (
    "w3schools",
    "mysite.com",
    "microsoft.com",
    "juraforum",
    "example.com",
    "exemple.com",
    "google.com",
    "facebook.com",
    "linkedin.com",
    "wikipedia.org",
    "financebuzz",
)

CONTACT_PATHS = [
    "",
    "/contact",
    "/contact-us",
    "/contactus",
    "/contacts",
    "/about",
    "/about-us",
    "/impressum",
    "/imprint",
    "/en/contact",
    "/pages/contact",
]

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(?:\+|00)?[\d][\d\s().\-/]{7,22}\d")
TEL_HREF_RE = re.compile(r"tel:([+\d][\d\-().\s]{6,})", re.I)
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
    "noreply",
    "no-reply",
    "donotreply",
)

OFFICE_LOCAL = re.compile(
    r"^(info|contact|sales|office|hello|enquiry|enquiries|support|customer|customerservice|"
    r"customerservices|customersupport|orders|export|mail|admin|team|trade|help|press|media|"
    r"technical|inquiry|kontakt|infouk|communications|customercare|service|promo)([._-]|$)",
    re.I,
)


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


def is_junk_website(url: str) -> bool:
    u = (url or "").lower()
    if emptyish(u):
        return True
    return any(j in u for j in JUNK_HOST_PARTS)


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
    return bool(
        re.search(
            r"(info|contact|sales|office|support|enquir|customer|hello|export|orders|service)",
            local,
            re.I,
        )
    )


def html_to_text(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    t = re.sub(r"&\w+;", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


def scrape_site(website: str) -> dict:
    url = normalize_website(website)
    if not url or is_junk_website(url):
        return {"phone": "", "email": "", "text": "", "url": url, "ok": False}
    phones: list[str] = []
    emails: list[str] = []
    texts: list[str] = []
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
                        texts.append(txt[:4000])
                    # also phone-like in text near contact words
                    for m in re.finditer(
                        r"(?:phone|tel|call|office|contact|telefon|telephone|fax)[^0-9+]{0,24}"
                        r"(\+?[\d][\d\s().\-/]{7,22}\d)",
                        txt,
                        re.I,
                    ):
                        p = clean_phone(m.group(1))
                        if p:
                            phones.append(p)
                    if phones and emails:
                        break
                except Exception:
                    continue
    except Exception:
        return {"phone": "", "email": "", "text": "", "url": url, "ok": False}

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
    host = urlparse(url).netloc.lower().removeprefix("www.").split(":")[0]
    domain = [e for e in ems if host and host in e.split("@")[-1]]
    email = (domain or ems)[0] if (domain or ems) else ""
    return {
        "phone": phone,
        "email": email,
        "text": "\n---\n".join(texts)[:12000],
        "url": url,
        "ok": bool(phone or email or texts),
    }


def ds_client() -> tuple[OpenAI, str]:
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key:
        raise SystemExit("OPENAI_API_KEY missing in D:\\vendor-intel\\.env")
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


def deepseek_fix_website(client: OpenAI, model: str, company: str, hq: str, bad_web: str) -> str:
    system = (
        "Return ONLY JSON. Find the company's official corporate website homepage URL. "
        "If unsure, found=false. Never invent a URL."
    )
    user = (
        f"Company: {company}\nHQ: {hq or 'unknown'}\nCurrent bad/missing URL: {bad_web or 'none'}\n"
        'JSON: {"found":true|false,"website":"https://...","confidence":"high|medium|low"}'
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=200,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        data = _parse_json(resp.choices[0].message.content or "{}")
    except Exception:
        return ""
    if not data.get("found"):
        return ""
    if str(data.get("confidence") or "").lower() != "high":
        return ""
    web = normalize_website(str(data.get("website") or ""))
    if not web or is_junk_website(web):
        return ""
    return web


def deepseek_extract_office(
    client: OpenAI,
    model: str,
    company: str,
    website: str,
    hq: str,
    page_text: str,
) -> dict:
    """Extract office phone/email. Prefer values present in page_text."""
    system = (
        "Extract the company's PUBLIC OFFICE phone and OFFICE email (general info/sales/contact). "
        "Prefer values that appear in the provided website text. "
        "Do NOT invent emails or phone numbers. "
        "Do NOT return personal employee emails like john.smith@... unless that is clearly the "
        "published company office contact on the page. Prefer info@, sales@, contact@, support@. "
        "If not clearly present / unsure: found_phone=false / found_email=false. "
        "Return ONLY JSON."
    )
    user = (
        f"Company: {company}\nWebsite: {website or 'unknown'}\nHQ: {hq or 'unknown'}\n\n"
        f"WEBSITE TEXT (About/Contact/footer excerpts):\n{(page_text or '')[:9000]}\n\n"
        "JSON schema:\n"
        '{"found_phone":true|false,"office_phone":"","found_email":true|false,'
        '"office_email":"","confidence":"high|medium|low","reason":"brief",'
        '"phone_in_text":true|false,"email_in_text":true|false}'
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            max_tokens=350,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        )
        data = _parse_json(resp.choices[0].message.content or "{}")
    except Exception as ex:
        return {"error": str(ex)[:120]}

    conf = str(data.get("confidence") or "").lower()
    out = {"phone": "", "email": "", "confidence": conf, "raw": data}
    if conf not in {"high", "medium"}:
        return out

    phone = clean_phone(str(data.get("office_phone") or "")) if data.get("found_phone") else ""
    email = clean_email(str(data.get("office_email") or "")) if data.get("found_email") else ""

    text_l = (page_text or "").lower()
    # Prefer in-text evidence; for high confidence allow if digits/email clearly match text
    if phone:
        digits = re.sub(r"\D", "", phone)
        if digits[-8:] in re.sub(r"\D", "", page_text or "") or data.get("phone_in_text") or conf == "high":
            # still require some digit presence in page when page exists
            if page_text and digits[-7:] not in re.sub(r"\D", "", page_text):
                if conf != "high":
                    phone = ""
            out["phone"] = phone
        else:
            phone = ""
    if email:
        if not is_office_email(email):
            # allow only if explicitly in text and confidence high
            if email.lower() not in text_l or conf != "high":
                email = ""
        if email and page_text and email.lower() not in text_l and conf != "high":
            email = ""
        out["email"] = email
    return out


def main() -> None:
    client, model = ds_client()
    print(f"DeepSeek model={model}", flush=True)

    cache: dict = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[2]]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}

    gaps = []
    for row in range(3, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        if not co:
            continue
        web = str(ws.cell(row, ic["Website"]).value or "").strip()
        ph = str(ws.cell(row, ic["Office No."]).value or "").strip()
        em = str(ws.cell(row, ic["Office Email"]).value or "").strip()
        hq = str(ws.cell(row, ic["Headquarters"]).value or "").strip()
        need_p = emptyish(ph)
        need_e = emptyish(em)
        if need_p or need_e or is_junk_website(web):
            gaps.append(
                {
                    "row": row,
                    "company": co,
                    "website": web,
                    "hq": hq,
                    "need_phone": need_p,
                    "need_email": need_e,
                    "junk_web": is_junk_website(web),
                }
            )

    print(f"gap rows={len(gaps)}", flush=True)

    # Phase 1: fix junk websites via DeepSeek (serial, careful)
    web_fixed = 0
    for g in gaps:
        if not g["junk_web"]:
            continue
        key = f"webfix::{g['company']}"
        if key in cache and cache[key].get("website"):
            new_web = cache[key]["website"]
        else:
            print(f"  fix website: {g['company'][:40]}", flush=True)
            new_web = deepseek_fix_website(client, model, g["company"], g["hq"], g["website"])
            cache[key] = {"website": new_web, "ts": datetime.now(timezone.utc).isoformat()}
            time.sleep(0.2)
        if new_web:
            ws.cell(g["row"], ic["Website"]).value = new_web
            g["website"] = new_web
            g["junk_web"] = False
            web_fixed += 1
        else:
            ws.cell(g["row"], ic["Website"]).value = NPD if emptyish(g["website"]) or is_junk_website(g["website"]) else g["website"]

    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"websites fixed={web_fixed}", flush=True)

    # Phase 2: scrape all gap websites concurrently
    scrape_map: dict[str, dict] = {}
    todo = []
    for g in gaps:
        web = normalize_website(g["website"])
        if web and not is_junk_website(web):
            todo.append((g["company"], web))

    print(f"scraping {len(todo)} sites ...", flush=True)
    with ThreadPoolExecutor(max_workers=12) as ex:
        futs = {ex.submit(scrape_site, web): (co, web) for co, web in todo}
        done = 0
        for fut in as_completed(futs):
            co, web = futs[fut]
            try:
                scrape_map[web] = fut.result()
            except Exception as exn:
                scrape_map[web] = {"phone": "", "email": "", "text": "", "url": web, "ok": False, "error": str(exn)[:80]}
            done += 1
            if done % 40 == 0 or done == len(todo):
                print(f"  scraped {done}/{len(todo)}", flush=True)

    # Phase 3: DeepSeek extract for remaining gaps (use scrape first, DS to parse/confirm)
    filled_phone = filled_email = 0
    ds_calls = 0
    for i, g in enumerate(gaps, 1):
        row = g["row"]
        co = g["company"]
        web = normalize_website(g["website"])
        cur_p = str(ws.cell(row, ic["Office No."]).value or "").strip()
        cur_e = str(ws.cell(row, ic["Office Email"]).value or "").strip()
        need_p = emptyish(cur_p)
        need_e = emptyish(cur_e)
        if not need_p and not need_e:
            continue

        scrap = scrape_map.get(web, {})
        phone = scrap.get("phone") or ""
        email = scrap.get("email") or ""

        cache_key = f"ds::{co}::{web}"
        ds_hit = cache.get(cache_key) or {}
        # Prefer scrape first (regex from live About/Contact pages)
        if need_p and scrap.get("phone"):
            phone = scrap["phone"]
        if need_e and scrap.get("email"):
            email = scrap["email"]

        # Re-call DeepSeek if cache empty / failed and we still need fields
        force = os.getenv("DEEPSEEK_FORCE_REFRESH", "").strip() in {"1", "true", "yes"}
        cached_empty = not (ds_hit.get("phone") or ds_hit.get("email"))
        still_need = (need_p and not phone) or (need_e and not email)
        if still_need and (force or cached_empty or not ds_hit):
            if scrap.get("text") or (need_p and need_e):
                print(f"[{i}/{len(gaps)}] DeepSeek {co[:42]}", flush=True)
                ds_hit = deepseek_extract_office(
                    client, model, co, web, g["hq"], scrap.get("text") or ""
                )
                cache[cache_key] = ds_hit
                ds_calls += 1
                if ds_calls % 20 == 0:
                    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
                time.sleep(0.12)

        if need_p and not phone and ds_hit.get("phone"):
            phone = ds_hit["phone"]
        if need_e and not email and ds_hit.get("email") and is_office_email(ds_hit["email"]):
            email = ds_hit["email"]

        if need_p and phone:
            ws.cell(row, ic["Office No."]).value = phone
            filled_phone += 1
        elif need_p:
            ws.cell(row, ic["Office No."]).value = NPD

        if need_e and email and is_office_email(email):
            ws.cell(row, ic["Office Email"]).value = email
            filled_email += 1
        elif need_e:
            ws.cell(row, ic["Office Email"]).value = NPD

    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    wb.save(XLSX)
    wb.close()

    # recount
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c or "") for c in rows[1]]
    ic2 = {h: i for i, h in enumerate(hdr) if h}
    phone_n = sum(
        1
        for r in rows[2:]
        if r[ic2["Office No."]] and str(r[ic2["Office No."]]).strip() != NPD
    )
    email_n = sum(
        1
        for r in rows[2:]
        if r[ic2["Office Email"]] and str(r[ic2["Office Email"]]).strip() != NPD
    )
    wb.close()

    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "file": str(XLSX),
        "gap_rows": len(gaps),
        "websites_fixed": web_fixed,
        "deepseek_calls": ds_calls,
        "newly_filled_phone": filled_phone,
        "newly_filled_email": filled_email,
        "total_phone_filled": phone_n,
        "total_email_filled": email_n,
        "total_rows": len(rows) - 2,
        "model": model,
        "rule": "DeepSeek extract from scraped About/Contact text; high/medium confidence only; no invented contacts",
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(audit, indent=2))


if __name__ == "__main__":
    main()
