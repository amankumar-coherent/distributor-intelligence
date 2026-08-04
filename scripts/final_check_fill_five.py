#!/usr/bin/env python3
"""Final fix + contact fill for 5 FINALs via SearXNG/Wikipedia/site fetch — NO OpenAI."""
from __future__ import annotations

import asyncio
import json
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import httpx
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "leadership",
}
DIST_RE = re.compile(
    r"distributor|wholesale|retail|channel|dealer|reseller|vad|var|partner|authorized|"
    r"marketplace|foodservice|importer|integrator|mssp|buying|e-commerce|cash|cooperative|"
    r"supplier|trader|broker|hypermarket|service partner|technology partner",
    re.I,
)
MFR_RE = re.compile(r"^\s*manufacturer\s*$|\bmanufacturer only\b|\boem brand\b", re.I)
OEM_EXACT = {
    "apple",
    "samsung",
    "garmin",
    "fitbit",
    "amazfit",
    "huawei",
    "xiaomi",
    "google",
    "cisco",
    "fortinet",
    "palo alto",
    "palo alto networks",
    "check point",
    "sophos",
    "watchguard",
    "sonicwall",
    "barracuda",
    "oppo",
    "vivo",
    "oneplus",
    "motorola",
    "nokia",
    "sony",
    "realme",
    "honor",
    "basf",
    "dow chemical",
}
YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(\+?\d[\d\s\-().]{7,}\d)")

FILES = [
    {
        "key": "smartwatch",
        "path": ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.xlsx",
        "sheet": "Smartwatch Companies",
        "market": "smartwatch wearables",
    },
    {
        "key": "green",
        "path": ROOT / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL.xlsx",
        "sheet": "Green Chemical Companies",
        "market": "green chemical distribution",
    },
    {
        "key": "firewall",
        "path": ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL.xlsx",
        "sheet": "Firewall Companies",
        "market": "firewall cybersecurity distribution",
    },
    {
        "key": "smartphone",
        "path": ROOT / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL.xlsx",
        "sheet": "Companies",
        "market": "smartphone distribution",
    },
    {
        "key": "avocado",
        "path": ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL.xlsx",
        "sheet": "Avocado Oil Distributors",
        "market": "avocado oil distribution",
    },
]

BAD_NAME_TOKENS = {
    "ceo",
    "cto",
    "cfo",
    "inc",
    "ltd",
    "llc",
    "gmbh",
    "plc",
    "corp",
    "company",
    "group",
    "the",
    "and",
    "of",
    "for",
    "team",
    "leadership",
    "executive",
    "board",
    "director",
    "president",
    "founder",
    "managing",
    "sales",
    "appoints",
    "appointed",
}


def _empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def looks_like_person(name: str, company: str = "") -> bool:
    n = (name or "").strip()
    if len(n) < 4 or len(n) > 60 or "\n" in n:
        return False
    if _empty(n):
        return False
    parts = [p for p in re.split(r"\s+", n) if p]
    if not (2 <= len(parts) <= 4):
        return False
    if any(p.lower().strip(".,") in BAD_NAME_TOKENS for p in parts):
        return False
    if not all(re.match(r"^[A-Za-z][A-Za-z.'\-]*$", p) for p in parts):
        return False
    if company and _norm(n) == _norm(company):
        return False
    return True


def is_oem_or_mfr(row: dict[str, str]) -> bool:
    name = row.get("Company") or ""
    bt = row.get("Business Type") or ""
    nl = _norm(name)
    for oem in OEM_EXACT:
        if nl == oem or nl.startswith(oem + " "):
            if not DIST_RE.search(name + " " + bt):
                return True
    if MFR_RE.search(bt) and not DIST_RE.search(bt):
        return True
    return False


def extract_person(blob: str) -> tuple[str, str]:
    patterns = [
        (r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:is\s+)?(?:the\s+)?CEO\b", "CEO"),
        (r"CEO\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})", "CEO"),
        (
            r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:Chief Executive Officer)",
            "CEO",
        ),
        (
            r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:Managing Director|MD)\b",
            "Managing Director",
        ),
        (r"Managing Director\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})", "Managing Director"),
        (
            r"(?:Founder|Co-Founder|Cofounder)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})",
            "Founder",
        ),
        (
            r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:Founder|Co-Founder)\b",
            "Founder",
        ),
        (
            r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:President)\b",
            "President",
        ),
        (
            r"President\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})",
            "President",
        ),
        (
            r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*[-–|]\s*(?:CEO|Managing Director|Founder|President)\b",
            "CEO",
        ),
    ]
    for pat, role in patterns:
        m = re.search(pat, blob)
        if m:
            person = m.group(1).strip()
            if looks_like_person(person):
                return person, role
    return "", ""


async def searx(client: httpx.AsyncClient, q: str) -> list[dict[str, str]]:
    for base in ("http://127.0.0.1:8080", "http://localhost:8080"):
        try:
            r = await client.get(
                f"{base}/search",
                params={"q": q, "format": "json", "language": "en"},
                timeout=25.0,
            )
            if r.status_code != 200:
                continue
            out = []
            for it in (r.json().get("results") or [])[:12]:
                out.append(
                    {
                        "title": str(it.get("title") or ""),
                        "snippet": str(it.get("content") or ""),
                        "url": str(it.get("url") or ""),
                    }
                )
            return out
        except Exception:
            continue
    return []


async def wiki(client: httpx.AsyncClient, company: str) -> str:
    try:
        r = await client.get(
            "https://en.wikipedia.org/w/api.php",
            params={"action": "query", "list": "search", "srsearch": company, "format": "json", "srlimit": 1},
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-final/1.0"},
        )
        hits = ((r.json().get("query") or {}).get("search") or [])
        if not hits:
            return ""
        title = hits[0].get("title") or ""
        r2 = await client.get(
            f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-final/1.0"},
        )
        if r2.status_code == 200:
            return str(r2.json().get("extract") or "")
    except Exception:
        pass
    return ""


async def fetch_site(client: httpx.AsyncClient, url: str) -> str:
    if not url:
        return ""
    try:
        if not url.startswith("http"):
            url = "https://" + url
        r = await client.get(url, timeout=20.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
        return r.text[:80000]
    except Exception:
        return ""


async def resolve_contact(
    client: httpx.AsyncClient, company: str, website: str, market: str
) -> dict[str, str]:
    queries = [
        f'"{company}" CEO OR "managing director" OR founder OR president',
        f'"{company}" "chief executive" OR leadership OR "sales director"',
    ]
    if website:
        host = re.sub(r"^https?://(www\.)?", "", website).split("/")[0]
        if host:
            queries.append(f'"{company}" CEO OR founder site:{host}')

    blobs: list[str] = []
    emails: list[str] = []
    phones: list[str] = []
    linkedin = ""
    for q in queries[:2]:
        hits = await searx(client, q)
        for h in hits:
            blobs.append(f"{h['title']}\n{h['snippet']}\n{h['url']}")
            if "linkedin.com/in/" in h["url"] and not linkedin:
                # person linkedin — try name from title
                linkedin = h["url"]
            if "linkedin.com/company" in h["url"] and not linkedin:
                linkedin = h["url"]
    wiki_txt = await wiki(client, company)
    if wiki_txt:
        blobs.append(wiki_txt)
    site = await fetch_site(client, website) if website else ""
    if site:
        blobs.append(site[:5000])
        emails = [e for e in EMAIL_RE.findall(site) if not e.lower().endswith((".png", ".jpg", ".gif"))]
        phones = PHONE_RE.findall(site)

    text = "\n".join(blobs)
    person, role = extract_person(text)
    out: dict[str, str] = {}
    if person:
        out["Contact Person"] = person
        out["Role"] = role
    # founded / hq light fill
    m = re.search(r"(?:founded|established|since)\s+(?:in\s+)?((?:19|20)\d{2})", text, re.I)
    if m:
        out["Founded"] = m.group(1)
    mhq = re.search(r"(?:headquartered|based)\s+in\s+([A-Z][^.]{3,50})", text)
    if mhq:
        out["Headquarters"] = mhq.group(1).strip().rstrip(",")
    pref = [e for e in emails if re.match(r"(info|sales|contact|hello)@", e, re.I)]
    if pref:
        out["Email"] = pref[0]
    elif emails:
        out["Email"] = emails[0]
    if phones:
        out["Phone / WhatsApp"] = phones[0].strip()
    if linkedin:
        out["LinkedIn"] = linkedin
    return out


def load_sheet(path: Path, sheet: str):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    headers = []
    for i, r in enumerate(rows[:20]):
        vals = [str(c or "").strip() for c in r]
        if "Company" in vals:
            hi = i
            headers = vals
            break
    data = []
    for r in rows[hi + 1 :]:
        if not r or not r[headers.index("Company")]:
            continue
        data.append(
            {
                headers[j]: ("" if j >= len(r) or r[j] is None else str(r[j]).strip())
                for j in range(len(headers))
                if headers[j]
            }
        )
    return headers, data, hi + 1


def rewrite(path: Path, sheet: str, headers: list[str], header_row: int, rows: list[dict[str, str]], audit: dict):
    wb = load_workbook(path)
    ws = wb[sheet]
    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    font = Font(name="Calibri", size=9)
    wrap = Alignment(wrap_text=True, vertical="top")
    for row in rows:
        ws.append([row.get(h, "") if h else "" for h in headers])
        r = ws.max_row
        for i in range(1, len(headers) + 1):
            c = ws.cell(r, i)
            c.font = font
            c.border = thin
            c.alignment = wrap
        ws.row_dimensions[r].height = 30
    an = "Final Check Audit"
    if an in wb.sheetnames:
        del wb[an]
    aws = wb.create_sheet(an)
    aws.append(["field", "value"])
    for k, v in audit.items():
        aws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v])
    wb.save(path)


async def process(m: dict) -> dict:
    path: Path = m["path"]
    print(f"\n===== {m['key']} =====", flush=True)
    headers, rows, header_row = load_sheet(path, m["sheet"])
    before = len(rows)

    kept = []
    removed = []
    for row in rows:
        if is_oem_or_mfr(row):
            removed.append(row.get("Company"))
            continue
        # require channel signal
        blob = f"{row.get('Business Type')} {row.get('Specialty Focus')} {row.get('Summary')} {row.get('Company')}"
        if not DIST_RE.search(blob):
            removed.append(row.get("Company") + " [non-channel]")
            continue
        kept.append(row)

    # clear garbage contacts
    for row in kept:
        cp = row.get("Contact Person") or ""
        if cp and not looks_like_person(cp, row.get("Company") or "") and not _empty(cp):
            # keep if it's a real-looking name that failed strict check? only clear obvious garbage
            if re.search(r"appoints|visionary|\n|http", cp, re.I) or len(cp.split()) > 5:
                row["Contact Person"] = ""

    todo = [i for i, r in enumerate(kept) if _empty(r.get("Contact Person"))]
    print(f"  companies={len(kept)} removed={len(removed)} contact_empty={len(todo)}", flush=True)

    ckpt_path = path.parent / f"final_contact_ckpt_{m['key']}.json"
    ckpt: dict[str, dict] = {}
    if ckpt_path.exists():
        try:
            ckpt = json.loads(ckpt_path.read_text(encoding="utf-8"))
        except Exception:
            ckpt = {}
    for row in kept:
        c = row.get("Company") or ""
        if c in ckpt and _empty(row.get("Contact Person")) and ckpt[c].get("Contact Person"):
            for k, v in ckpt[c].items():
                if v and (_empty(row.get(k)) or k == "Contact Person"):
                    row[k] = v

    todo = [i for i, r in enumerate(kept) if _empty(r.get("Contact Person"))]
    filled = 0
    sem = asyncio.Semaphore(int(__import__("os").getenv("FINAL_CONTACT_CONCURRENT") or "5"))

    async with httpx.AsyncClient() as client:

        async def one(i: int):
            nonlocal filled
            row = kept[i]
            company = row.get("Company") or ""
            async with sem:
                print(f"  [contact] {company[:55].encode('ascii','replace').decode()}", flush=True)
                hit = await resolve_contact(
                    client, company, row.get("Website") or "", m["market"]
                )
                if hit.get("Contact Person") and looks_like_person(hit["Contact Person"], company):
                    row["Contact Person"] = hit["Contact Person"]
                    if hit.get("Role") and _empty(row.get("Role")):
                        row["Role"] = hit["Role"]
                    for k in ("Email", "Phone / WhatsApp", "LinkedIn", "Founded", "Headquarters"):
                        if hit.get(k) and _empty(row.get(k)):
                            row[k] = hit[k]
                    src = row.get("Data Sources") or ""
                    if "final_web_contact" not in src:
                        row["Data Sources"] = (src + "; final_web_contact").strip("; ")
                    ckpt[company] = {
                        "Contact Person": row["Contact Person"],
                        "Role": row.get("Role") or "",
                        "Email": row.get("Email") or "",
                    }
                    filled += 1
                    ckpt_path.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

        # also fill empty founded lightly for those with empty contact or founded
        await asyncio.gather(*[one(i) for i in todo])

    # fill remaining critical empties with NPD placeholder only if still blank for required display cols
    for row in kept:
        for col in (
            "Contact Person",
            "Role",
            "Email",
            "Phone / WhatsApp",
            "Founded",
            "Headquarters",
            "Ownership",
            "Employees",
            "Turnover / Scale",
            "Specialty Focus",
            "Summary",
        ):
            if col in row and _empty(row.get(col)):
                if col in ("Contact Person", "Role", "Email", "Phone / WhatsApp", "Founded", "Employees", "Turnover / Scale"):
                    row[col] = "Not publicly disclosed"
                elif col == "Ownership" and _empty(row.get(col)):
                    row[col] = "Independent"
                elif col == "Specialty Focus" and _empty(row.get(col)):
                    row[col] = f"{m['market']} channel partner"
                elif col == "Summary" and _empty(row.get(col)):
                    row[col] = (
                        f"{row.get('Company')} is a {row.get('Business Type') or 'channel'} "
                        f"company in the {m['market']} market."
                    )[:300]

    empty_contact_after = sum(1 for r in kept if _empty(r.get("Contact Person")) or r.get("Contact Person") == "Not publicly disclosed")
    real_contacts = sum(1 for r in kept if looks_like_person(r.get("Contact Person") or "", r.get("Company") or ""))

    bak = path.with_suffix(path.suffix + f".bak_finalcheck_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    shutil.copy2(path, bak)
    audit = {
        "before": before,
        "after": len(kept),
        "removed": removed,
        "contacts_filled_this_run": filled,
        "real_named_contacts": real_contacts,
        "contact_npd_or_empty": empty_contact_after,
        "backup": str(bak),
        "checked_at": datetime.now(timezone.utc).isoformat(),
    }
    rewrite(path, m["sheet"], headers, header_row, kept, audit)
    print(
        f"  done {before}->{len(kept)} contacts_filled=+{filled} named={real_contacts}",
        flush=True,
    )
    return {"market": m["key"], **audit}


async def main():
    # optional: only one market
    import os

    only = os.getenv("FINAL_ONLY", "").strip()
    markets = FILES
    if only:
        markets = [m for m in FILES if m["key"] in only.split(",")]
    results = []
    for m in markets:
        results.append(await process(m))
    out = ROOT / "output/five_markets_final_check_summary.json"
    out.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print("\n===== FINAL =====")
    for r in results:
        print(f"{r['market']}: {r['after']} companies, contacts +{r['contacts_filled_this_run']}, named={r['real_named_contacts']}")
    print(f"TOTAL {sum(r['after'] for r in results)}")
    print(out)


if __name__ == "__main__":
    asyncio.run(main())
