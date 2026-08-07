#!/usr/bin/env python3
"""Fill ADC Contact Person / Role / LinkedIn via Wikipedia + company websites.

No SearXNG, no DuckDuckGo, no Apollo wait. Uses:
  - Wikipedia REST / opensearch + page extract
  - Company about/leadership pages (httpx)
  - Built-in known leadership map

Keeps existing Apollo partial names and verified full names.
"""
from __future__ import annotations

import asyncio
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote, urljoin

import httpx

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import fill_adc_contact_role_linkedin as xls

OUT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT / "global_antibody_drug_conjugates_market_FINAL.xlsx"
PROGRESS = OUT / "adc_wiki_web_contact_progress.jsonl"
SUMMARY = OUT / "adc_wiki_web_contact_summary.json"

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}

PERSON_LI_RE = re.compile(
    r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/in/[a-zA-Z0-9\-_%]+/?",
    re.I,
)

CEO_PATS = [
    re.compile(
        r"(?:CEO|Chief Executive Officer|Chief Executive|President(?:\s+and\s+CEO)?|"
        r"Managing Director|Group CEO|Vorstandsvorsitzender|"
        r"President and Chief Executive Officer|Chairman(?:\s+and\s+CEO)?)\s*[:\-–]?\s*"
        r"(Dr\.?\s+)?([A-Z][a-zA-Z.''\-]+(?:\s+[A-Z][a-zA-Z.''\-]+){1,3})",
        re.I,
    ),
    re.compile(
        r"(Dr\.?\s+)?([A-Z][a-zA-Z.''\-]+(?:\s+[A-Z][a-zA-Z.''\-]+){1,3})\s*,?\s*"
        r"(?:is|was|serves as|has been)?\s*(?:the\s+)?"
        r"(?:CEO|Chief Executive Officer|President(?:\s+and\s+CEO)?|Managing Director|Group CEO)\b",
        re.I,
    ),
    re.compile(
        r"\|\s*[Cc]hief [Ee]xecutive [Oo]fficer\s*\|\s*(Dr\.?\s+)?([A-Z][^\|\n]{2,40})",
    ),
    re.compile(
        r"\|\s*[Kk]ey people\s*\|\s*.*?\b([A-Z][a-zA-Z.''\-]+(?:\s+[A-Z][a-zA-Z.''\-]+)+)\s*\((?:CEO|Chief)",
        re.I,
    ),
]

# Extra known contacts from public web / parent companies (Wikipedia-backed where noted)
EXTRA_KNOWN: dict[str, tuple[str, str, str]] = {
    "Alliance Healthcare Greece": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Ireland": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Polska": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare España": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare France": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Italia": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Nederland": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Norge": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Portugal": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare România": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "BriovaRx Specialty Pharmacy": (
        "Stephen Hemsley",
        "CEO (UnitedHealth Group / OptumRx)",
        "",
    ),
    "Diplomat Specialty Pharmacy (Optum)": (
        "Stephen Hemsley",
        "CEO (UnitedHealth Group / Optum)",
        "",
    ),
    "Bellco Drug Corp": (
        "Jason Hollar",
        "CEO (Cardinal Health)",
        "https://www.linkedin.com/in/jasonmhollar",
    ),
    "HD Smith / Cencora affiliate network": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "J.M. Blanco": (
        "Lisa Thompson Harrison",
        "Cencora affiliate leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "United Drug / McKesson Ireland": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "Symbion": (
        "Vikesh Ramsunder",
        "CEO (Sigma Healthcare)",
        "",
    ),
}


def empty(v: object) -> bool:
    return str(v or "").strip().lower() in NPD


def needs_fill(row: dict) -> bool:
    cp = str(row.get("Contact Person") or "").strip()
    if empty(cp):
        return True
    low = cp.lower()
    if "leadership" in low or "joint md" in low:
        return True
    return False


def looks_person(name: str, company: str = "") -> bool:
    n = (name or "").strip()
    if len(n) < 4 or len(n) > 60 or "\n" in n:
        return False
    low = n.lower()
    if empty(n) or "not publicly" in low:
        return False
    # allow Apollo partials
    if "*" in n:
        parts = n.replace("*", "x").replace("-", " ").split()
        return 2 <= len(parts) <= 5
    bad = {
        "ceo", "inc", "ltd", "llc", "gmbh", "team", "the", "and", "leadership",
        "board", "group", "company", "president", "director", "officer", "key",
        "people", "wikipedia", "retrieved", "edit", "references",
    }
    if any(x in low for x in ("leadership", "curriculum", "retrieved", "wikipedia")):
        return False
    parts = n.replace("-", " ").split()
    if n.startswith("Dr") and len(parts) >= 2:
        parts = [p for p in parts if p.lower() not in {"dr", "dr."}]
    if not (2 <= len(parts) <= 4):
        return False
    if any(p.lower().strip(".,") in bad for p in parts):
        return False
    caps = sum(1 for p in parts if p[:1].isupper())
    if caps < 2:
        return False
    if company and n.lower() == company.strip().lower():
        return False
    return True


def strip_html(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"&nbsp;|&amp;|&quot;|&#39;|&#160;", " ", t)
    return re.sub(r"\s+", " ", t)


def clean_name(raw: str) -> str:
    n = re.sub(r"\s+", " ", (raw or "")).strip(" .,;|[]()")
    n = re.sub(r"\s*\(.*?\)\s*", " ", n).strip()
    # drop trailing junk
    n = re.split(r"\s+is\s+|\s+was\s+", n, maxsplit=1)[0].strip()
    return n[:60]


def extract_from_blob(blob: str, company: str) -> dict[str, str]:
    out: dict[str, str] = {}
    if not blob:
        return out
    for m in PERSON_LI_RE.finditer(blob):
        out["linkedin"] = m.group(0).rstrip("/")
        break
    for pat in CEO_PATS:
        m = pat.search(blob)
        if not m:
            continue
        groups = [g for g in m.groups() if g]
        name = clean_name(groups[-1] if groups else "")
        if looks_person(name, company):
            out["person"] = name
            # role guess
            ctx = blob[max(0, m.start() - 30) : m.end() + 40]
            if re.search(r"Managing Director", ctx, re.I):
                out["role"] = "Managing Director"
            elif re.search(r"President and CEO|Chairman and CEO", ctx, re.I):
                out["role"] = "President and CEO"
            elif re.search(r"Group CEO", ctx, re.I):
                out["role"] = "Group CEO"
            elif re.search(r"Chairman", ctx, re.I):
                out["role"] = "Chairman"
            elif re.search(r"Founder", ctx, re.I):
                out["role"] = "Founder"
            else:
                out["role"] = "CEO"
            break
    return out


async def fetch(client: httpx.AsyncClient, url: str, timeout: float = 8.0) -> str:
    try:
        r = await client.get(url, timeout=timeout, follow_redirects=True)
        if r.status_code >= 400:
            return ""
        return r.text or ""
    except Exception:
        return ""


async def wiki_text(client: httpx.AsyncClient, company: str) -> str:
    # opensearch → extract
    q = company.split("/")[0].split("(")[0].strip()
    # strip trailing descriptors
    for cut in (" specialty", " distribution", " wholesale", " network", " affiliate"):
        if cut in q.lower():
            q = q[: q.lower().index(cut)].strip()
    if len(q) < 3:
        return ""
    try:
        search_url = (
            "https://en.wikipedia.org/w/api.php?"
            f"action=opensearch&search={quote(q)}&limit=3&namespace=0&format=json"
        )
        raw = await fetch(client, search_url, timeout=10.0)
        if not raw:
            return ""
        data = json.loads(raw)
        titles = data[1] if isinstance(data, list) and len(data) > 1 else []
        blobs = []
        for title in titles[:2]:
            # skip clearly unrelated
            tlow = title.lower()
            qtok = [t for t in re.split(r"\W+", q.lower()) if len(t) > 3][:2]
            if qtok and not any(t in tlow for t in qtok):
                continue
            ext_url = (
                "https://en.wikipedia.org/w/api.php?"
                f"action=query&prop=extracts|revisions&rvprop=content&explaintext=1"
                f"&titles={quote(title)}&format=json&redirects=1"
            )
            jtxt = await fetch(client, ext_url, timeout=12.0)
            if not jtxt:
                continue
            j = json.loads(jtxt)
            pages = (j.get("query") or {}).get("pages") or {}
            for page in pages.values():
                extract = str(page.get("extract") or "")[:4000]
                revs = page.get("revisions") or []
                wikitext = ""
                if revs:
                    wikitext = str(revs[0].get("*") or revs[0].get("slots", {}).get("main", {}).get("*") or "")[:5000]
                blobs.append(extract)
                blobs.append(wikitext)
        return " ".join(blobs)[:8000]
    except Exception:
        return ""


async def company_pages(client: httpx.AsyncClient, website: str) -> str:
    if not website or not website.startswith("http"):
        return ""
    blobs = []
    base = website.rstrip("/")
    paths = (
        "",
        "/about",
        "/about-us",
        "/company",
        "/leadership",
        "/our-team",
        "/management",
        "/en/about",
        "/en/company",
        "/who-we-are",
        "/about/leadership",
        "/our-company",
    )
    for path in paths:
        url = base if not path else urljoin(base + "/", path.lstrip("/"))
        html = await fetch(client, url, timeout=7.0)
        if html and len(html) > 400:
            blobs.append(strip_html(html)[:3000])
        if len(blobs) >= 3:
            break
    return " ".join(blobs)[:7000]


async def enrich_one(client: httpx.AsyncClient, row: dict) -> dict[str, str]:
    company = row["Company"]
    known = {**xls.KNOWN, **EXTRA_KNOWN}
    if company in known:
        p, role, li = known[company]
        if not empty(p) and "not publicly" not in p.lower():
            return {"person": p, "role": role, "linkedin": li or "", "source": "known"}

    website = str(row.get("Website") or "")
    site_blob, wiki_blob = await asyncio.gather(
        company_pages(client, website),
        wiki_text(client, company),
    )
    found = extract_from_blob(site_blob + " " + wiki_blob, company)
    if found.get("person"):
        found["source"] = "wiki_web"
        return found
    return {"source": "none"}


async def main() -> None:
    bak = XLSX.with_name(XLSX.stem + "_pre_wiki_web.xlsx")
    if not bak.exists():
        bak.write_bytes(XLSX.read_bytes())
        print(f"Backup: {bak.name}", flush=True)

    rows = xls.load_rows(XLSX)
    todo = [r for r in rows if needs_fill(r)]
    print(f"Wiki+web contact fill: {len(todo)} need / {len(rows)} total", flush=True)

    filled = 0
    headers = {
        "User-Agent": "VendorIntelContactBot/1.0 (research; wikipedia+site)",
        "Accept-Language": "en-US,en;q=0.9",
    }
    async with httpx.AsyncClient(headers=headers) as client:
        for i in range(0, len(todo), 8):
            chunk = todo[i : i + 8]
            results = await asyncio.gather(*[enrich_one(client, r) for r in chunk])
            for r, hit in zip(chunk, results):
                person = str(hit.get("person") or "").strip()
                role = str(hit.get("role") or "").strip()
                li = str(hit.get("linkedin") or "").strip()
                if person and looks_person(person, r["Company"]):
                    r["Contact Person"] = person
                    r["Role"] = role or "CEO"
                    if li and "/in/" in li:
                        r["LinkedIn"] = li
                    elif empty(r.get("LinkedIn")) or "/company/" in str(r.get("LinkedIn") or ""):
                        r["LinkedIn"] = "Not publicly disclosed"
                    filled += 1
                    print(
                        f"  [{i+1}+] {r['Company'][:40]} -> {person} | {r['Role'][:40]} "
                        f"({hit.get('source')})",
                        flush=True,
                    )
                    with PROGRESS.open("a", encoding="utf-8") as f:
                        f.write(
                            json.dumps(
                                {
                                    "company": r["Company"],
                                    "person": person,
                                    "role": r["Role"],
                                    "linkedin": r.get("LinkedIn"),
                                    "source": hit.get("source"),
                                },
                                ensure_ascii=False,
                            )
                            + "\n"
                        )
                else:
                    print(f"  [{i+1}+] {r['Company'][:40]} -> no public contact", flush=True)
            # checkpoint
            rows_sorted = sorted(rows, key=lambda x: x["Company"].lower())
            xls.write_xlsx(XLSX, rows_sorted)
            print(f"  … checkpoint ({filled} filled)", flush=True)

    rows.sort(key=lambda x: x["Company"].lower())
    xls.write_xlsx(XLSX, rows)

    named = sum(1 for r in rows if not empty(r.get("Contact Person")))
    partial = sum(1 for r in rows if "*" in str(r.get("Contact Person") or ""))
    roles = sum(1 for r in rows if not empty(r.get("Role")))
    lis = sum(1 for r in rows if "/in/" in str(r.get("LinkedIn") or ""))
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total": len(rows),
        "todo": len(todo),
        "filled_this_run": filled,
        "contact_person_filled": named,
        "partial_names": partial,
        "role_filled": roles,
        "person_linkedin": lis,
        "note": "Wikipedia + company sites only; no SearXNG/DDG/Apollo wait.",
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=2), flush=True)


if __name__ == "__main__":
    asyncio.run(main())
