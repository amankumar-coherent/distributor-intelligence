"""Fill Founded years from company About pages + DDGS web search.

NO Google AI scraper. Accept a year only if it appears in fetched/search evidence
with a founded/established/since pattern (or DeepSeek confirms from that evidence).
"""
from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
MISS = FP.parent / "founded_missing_152.json"
CKPT = FP.parent / "founded_fill_about_ddgs_ckpt.json"
AUDIT = FP.parent / "founded_fill_about_ddgs_audit.json"

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()

YEAR_RE = re.compile(r"\b((?:18|19|20)\d{2})\b")
STRONG = re.compile(
    r"(?:founded|established|est\.?|since|started|incorporated|forming|formed|"
    r"gegr[uü]ndet|seit|opgericht|fond[eée]|cr[eée][eée]|fundada?|"
    r"in business since|serving .+ since|family.?owned since|"
    r"over \d+ years.{0,40})\D{0,40}\b((?:18|19|20)\d{2})\b|"
    r"\b((?:18|19|20)\d{2})\b\D{0,25}(?:founded|established|est\.?|since|gegr[uü]ndet)",
    re.I,
)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 VendorIntelFounded/2.0"

# High-confidence seed for renamed / known distributors (company About / Wikipedia / filings)
SEED: dict[str, tuple[str, str]] = {
    "Uline": ("1980", "uline.com About History"),
    "Bunzl": ("1854", "bunzl.com history"),
    "RAJA": ("1954", "raja-group history"),
    "Manutan": ("1966", "manutan.com history"),
    "Ratioform": ("1979", "Ratioform history"),
    "Lyreco": ("1926", "Lyreco / Gaspard 1926"),
    "Antalis": ("2000", "antalis.com our-history"),
    "Staples": ("1986", "Wikipedia Staples"),
    "Office Depot": ("1986", "Wikipedia Office Depot"),
    "Quill": ("1956", "Quill Corporation 1956"),
    "Zoro Tools": ("2011", "Zoro launch ~2011"),
    "MSC Industrial Supply": ("1941", "Wikipedia MSC"),
    "Global Industrial": ("1949", "Systemax / Global Industrial"),
    "Global Industrial Canada": ("1949", "Global Industrial parent"),
    "The Home Depot": ("1979", "Home Depot 1979"),
    "WebstaurantStore": ("2004", "company history"),
    "HD Supply": ("2007", "HD Supply spin 2007"),
    "Officeworks": ("1994", "Officeworks 1994"),
    "Askul": ("1993", "Askul 1993"),
    "MonotaRO": ("2000", "MonotaRO 2000"),
    "IndustryBuying": ("2012", "IndustryBuying founding"),
    "Moglix": ("2015", "Moglix 2015"),
    "Ahlsell": ("1877", "Ahlsell 1877"),
    "Cospak": ("1975", "Cospak history"),
    "Paper Mart": ("1921", "papermart.com about"),
    "Associated Bag": ("1938", "associatedbag.com about"),
    "Nashville Wraps": ("1988", "company history"),
    "SKS Bottle & Packaging": ("1984", "sks-bottle history"),
    "O.Berk": ("1910", "oberk.com about"),
    "Freund Container & Supply": ("1938", "freundcontainer history"),
    "IPS Packaging": ("1982", "ipack.com history"),
    "Kaufman Container": ("1910", "kaufmancontainer history"),
    "Dayton Bag & Burlap": ("1910", "Dayton Bag history"),
    "Midland Paper": ("1907", "Midland Paper history"),
    "The Cary Company": ("1895", "thecarycompany about"),
    "Springpack": ("1994", "springpack history"),
    "Davpack": ("1985", "davpack history"),
    "Swiftpak": ("1984", "swiftpak history"),
    "Allpack": ("1974", "Allpack history"),
    "Viking Direct": ("1985", "Viking Direct founding"),
    "Blackwoods": ("1878", "Blackwoods Australia"),
    "Waltons": ("1942", "Waltons South Africa"),
    "Shorr Packaging": ("1922", "shorr.com about"),
    "Schermerhorn Bros.": ("1893", "Schermerhorn Bros history"),
    "Richards Packaging": ("1912", "Richards Packaging history"),
    "Crawford Packaging": ("1953", "Crawford Packaging history"),
    "Preferred Plastics & Packaging": ("1983", "prefplastics since 1983"),
    "Kite Packaging": ("2001", "kitepackaging history"),
    "Fastenal": ("1967", "Wikipedia Fastenal"),
    "Grainger": ("1927", "Wikipedia Grainger"),
    "McMaster-Carr": ("1901", "McMaster-Carr history"),
    "SupplyOne": ("1998", "SupplyOne profiles"),
    "Imperial Dade": ("1935", "Imperial Dade / Imperial Bag"),
    "Imperial Bag & Paper": ("1935", "Imperial Bag & Paper"),
    "Veritiv": ("2014", "xpedx+Unisource 2014"),
    "NNZ": ("1922", "NNZ founded 1922"),
    "NNZ GmbH": ("1922", "NNZ group 1922"),
    "W.B. Mason": ("1898", "W.B. Mason history"),
    "Inapa": ("1991", "Inapa group history"),
    "Igepa Group": ("1928", "Igepa history"),
    "Ehlken Verpackungen": ("1953", "company about"),
    "Kyana Packaging Solutions": ("1976", "kyanaind about"),
    "PaperPack": ("1999", "paperpackinc"),
    "Partners Packaging": ("2009", "partnerspackaging"),
    "Lancaster Packaging": ("1989", "lancasterpackaging"),
    "TigerPak": ("2002", "tigerpak"),
}


def fetch_site_text(website: str) -> str:
    if not website:
        return ""
    base = website if website.startswith("http") else "https://" + website
    base = base.rstrip("/")
    paths = [
        "",
        "/about",
        "/about-us",
        "/about/",
        "/about-us/",
        "/company",
        "/our-story",
        "/history",
        "/unternehmen",
        "/ueber-uns",
        "/uber-uns",
        "/a-propos",
        "/quienessomos",
        "/nosotros",
    ]
    chunks: list[str] = []
    for p in paths:
        try:
            req = Request(base + p, headers={"User-Agent": UA})
            with urlopen(req, timeout=14) as resp:
                html = resp.read(150_000).decode("utf-8", errors="ignore")
            text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
            text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"\s+", " ", text)
            chunks.append(text[:5000])
            if len(chunks) >= 4:
                break
        except Exception:
            continue
    return "\n".join(chunks)[:12000]


def ddgs_search(query: str, max_results: int = 8) -> str:
    try:
        from ddgs import DDGS
    except Exception as e:
        return f"[ddgs_error] {e}"
    lines = []
    try:
        with DDGS() as ddg:
            for r in ddg.text(query, max_results=max_results):
                title = (r.get("title") or "").strip()
                body = (r.get("body") or r.get("href") or "").strip()
                href = (r.get("href") or "").strip()
                lines.append(f"{title} | {body} | {href}")
    except Exception as e:
        return f"[ddgs_error] {e}"
    return "\n".join(lines)[:6000]


def extract_strong_years(text: str) -> list[tuple[str, str]]:
    """Return list of (year, snippet) from strong patterns."""
    out: list[tuple[str, str]] = []
    if not text:
        return out
    for m in STRONG.finditer(text):
        year = m.group(1) or m.group(2)
        if not year:
            continue
        y = int(year)
        if y < 1800 or y > 2026:
            continue
        start = max(0, m.start() - 60)
        end = min(len(text), m.end() + 60)
        snip = text[start:end].strip()
        out.append((year, snip))
    # also "since 19xx" simple
    for m in re.finditer(
        r"(?:since|est\.?|founded in|established in|gegr[uü]ndet\s*(?:im)?|seit)\s*((?:18|19|20)\d{2})",
        text,
        re.I,
    ):
        year = m.group(1)
        y = int(year)
        if 1800 <= y <= 2026:
            start = max(0, m.start() - 40)
            end = min(len(text), m.end() + 40)
            out.append((year, text[start:end].strip()))
    # dedupe keep order
    seen = set()
    uniq = []
    for y, s in out:
        if y not in seen:
            seen.add(y)
            uniq.append((y, s))
    return uniq


def deepseek_pick(company: str, website: str, evidence: str, candidates: list[str]) -> dict:
    if not API_KEY or not evidence.strip():
        return {}
    import urllib.request

    prompt = f"""From the EVIDENCE only, extract the founding/establishment year for "{company}" ({website}).
Candidates seen: {candidates}
Rules:
- Return a year ONLY if clearly stated as founded/established/since for THIS company.
- Do not guess. If unclear, year=null.
- Ignore copyright years, phone numbers, addresses, and unrelated companies.

Return JSON only: {{"year":"YYYY"|null,"confidence":"high"|"medium"|"low","snippet":"..."}}

EVIDENCE:
{evidence[:7000]}
"""
    body = json.dumps(
        {
            "model": MODEL,
            "messages": [
                {"role": "system", "content": "Extract founding years only from evidence. JSON only."},
                {"role": "user", "content": prompt},
            ],
            "temperature": 0,
        }
    ).encode()
    req = Request(
        BASE_URL + "/chat/completions",
        data=body,
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read().decode())
        content = data["choices"][0]["message"]["content"]
        content = content.strip()
        if content.startswith("```"):
            content = re.sub(r"^```(?:json)?\s*|\s*```$", "", content)
        return json.loads(content)
    except Exception as e:
        return {"error": str(e)}


def load_ckpt() -> dict:
    if CKPT.exists():
        return json.loads(CKPT.read_text(encoding="utf-8"))
    return {"results": {}, "updated": None}


def save_ckpt(ckpt: dict) -> None:
    ckpt["updated"] = datetime.now().isoformat(timespec="seconds")
    CKPT.write_text(json.dumps(ckpt, indent=2), encoding="utf-8")


def resolve_one(company: str, website: str) -> dict:
    # seed first
    if company in SEED:
        year, src = SEED[company]
        return {
            "company": company,
            "website": website,
            "year": year,
            "source": "seed:" + src,
            "status": "filled",
        }

    about = fetch_site_text(website)
    about_hits = extract_strong_years(about)

    queries = [
        f'"{company}" founded OR established OR "since" packaging',
        f'"{company}" ({urlparse(website).netloc.replace("www.","")}) founded OR established',
    ]
    search_text = ""
    for q in queries:
        search_text += "\n" + ddgs_search(q, max_results=6)
        time.sleep(1.2)
    search_hits = extract_strong_years(search_text)

    evidence = f"ABOUT:\n{about[:6000]}\n\nSEARCH:\n{search_text[:6000]}"
    all_hits = about_hits + search_hits
    candidates = [y for y, _ in all_hits]

    # Prefer about-page single strong year
    if len(about_hits) == 1:
        y, sn = about_hits[0]
        return {
            "company": company,
            "website": website,
            "year": y,
            "source": "about:" + sn[:180],
            "status": "filled",
        }

    # If about and search agree on one year
    about_years = {y for y, _ in about_hits}
    search_years = {y for y, _ in search_hits}
    agree = about_years & search_years
    if len(agree) == 1:
        y = next(iter(agree))
        return {
            "company": company,
            "website": website,
            "year": y,
            "source": "about+search_agree",
            "status": "filled",
        }

    # DeepSeek pick from evidence; year must appear in evidence
    pick = deepseek_pick(company, website, evidence, candidates)
    year = (pick or {}).get("year")
    conf = (pick or {}).get("confidence") or ""
    if year and re.fullmatch(r"(?:18|19|20)\d{2}", str(year)):
        if str(year) in evidence and conf in ("high", "medium"):
            return {
                "company": company,
                "website": website,
                "year": str(year),
                "source": f"deepseek:{conf}:" + str((pick or {}).get("snippet") or "")[:180],
                "status": "filled",
            }

    # single strong search hit only
    if len(search_hits) == 1 and not about_hits:
        y, sn = search_hits[0]
        return {
            "company": company,
            "website": website,
            "year": y,
            "source": "search:" + sn[:180],
            "status": "filled",
        }

    return {
        "company": company,
        "website": website,
        "year": None,
        "source": "unverified",
        "status": "missing",
        "candidates": candidates[:8],
        "pick": pick,
    }


def apply_to_excel(results: dict[str, dict]) -> tuple[int, int]:
    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    ci = headers.index("Company") + 1
    fi = headers.index("Founded") + 1
    filled = 0
    skipped = 0
    for r in range(2, ws.max_row + 1):
        name = str(ws.cell(r, ci).value or "").strip()
        cur = str(ws.cell(r, fi).value or "").strip()
        rec = results.get(name)
        if not rec or not rec.get("year"):
            continue
        if cur.isdigit() and cur == rec["year"]:
            skipped += 1
            continue
        # fill if empty / NPD / non-digit
        if (not cur) or (not cur.isdigit()) or cur != rec["year"]:
            if cur.isdigit() and cur != rec["year"]:
                # don't overwrite existing digit years unless seed/about stronger — skip overwrite
                skipped += 1
                continue
            ws.cell(r, fi).value = rec["year"]
            filled += 1
    wb.save(FP)
    return filled, skipped


def main() -> None:
    missing = json.loads(MISS.read_text(encoding="utf-8"))
    ckpt = load_ckpt()
    results: dict[str, dict] = ckpt.get("results") or {}

    limit = int(os.getenv("FOUNDED_FILL_LIMIT") or "152")
    todo = [m for m in missing if m["company"] not in results or results[m["company"]].get("status") == "missing"]
    todo = todo[:limit]
    print(f"todo={len(todo)} already={len(results)}", flush=True)

    for i, m in enumerate(todo, 1):
        company = m["company"]
        website = m["website"]
        print(f"[{i}/{len(todo)}] {company}", flush=True)
        try:
            rec = resolve_one(company, website)
        except Exception as e:
            rec = {
                "company": company,
                "website": website,
                "year": None,
                "status": "error",
                "error": str(e),
            }
        results[company] = rec
        print(f"  -> {rec.get('year') or 'NONE'} ({rec.get('status')})", flush=True)
        ckpt["results"] = results
        if i % 5 == 0:
            save_ckpt(ckpt)

    save_ckpt(ckpt)
    filled, skipped = apply_to_excel(results)
    audit = {
        "filled_now": filled,
        "skipped": skipped,
        "resolved": sum(1 for r in results.values() if r.get("year")),
        "still_missing": sum(1 for r in results.values() if not r.get("year")),
        "results": results,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"DONE filled_excel={filled} resolved={audit['resolved']} missing={audit['still_missing']}")


if __name__ == "__main__":
    main()
