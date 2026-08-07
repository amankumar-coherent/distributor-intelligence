#!/usr/bin/env python3
"""Replace Global / vague Continent / Geography with concrete country lists.

Priority:
1) Derive from Regions Served / HQ / Regional Extensions
2) Expand continent macros (Europe, Middle East, …) to country lists
3) Google AI + DeepSeek for residuals that still lack countries
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")
os.environ["GOOGLE_AI_SCRAPER_ENABLED"] = "true"
os.environ["GOOGLE_AI_SCRAPER_FORCE_OVERVIEW"] = "false"
os.environ["GOOGLE_AI_SCRAPER_OVERVIEW_RETRIES"] = "1"

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping  # noqa: E402

SLIM_FILES = [
    ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global/global_antibody_drug_conjugates_market_FINAL_SLIM.xlsx",
]

GEO = "Continent / Geography"
PROGRESS = ROOT / "scripts/_fix_global_geo_progress.jsonl"
SUMMARY = ROOT / "scripts/_fix_global_geo_summary.json"
BATCH = 20
CONCURRENCY = 3

# Macro → countries (used when source says Europe / Global / etc.)
MACROS: dict[str, list[str]] = {
    "europe": [
        "United Kingdom", "Germany", "France", "Italy", "Spain", "Netherlands",
        "Belgium", "Switzerland", "Austria", "Sweden", "Norway", "Denmark",
        "Finland", "Poland", "Portugal", "Ireland", "Czech Republic", "Hungary",
        "Romania", "Greece",
    ],
    "european union": [
        "Germany", "France", "Italy", "Spain", "Netherlands", "Belgium",
        "Sweden", "Denmark", "Finland", "Poland", "Portugal", "Ireland",
        "Austria", "Czech Republic", "Hungary", "Romania", "Greece",
    ],
    "north america": ["United States", "Canada", "Mexico"],
    "latin america": [
        "Mexico", "Brazil", "Argentina", "Chile", "Colombia", "Peru",
        "Ecuador", "Uruguay", "Paraguay", "Costa Rica", "Panama",
    ],
    "south america": [
        "Brazil", "Argentina", "Chile", "Colombia", "Peru", "Ecuador",
        "Uruguay", "Paraguay", "Bolivia", "Venezuela",
    ],
    "middle east": [
        "United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Bahrain",
        "Oman", "Jordan", "Lebanon", "Egypt", "Israel", "Turkey",
    ],
    "africa": [
        "South Africa", "Nigeria", "Egypt", "Kenya", "Morocco", "Ghana",
        "Tanzania", "Ethiopia", "Algeria", "Tunisia",
    ],
    "asia": [
        "China", "Japan", "South Korea", "India", "Singapore", "Hong Kong",
        "Taiwan", "Thailand", "Malaysia", "Indonesia", "Philippines",
        "Vietnam", "Pakistan",
    ],
    "asia-pacific": [
        "China", "Japan", "South Korea", "India", "Singapore", "Hong Kong",
        "Taiwan", "Thailand", "Malaysia", "Indonesia", "Philippines",
        "Vietnam", "Australia", "New Zealand",
    ],
    "apac": [
        "China", "Japan", "South Korea", "India", "Singapore", "Hong Kong",
        "Taiwan", "Thailand", "Malaysia", "Indonesia", "Philippines",
        "Vietnam", "Australia", "New Zealand",
    ],
    "oceania": ["Australia", "New Zealand"],
    "gcc": [
        "United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Bahrain", "Oman",
    ],
    "nordics": ["Sweden", "Norway", "Denmark", "Finland", "Iceland"],
    "benelux": ["Belgium", "Netherlands", "Luxembourg"],
    "dach": ["Germany", "Austria", "Switzerland"],
    "cee": [
        "Poland", "Czech Republic", "Hungary", "Romania", "Slovakia",
        "Bulgaria", "Croatia", "Slovenia", "Serbia", "Ukraine",
    ],
    "mena": [
        "United Arab Emirates", "Saudi Arabia", "Qatar", "Kuwait", "Bahrain",
        "Oman", "Jordan", "Lebanon", "Egypt", "Morocco", "Tunisia", "Algeria",
        "Israel", "Turkey",
    ],
    "emea": [
        "United Kingdom", "Germany", "France", "Italy", "Spain", "Netherlands",
        "Belgium", "Switzerland", "Sweden", "Poland", "United Arab Emirates",
        "Saudi Arabia", "South Africa", "Egypt", "Turkey", "India",
    ],
    "caribbean": [
        "Jamaica", "Trinidad and Tobago", "Bahamas", "Barbados", "Dominican Republic",
        "Puerto Rico", "Cuba",
    ],
}

# Broad major-market fallback when company is genuinely worldwide and no better list
WORLD_MAJOR = [
    "United States", "Canada", "Mexico", "Brazil", "United Kingdom", "Germany",
    "France", "Italy", "Spain", "Netherlands", "Belgium", "Poland", "Sweden",
    "Switzerland", "Turkey", "United Arab Emirates", "Saudi Arabia", "Egypt",
    "South Africa", "India", "Singapore", "Japan", "South Korea", "China",
    "Australia", "New Zealand", "Indonesia", "Malaysia", "Thailand",
    "Philippines", "Vietnam", "Argentina", "Chile", "Colombia",
]

HQ_COUNTRY_PATTERNS = [
    (r"\bunited states\b|\busa\b|\bu\.s\.a\.?\b|\bu\.s\.\b", "United States"),
    (r"\bunited kingdom\b|\bengland\b|\bscotland\b|\bwales\b|\buk\b", "United Kingdom"),
    (r"\bunited arab emirates\b|\buae\b|\bdubai\b|\babu dhabi\b", "United Arab Emirates"),
    (r"\bsouth korea\b|\bkorea\b", "South Korea"),
    (r"\bsouth africa\b", "South Africa"),
    (r"\bnew zealand\b", "New Zealand"),
    (r"\bhong kong\b", "Hong Kong"),
    (r"\bsaudi arabia\b", "Saudi Arabia"),
    (r"\bczech republic\b|\bczechia\b", "Czech Republic"),
    (r"\bgermany\b", "Germany"),
    (r"\bfrance\b", "France"),
    (r"\bitaly\b", "Italy"),
    (r"\bspain\b", "Spain"),
    (r"\bnetherlands\b|\bholland\b", "Netherlands"),
    (r"\bbelgium\b", "Belgium"),
    (r"\bswitzerland\b", "Switzerland"),
    (r"\baustria\b", "Austria"),
    (r"\bsweden\b", "Sweden"),
    (r"\bnorway\b", "Norway"),
    (r"\bdenmark\b", "Denmark"),
    (r"\bfinland\b", "Finland"),
    (r"\bpoland\b", "Poland"),
    (r"\bportugal\b", "Portugal"),
    (r"\bireland\b", "Ireland"),
    (r"\bcanada\b", "Canada"),
    (r"\bmexico\b", "Mexico"),
    (r"\bbrazil\b", "Brazil"),
    (r"\bargentina\b", "Argentina"),
    (r"\bchile\b", "Chile"),
    (r"\bcolombia\b", "Colombia"),
    (r"\bperu\b", "Peru"),
    (r"\buruguay\b", "Uruguay"),
    (r"\bparaguay\b", "Paraguay"),
    (r"\bindia\b", "India"),
    (r"\bchina\b", "China"),
    (r"\bjapan\b", "Japan"),
    (r"\bsingapore\b", "Singapore"),
    (r"\bmalaysia\b", "Malaysia"),
    (r"\bindonesia\b", "Indonesia"),
    (r"\bthailand\b", "Thailand"),
    (r"\bvietnam\b", "Vietnam"),
    (r"\bphilippines\b", "Philippines"),
    (r"\btaiwan\b", "Taiwan"),
    (r"\baustralia\b", "Australia"),
    (r"\bturkey\b|\btürkiye\b", "Turkey"),
    (r"\begypt\b", "Egypt"),
    (r"\bnigeria\b", "Nigeria"),
    (r"\bkenya\b", "Kenya"),
    (r"\bmorocco\b", "Morocco"),
    (r"\bisrael\b", "Israel"),
    (r"\bqatar\b", "Qatar"),
    (r"\bkuwait\b", "Kuwait"),
    (r"\bbahrain\b", "Bahrain"),
    (r"\boman\b", "Oman"),
    (r"\bjordan\b", "Jordan"),
    (r"\blebanon\b", "Lebanon"),
    (r"\brussia\b", "Russia"),
    (r"\bukraine\b", "Ukraine"),
    (r"\bromania\b", "Romania"),
    (r"\bhungary\b", "Hungary"),
    (r"\bgreece\b", "Greece"),
    (r"\bpakistan\b", "Pakistan"),
    (r"\bbangladesh\b", "Bangladesh"),
]


def norm(s: object) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def header_row(ws) -> int:
    for r in range(1, min(6, ws.max_row + 1)):
        vals = [norm(c.value) for c in ws[r] if c.value not in (None, "")]
        if "company" in vals:
            return r
    return 1


def normalize_countries(vals: object) -> str:
    if vals is None:
        return ""
    if isinstance(vals, str):
        s = vals.strip()
        if s.startswith("[") and s.endswith("]"):
            try:
                vals = json.loads(s.replace("'", '"'))
            except Exception:
                vals = re.split(r"[,;/|]", s.strip("[]"))
        elif ";" in s or "," in s or "/" in s or "|" in s:
            vals = re.split(r"[,;/|]", s)
        else:
            vals = [s]
    if isinstance(vals, list):
        items = [str(x).strip().strip("'\"") for x in vals if str(x).strip()]
    else:
        return ""

    seen: set[str] = set()
    out: list[str] = []
    for it in items:
        k = norm(it)
        if not k or k in {"n/a", "npd", "unknown", "none", "null"}:
            continue
        # expand macros inline
        if k in MACROS:
            for c in MACROS[k]:
                ck = norm(c)
                if ck not in seen:
                    seen.add(ck)
                    out.append(c)
            continue
        if k in {"global", "worldwide", "world", "international", "all countries", "all"}:
            continue
        if k in seen:
            continue
        seen.add(k)
        # title-case lightly if all lower
        if it.islower():
            it = it.title()
        out.append(it)
    return "; ".join(out)


def hq_country(hq: str) -> str | None:
    low = norm(hq)
    if not low or low in {"npd", "n/a", "global", "unknown"}:
        return None
    for pat, name in HQ_COUNTRY_PATTERNS:
        if re.search(pat, low):
            return name
    return None


def needs_fix(geo: str) -> bool:
    g = str(geo or "").strip()
    if not g:
        return True
    low = norm(g)
    if low in {"global", "worldwide", "world", "international"}:
        return True
    if low in MACROS:
        return True
    parts = [p.strip() for p in re.split(r"[,;/|]", g) if p.strip()]
    if not parts:
        return True
    # all parts are macros / global words
    if all(norm(p) in MACROS or norm(p) in {"global", "worldwide", "world", "international", "export markets"} for p in parts):
        return True
    # contains Global mixed with nothing useful
    if "global" in low and len(parts) <= 2:
        # e.g. "Europe; Global"
        return True
    return False


def derive_geo(regions: str, extensions: str, hq: str, current: str) -> str:
    """Build country list from existing columns without web search."""
    chunks: list[str] = []
    for src in (regions, extensions, current):
        s = str(src or "").strip()
        if not s:
            continue
        # python list string
        if s.startswith("["):
            try:
                parsed = json.loads(s.replace("'", '"'))
                if isinstance(parsed, list):
                    chunks.extend(str(x) for x in parsed)
                    continue
            except Exception:
                pass
        chunks.append(s)

    joined = "; ".join(chunks)
    geo = normalize_countries(joined)
    if geo:
        return geo

    # only macros left → expand from current/regions
    expanded: list[str] = []
    for src in (regions, current, extensions):
        for p in re.split(r"[,;/|]", str(src or "")):
            p = p.strip().strip("'\"")
            k = norm(p)
            if k in MACROS:
                expanded.extend(MACROS[k])
            elif k in {"global", "worldwide", "world", "international"}:
                expanded.extend(WORLD_MAJOR)
    geo = normalize_countries(expanded)
    if geo:
        return geo

    hc = hq_country(hq)
    if hc:
        return hc
    return ""


def collect_targets() -> list[dict]:
    rows: list[dict] = []
    for f in SLIM_FILES:
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        hr = header_row(ws)
        headers = [str(c.value or "").strip() for c in ws[hr]]
        if GEO not in headers:
            wb.close()
            continue
        ci = headers.index("Company")
        gi = headers.index(GEO)
        hi = headers.index("Headquarters") if "Headquarters" in headers else None
        wi = headers.index("Website") if "Website" in headers else None
        ri = headers.index("Regions Served") if "Regions Served" in headers else None
        ei = headers.index("Regional Extensions") if "Regional Extensions" in headers else None
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row or not row[ci]:
                continue
            geo = str(row[gi] or "").strip()
            if not needs_fix(geo):
                continue
            company = str(row[ci]).strip()
            rows.append(
                {
                    "file": str(f),
                    "company": company,
                    "hq": str(row[hi] or "").strip() if hi is not None else "",
                    "website": str(row[wi] or "").strip() if wi is not None else "",
                    "regions_served": str(row[ri] or "").strip() if ri is not None else "",
                    "regional_ext": str(row[ei] or "").strip() if ei is not None else "",
                    "old_geo": geo,
                    "key": f"{norm(f.name)}::{norm(company)}",
                }
            )
        wb.close()
    return rows


def deepseek() -> OpenAI:
    return OpenAI(
        api_key=os.getenv("OPENAI_API_KEY") or "",
        base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/"),
    )


def parse_geo_batch(companies: list[dict], markdown: str) -> dict[str, str]:
    if not (markdown or "").strip():
        return {}
    names = [c["company"] for c in companies]
    client = deepseek()
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    prompt = f"""Extract ALL countries served for each company.

Companies:
{json.dumps(names, ensure_ascii=False)}

Text:
```
{markdown[:14000]}
```

Return JSON array only:
[{{"company":"exact name","countries":["Country1","Country2"]}}]

Rules:
- List concrete country names only (never Global/Worldwide/Europe as the only answer).
- If text says Europe / Middle East / Asia-Pacific, expand to the main countries in that region.
- If clearly worldwide but no list, list major markets (US, Canada, UK, Germany, France, ...).
- If only HQ country known, return that country.
- Do not invent unrelated countries. Skip company if nothing useful.
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )
    text = (resp.choices[0].message.content or "").strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.M).strip()
    try:
        data = json.loads(text)
    except Exception:
        return {}
    if isinstance(data, dict):
        data = data.get("companies") or data.get("results") or []
    out: dict[str, str] = {}
    if not isinstance(data, list):
        return out
    for row in data:
        if not isinstance(row, dict):
            continue
        name = row.get("company")
        geo = normalize_countries(row.get("countries") or row.get("countries_served"))
        if name and geo and not needs_fix(geo):
            out[norm(name)] = geo
        elif name and geo:
            # still vague after normalize — reject
            continue
    return out


async def google_one(company: str, hq: str, website: str) -> str:
    q = f"{company} countries served markets regions headquarters all countries"
    if hq:
        q = f"{company} {hq} all countries regions served"
    try:
        data = await google_ai_ask(q, llm_clean=True)
        md = str(data.get("markdown") or data.get("markdown_raw") or "").strip()
    except Exception as e:
        md = f"(error {e})"
    found = parse_geo_batch(
        [{"company": company, "hq": hq, "website": website}],
        md,
    )
    return found.get(norm(company), "")


def ddg_text(company: str, hq: str = "") -> str:
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS  # type: ignore
        except ImportError:
            return ""
    q = f"{company} countries served markets regions distribution"
    if hq:
        q += f" {hq}"
    bits = []
    try:
        with DDGS() as ddgs:
            for r in ddgs.text(q, max_results=5):
                bits.append(
                    f"- {r.get('title','')}: {r.get('body') or r.get('description') or ''}"
                )
    except Exception as e:
        bits.append(f"(ddg error: {e})")
    return "\n".join(bits)


def deepseek_one(c: dict, extra_text: str = "") -> str:
    client = deepseek()
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    snippets = ddg_text(c["company"], c.get("hq", ""))
    prompt = f"""List ALL countries this company serves / operates in / distributes to.

Company: {c['company']}
HQ: {c.get('hq') or 'unknown'}
Website: {c.get('website') or 'unknown'}
Regions hint: {c.get('regions_served') or 'none'}
Old geo: {c.get('old_geo') or 'none'}

Google text:
{extra_text[:6000] or '(none)'}

Web snippets:
{snippets or '(none)'}

Return JSON only: {{"countries":["Country1","Country2"]}}
Rules:
- Concrete country names only (never just Global/Europe).
- Expand region labels to countries.
- If worldwide with no list, use major world markets.
- If only HQ known, return HQ country.
- Do not invent.
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )
    text = (resp.choices[0].message.content or "").strip()
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.M).strip()
    try:
        data = json.loads(text)
    except Exception:
        return ""
    geo = normalize_countries(data.get("countries") if isinstance(data, dict) else data)
    if geo and not needs_fix(geo):
        return geo
    return ""


def apply(fills: dict[str, str], rows: list[dict]) -> int:
    by_file: dict[str, dict[str, str]] = {}
    key_map = {r["key"]: r for r in rows}
    for key, geo in fills.items():
        r = key_map.get(key)
        if not r or not geo:
            continue
        by_file.setdefault(r["file"], {})[norm(r["company"])] = geo

    n_total = 0
    for fpath, mapping in by_file.items():
        wb = load_workbook(fpath)
        ws = wb.active
        hr = header_row(ws)
        headers = [str(c.value or "").strip() for c in ws[hr]]
        ci = headers.index("Company")
        gi = headers.index(GEO)
        n = 0
        for r in range(hr + 1, ws.max_row + 1):
            company = ws.cell(r, ci + 1).value
            if not company:
                continue
            geo = mapping.get(norm(str(company)))
            if not geo:
                continue
            ws.cell(r, gi + 1).value = geo
            n += 1
        wb.save(fpath)
        print(f"Updated {n} rows -> {Path(fpath).name}")
        n_total += n
    return n_total


async def search_residual(items: list[dict], google_ok: bool) -> dict[str, str]:
    fills: dict[str, str] = {}
    sem = asyncio.Semaphore(CONCURRENCY)

    async def one(c: dict) -> tuple[str, str]:
        async with sem:
            print(f"  search → {c['company']}", flush=True)
            geo = ""
            gtext = ""
            if google_ok:
                geo = await google_one(c["company"], c.get("hq", ""), c.get("website", ""))
                if not geo:
                    try:
                        data = await google_ai_ask(
                            f"{c['company']} all countries regions served markets",
                            llm_clean=True,
                        )
                        gtext = str(data.get("markdown") or data.get("markdown_raw") or "")
                    except Exception:
                        gtext = ""
            if not geo:
                geo = await asyncio.to_thread(deepseek_one, c, gtext)
            if not geo:
                # last resort: HQ or world major if old was Global
                hc = hq_country(c.get("hq", ""))
                if hc:
                    geo = hc
                elif needs_fix(c.get("old_geo", "Global")):
                    geo = "; ".join(WORLD_MAJOR)
            return c["key"], geo

    for i in range(0, len(items), BATCH):
        batch = items[i : i + BATCH]
        print(f"\n=== Search batch {i//BATCH+1}/{(len(items)+BATCH-1)//BATCH} ({len(batch)}) ===", flush=True)
        recs = await asyncio.gather(*[one(c) for c in batch])
        for key, geo in recs:
            if geo:
                fills[key] = geo
                # progress
                with PROGRESS.open("a", encoding="utf-8") as fh:
                    fh.write(json.dumps({"key": key, "geo": geo}, ensure_ascii=False) + "\n")
    return fills


async def main() -> None:
    targets = collect_targets()
    print(f"Rows needing Continent/Geography fix: {len(targets)}")

    fills: dict[str, str] = {}
    residual: list[dict] = []

    # Phase 1: local derive
    for c in targets:
        geo = derive_geo(
            c.get("regions_served", ""),
            c.get("regional_ext", ""),
            c.get("hq", ""),
            c.get("old_geo", ""),
        )
        # If derived is only HQ country but old was Global and regions said Global → prefer world major
        # unless regions had concrete countries (already handled).
        rs = norm(c.get("regions_served", ""))
        if geo and not needs_fix(geo):
            # if old Global and regions Global and we only got HQ — expand to world major
            if (
                needs_fix(c.get("old_geo", ""))
                and (not rs or rs in {"global", "worldwide", "international"})
                and geo.count(";") == 0
                and hq_country(c.get("hq", "")) == geo
            ):
                # keep HQ only for clearly local distributors; world major for "Global" labeled
                if norm(c.get("old_geo", "")) in {"global", "worldwide", "world", "international"}:
                    geo = "; ".join(WORLD_MAJOR)
            fills[c["key"]] = geo
        else:
            residual.append(c)

    print(f"Derived locally: {len(fills)}")
    print(f"Need web search: {len(residual)}")

    if PROGRESS.exists():
        PROGRESS.unlink()

    google_ok = await google_ai_ping()
    print(f"Google AI scraper: {'connected' if google_ok else 'DOWN'}")

    if residual:
        searched = await search_residual(residual, google_ok)
        fills.update(searched)

    print("\nApplying to Excel...")
    n = apply(fills, targets)

    # verify remaining
    left = collect_targets()
    summary = {
        "targets": len(targets),
        "filled": len(fills),
        "excel_rows_updated": n,
        "still_vague_after": len(left),
        "still_names": [x["company"] for x in left[:50]],
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print("\n=== SUMMARY ===")
    for k, v in summary.items():
        if k != "still_names":
            print(f"{k}: {v}")
    if left:
        print(f"still vague sample ({len(left)}):", [x['company'] for x in left[:15]])


if __name__ == "__main__":
    asyncio.run(main())
