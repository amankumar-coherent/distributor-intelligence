#!/usr/bin/env python3
"""Add + fill Operational Presence (and refine Continent / Geography) for all 6 markets.

DEFINITIONS
-----------
Continent / Geography
  Countries / markets where the company SELLS, DISTRIBUTES, or SERVES customers
  (commercial reach / markets served). Export-only markets belong here.

Operational Presence
  Countries where the company has PHYSICAL operations: headquarters, offices,
  warehouses, plants, retail stores, subsidiaries, distribution centres.
  Export-only markets without a physical footprint do NOT belong here.
  Always include the HQ country when known.

Column order: … Headquarters | Continent / Geography | Operational Presence | …

Sources: Google AI scraper → DDGS → DeepSeek (OPENAI_* → api.deepseek.com).
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")
os.environ["GOOGLE_AI_SCRAPER_ENABLED"] = "true"
os.environ["GOOGLE_AI_SCRAPER_FORCE_OVERVIEW"] = "false"
os.environ["GOOGLE_AI_SCRAPER_OVERVIEW_RETRIES"] = "1"

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping  # noqa: E402

FILES = [
    ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market.xlsx",
    ROOT / "output/web_expand/global_smartphone_market_global/global_smartphone_market.xlsx",
    ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global/global_antibody_drug_conjugates_market.xlsx",
    ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market.xlsx",
    ROOT / "output/run4/global_green_chemical_market_global/global_green_chemical_market.xlsx",
    ROOT / "output/run4/global_firewall_market_global/global_firewall_market.xlsx",
]

GEO_COL = "Continent / Geography"
OP_COL = "Operational Presence"
PROGRESS = ROOT / "scripts/_op_presence_progress.jsonl"
SUMMARY = ROOT / "scripts/_op_presence_summary.json"
BATCH = 15
CONCURRENCY = 5
# Prefer DDGS+DeepSeek for speed; Google AI for harder multi-country residuals
USE_GOOGLE_FIRST = False

HQ_PATTERNS = [
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
    (r"\bbrazil\b|\bbrasil\b", "Brazil"),
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
    (r"\bsri lanka\b", "Sri Lanka"),
    (r"\bbermuda\b", "Bermuda"),
    (r"\bsyria\b", "Syria"),
]

US_HINTS = (
    "california", "new york", "texas", "florida", "illinois", "washington",
    "massachusetts", "pennsylvania", "ohio", "colorado", "arizona", "nevada",
    "virginia", "michigan", "minnesota", "georgia", "new jersey", "seattle",
    "san francisco", "los angeles", "chicago", "boston", "miami", "irvine",
    "melville", "alpharetta", "bentonville", "thief river falls",
)


def norm(s: object) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def header_row(ws) -> int:
    for r in range(1, min(8, ws.max_row + 1)):
        vals = [norm(c.value) for c in ws[r] if c.value not in (None, "")]
        if "company" in vals:
            return r
    return 1


def hq_country(hq: str) -> str | None:
    low = norm(hq)
    if not low or low in {"npd", "n/a", "unknown", "global", "not publicly disclosed"}:
        return None
    for pat, name in HQ_PATTERNS:
        if re.search(pat, low):
            return name
    if any(h in low for h in US_HINTS):
        return "United States"
    if "lyon" in low or "auvergne" in low or "rhône" in low or "rhone" in low:
        return "France"
    if "johannesburg" in low or "gauteng" in low:
        return "South Africa"
    if "holzmaden" in low:
        return "Germany"
    if "hamilton" in low and "bermuda" in low:
        return "Bermuda"
    return None


def normalize_countries(vals: object) -> str:
    if vals is None:
        return ""
    if isinstance(vals, str):
        s = vals.strip()
        if not s or norm(s) in {"global", "worldwide", "n/a", "npd", "unknown"}:
            return ""
        if s.startswith("["):
            try:
                vals = json.loads(s.replace("'", '"'))
            except Exception:
                vals = re.split(r"[,;/|]", s.strip("[]"))
        else:
            vals = re.split(r"[,;/|]", s)
    if not isinstance(vals, list):
        return ""
    seen: set[str] = set()
    out: list[str] = []
    skip = {
        "global", "worldwide", "world", "international", "europe", "asia",
        "africa", "middle east", "north america", "latin america", "apac",
        "emea", "latam", "n/a", "npd", "unknown", "none",
    }
    for it in vals:
        c = str(it).strip().strip("'\"")
        if not c:
            continue
        # map aliases via hq_country helper on the token
        mapped = hq_country(c) or (c.title() if c.islower() else c)
        k = norm(mapped)
        if k in skip or k in seen:
            continue
        # reject city-only short junk without country pattern
        if len(mapped) < 4:
            continue
        seen.add(k)
        out.append(mapped)
    return "; ".join(out)


def ensure_op_column(path: Path) -> None:
    wb = load_workbook(path)
    ws = wb.active
    hr = header_row(ws)
    headers = [str(c.value or "").strip() for c in ws[hr]]
    if OP_COL in headers:
        # already present — ensure position after GEO
        gi = headers.index(GEO_COL) if GEO_COL in headers else None
        oi = headers.index(OP_COL)
        if gi is not None and oi != gi + 1:
            # move: read values, delete, insert after geo
            vals = [ws.cell(r, oi + 1).value for r in range(1, ws.max_row + 1)]
            ws.delete_cols(oi + 1)
            headers = [str(c.value or "").strip() for c in ws[hr]]
            gi = headers.index(GEO_COL)
            ws.insert_cols(gi + 2)
            for r, v in enumerate(vals, start=1):
                ws.cell(r, gi + 2).value = v if r != hr else OP_COL
        wb.save(path)
        return

    if GEO_COL not in headers:
        print(f"  WARN no {GEO_COL} in {path.name}")
        wb.close()
        return
    gi = headers.index(GEO_COL)  # 0-based
    ws.insert_cols(gi + 2)  # after geo
    ws.cell(hr, gi + 2).value = OP_COL
    # style like neighbors
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    wrap = Alignment(wrap_text=True, vertical="top", horizontal="left")
    font = Font(name="Calibri", size=11)
    hfont = Font(name="Calibri", size=11, bold=True)
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(r, gi + 2)
        cell.border = thin
        cell.alignment = wrap
        cell.font = hfont if r == hr else font
        cell.fill = PatternFill(fill_type=None)
    from openpyxl.utils import get_column_letter
    ws.column_dimensions[get_column_letter(gi + 2)].width = 36
    wb.save(path)
    print(f"  inserted {OP_COL} in {path.name}")


def collect_rows() -> list[dict]:
    rows: list[dict] = []
    for f in FILES:
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        hr = header_row(ws)
        headers = [str(c.value or "").strip() for c in ws[hr]]
        ci = headers.index("Company")
        hi = headers.index("Headquarters") if "Headquarters" in headers else None
        wi = headers.index("Website") if "Website" in headers else None
        gi = headers.index(GEO_COL) if GEO_COL in headers else None
        oi = headers.index(OP_COL) if OP_COL in headers else None
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row or not row[ci]:
                continue
            company = str(row[ci]).strip()
            rows.append(
                {
                    "file": str(f),
                    "company": company,
                    "hq": str(row[hi] or "").strip() if hi is not None else "",
                    "website": str(row[wi] or "").strip() if wi is not None else "",
                    "old_geo": str(row[gi] or "").strip() if gi is not None else "",
                    "old_op": str(row[oi] or "").strip() if oi is not None else "",
                    "key": f"{norm(f.name)}::{norm(company)}",
                }
            )
        wb.close()
    return rows


def load_done() -> dict[str, dict]:
    done: dict[str, dict] = {}
    if PROGRESS.exists():
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                obj = json.loads(line)
                if obj.get("key") and (obj.get("geo") or obj.get("op")):
                    done[obj["key"]] = obj
            except Exception:
                pass
    return done


def deepseek() -> OpenAI:
    return OpenAI(
        api_key=os.getenv("OPENAI_API_KEY") or "",
        base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/"),
    )


def ddg_text(company: str, hq: str = "", website: str = "") -> str:
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS  # type: ignore
        except ImportError:
            return ""
    queries = [
        f"{company} headquarters offices warehouses plants countries operations",
        f"{company} markets served countries distribution",
    ]
    if hq:
        queries[0] += f" {hq}"
    bits: list[str] = []
    try:
        with DDGS() as ddgs:
            for q in queries:
                for r in ddgs.text(q, max_results=4):
                    bits.append(
                        f"- {r.get('title','')}: {r.get('body') or r.get('description') or ''} | {r.get('href') or ''}"
                    )
    except Exception as e:
        bits.append(f"(ddg error: {e})")
    return "\n".join(bits)


def extract_both(company: str, hq: str, website: str, old_geo: str, text: str, source: str) -> dict:
    if not (text or "").strip():
        return {}
    client = deepseek()
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    prompt = f"""Distinguish TWO fields for this company. Use ONLY evidence in the text + known HQ.

Company: {company}
Known HQ text: {hq or 'unknown'}
Website: {website or 'unknown'}
Existing markets-served hint: {old_geo or 'none'}

Text ({source}):
```
{text[:12000]}
```

DEFINITIONS (strict):
1) markets_served (Continent / Geography): countries where the company SELLS / DISTRIBUTES / SERVES customers (commercial reach). Include export markets.
2) operational_presence: countries with PHYSICAL footprint only — HQ, offices, warehouses, plants, retail stores, subsidiaries, DCs. Do NOT list export-only markets here.

Return JSON only:
{{
  "markets_served": ["Country1","Country2"],
  "operational_presence": ["Country1","Country2"],
  "confidence": "high"|"medium"|"low"
}}
Rules:
- Always include HQ country in operational_presence when HQ is known.
- Prefer concrete country names (never Global/Europe alone).
- If operational countries unknown beyond HQ, return only HQ country for operational_presence.
- If markets unknown, reuse markets_served from hint if it looks like country list, else HQ country.
- Do not invent.
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )
    raw = (resp.choices[0].message.content or "").strip()
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.I | re.M).strip()
    try:
        data = json.loads(raw)
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}
    geo = normalize_countries(data.get("markets_served"))
    op = normalize_countries(data.get("operational_presence"))
    conf = str(data.get("confidence") or "low").lower()
    hc = hq_country(hq)
    if hc:
        # ensure HQ in OP
        op_list = [x.strip() for x in op.split(";") if x.strip()] if op else []
        if hc not in op_list and norm(hc) not in {norm(x) for x in op_list}:
            op_list = [hc] + op_list
        op = "; ".join(op_list)
        if not geo:
            geo = normalize_countries(old_geo) or hc
    return {"geo": geo, "op": op, "confidence": conf, "source": source}


async def google_text(company: str, hq: str) -> str:
    q = f"{company} headquarters offices warehouses plants countries served markets operations"
    if hq:
        q = f"{company} {hq} offices warehouses countries served markets"
    try:
        data = await google_ai_ask(q, llm_clean=True)
        return str(data.get("markdown") or data.get("markdown_raw") or "").strip()
    except Exception as e:
        return f"(google error: {e})"


async def process_one(item: dict, sem: asyncio.Semaphore, google_ok: bool) -> dict:
    async with sem:
        company = item["company"]
        print(f"→ {company}", flush=True)
        gtext = ""
        parsed: dict = {}

        # Primary: DDGS + DeepSeek (fast, web-based)
        dtext = await asyncio.to_thread(
            ddg_text, company, item.get("hq", ""), item.get("website", "")
        )
        parsed = await asyncio.to_thread(
            extract_both,
            company,
            item.get("hq", ""),
            item.get("website", ""),
            item.get("old_geo", ""),
            dtext,
            "ddg+deepseek",
        )

        # Optional Google pass for low-confidence / empty OP
        need_google = google_ok and (
            USE_GOOGLE_FIRST
            or not parsed.get("op")
            or parsed.get("confidence") == "low"
        )
        if need_google:
            gtext = await google_text(company, item.get("hq", ""))
            bundle = f"GOOGLE:\n{gtext}\n\nDDG:\n{dtext}"
            parsed2 = await asyncio.to_thread(
                extract_both,
                company,
                item.get("hq", ""),
                item.get("website", ""),
                item.get("old_geo", ""),
                bundle,
                "google+ddg+deepseek",
            )
            if parsed2.get("op"):
                if (
                    not parsed.get("op")
                    or parsed2.get("confidence") in {"high", "medium"}
                    or len(parsed2.get("op", "")) > len(parsed.get("op", ""))
                ):
                    parsed = parsed2

        hc = hq_country(item.get("hq", ""))
        geo = normalize_countries(parsed.get("geo") or item.get("old_geo") or "") or (hc or "")
        op = normalize_countries(parsed.get("op") or "") or (hc or "")
        if not op and hc:
            op = hc
        if not geo and hc:
            geo = hc
        # Low-conf: do not copy a huge markets-served list into OP
        if (
            parsed.get("confidence") == "low"
            and op
            and geo
            and op == geo
            and op.count(";") >= 4
            and hc
        ):
            op = hc

        return {
            "key": item["key"],
            "file": item["file"],
            "company": company,
            "geo": geo,
            "op": op,
            "confidence": parsed.get("confidence") or "fallback",
            "source": parsed.get("source") or "hq_fallback",
        }


def apply(results: list[dict]) -> None:
    by_file: dict[str, dict[str, dict]] = {}
    for r in results:
        if not r.get("op") and not r.get("geo"):
            continue
        by_file.setdefault(r["file"], {})[norm(r["company"])] = r

    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    wrap = Alignment(wrap_text=True, vertical="top", horizontal="left")
    font = Font(name="Calibri", size=11)

    for fpath, mapping in by_file.items():
        wb = load_workbook(fpath)
        ws = wb.active
        hr = header_row(ws)
        headers = [str(c.value or "").strip() for c in ws[hr]]
        if OP_COL not in headers or GEO_COL not in headers:
            print("skip missing cols", fpath)
            continue
        ci = headers.index("Company")
        gi = headers.index(GEO_COL)
        oi = headers.index(OP_COL)
        n = 0
        for r in range(hr + 1, ws.max_row + 1):
            company = ws.cell(r, ci + 1).value
            if not company:
                continue
            rec = mapping.get(norm(str(company)))
            if not rec:
                continue
            if rec.get("geo"):
                cell = ws.cell(r, gi + 1)
                cell.value = rec["geo"]
                cell.alignment = wrap
                cell.font = font
                cell.border = thin
            if rec.get("op"):
                cell = ws.cell(r, oi + 1)
                cell.value = rec["op"]
                cell.alignment = wrap
                cell.font = font
                cell.border = thin
            n += 1
        wb.save(fpath)
        print(f"Updated {n} rows -> {Path(fpath).name}")


async def main() -> None:
    print("=== Ensuring Operational Presence column ===")
    for f in FILES:
        ensure_op_column(f)

    rows = collect_rows()
    print(f"Total companies: {len(rows)}")
    done = load_done()
    print(f"Already in progress: {len(done)}")

    def needs_enrich(r: dict) -> bool:
        """Enrich when geo is multi-country, OP missing/NPD, or OP==huge geo."""
        if r["key"] in done:
            return False
        geo = r.get("old_geo") or ""
        op = r.get("old_op") or ""
        if not op or op.upper() == "NPD":
            return True
        if geo.count(";") >= 1:  # multi-country markets — verify physical OP
            return True
        if op.count(";") >= 4:
            return True
        return False

    pending = [r for r in rows if needs_enrich(r)]
    # Also include remaining not-done single-country for light pass? Skip — HQ seed OK.
    print(f"Pending web enrich (multi-country / missing OP): {len(pending)}")

    google_ok = await google_ai_ping()
    print(f"Google AI scraper: {'connected' if google_ok else 'DOWN'}")

    sem = asyncio.Semaphore(CONCURRENCY)
    new_results: list[dict] = []

    with PROGRESS.open("a", encoding="utf-8") as prog:
        for i in range(0, len(pending), BATCH):
            batch = pending[i : i + BATCH]
            print(
                f"\n=== Batch {i//BATCH+1}/{(len(pending)+BATCH-1)//BATCH} ({len(batch)}) ===",
                flush=True,
            )
            recs = await asyncio.gather(*[process_one(x, sem, google_ok) for x in batch])
            for rec in recs:
                prog.write(json.dumps(rec, ensure_ascii=False) + "\n")
                prog.flush()
                new_results.append(rec)
                print(
                    f"  OP={rec.get('op','')[:60]} | GEO={rec.get('geo','')[:60]} | {rec.get('source')}",
                    flush=True,
                )

    # merge all done + new for apply
    all_map = dict(done)
    for r in new_results:
        all_map[r["key"]] = r
    # attach file/company from rows
    key_to_row = {r["key"]: r for r in rows}
    apply_list = []
    for key, rec in all_map.items():
        base = key_to_row.get(key)
        if not base:
            continue
        apply_list.append(
            {
                "key": key,
                "file": base["file"],
                "company": base["company"],
                "geo": rec.get("geo") or "",
                "op": rec.get("op") or "",
            }
        )

    print("\nApplying to Excel...")
    apply(apply_list)

    summary = {
        "total_companies": len(rows),
        "filled_this_run": len(new_results),
        "filled_total": len(apply_list),
        "with_op": sum(1 for r in apply_list if r.get("op")),
        "with_geo": sum(1 for r in apply_list if r.get("geo")),
    }
    SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print("\n=== SUMMARY ===")
    for k, v in summary.items():
        print(f"{k}: {v}")


if __name__ == "__main__":
    asyncio.run(main())
