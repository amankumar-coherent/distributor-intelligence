#!/usr/bin/env python3
"""Verify Website, Founded, Headquarters — SIMPLE Google search queries only.

WRONG (shows nothing on Google):
  For EACH company, give: 1) Official website URL ...

RIGHT (what we put in the Google search bar, same style as you):
  {company} official website founded year headquarters

Pipeline:
  1) Google AI scraper — one simple query per company (25 at a time concurrency)
  2) DeepSeek residual if Google returns nothing useful
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
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")
os.environ["GOOGLE_AI_SCRAPER_ENABLED"] = "true"
# Keep simple query — do NOT rewrite into long instruction prompts
os.environ["GOOGLE_AI_SCRAPER_FORCE_OVERVIEW"] = "false"
os.environ["GOOGLE_AI_SCRAPER_OVERVIEW_RETRIES"] = "1"
os.environ["GOOGLE_AI_SCRAPER_MODE"] = "pro"

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping  # noqa: E402

SLIM_FILES = [
    ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global/global_antibody_drug_conjugates_market_FINAL_SLIM.xlsx",
]

BATCH = 25
CONCURRENCY = 5
CURRENT_YEAR = 2026
PROGRESS = ROOT / "scripts/_verify_wfh_simple_progress.jsonl"
SUMMARY = ROOT / "scripts/_verify_wfh_simple_summary.json"


def norm(s: object) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def header_row(ws) -> int:
    for r in range(1, min(6, ws.max_row + 1)):
        vals = [norm(c.value) for c in ws[r] if c.value not in (None, "")]
        if "company" in vals:
            return r
    return 1


def google_search_bar_query(company: str) -> str:
    """Exact style for Google search bar — NOT an instruction prompt."""
    return f"{company} official website founded year headquarters"


def collect_rows() -> list[dict]:
    rows: list[dict] = []
    for f in SLIM_FILES:
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        hr = header_row(ws)
        headers = [str(c.value or "").strip() for c in ws[hr]]
        ci = headers.index("Company")
        wi = headers.index("Website") if "Website" in headers else None
        fi = headers.index("Founded") if "Founded" in headers else None
        hi = headers.index("Headquarters") if "Headquarters" in headers else None
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row or not row[ci]:
                continue
            company = str(row[ci]).strip()
            rows.append(
                {
                    "file": str(f),
                    "company": company,
                    "website": str(row[wi] or "").strip() if wi is not None else "",
                    "founded": str(row[fi] or "").strip() if fi is not None else "",
                    "hq": str(row[hi] or "").strip() if hi is not None else "",
                    "key": f"{norm(f.name)}::{norm(company)}",
                }
            )
        wb.close()
    return rows


def load_done() -> set[str]:
    done: set[str] = set()
    if PROGRESS.exists():
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                obj = json.loads(line)
                if obj.get("key") and obj.get("processed"):
                    done.add(obj["key"])
            except Exception:
                pass
    return done


def clean_url(u: object) -> str | None:
    s = str(u or "").strip()
    if not s or s.upper() in {"NPD", "N/A", "NA", "NONE", "UNKNOWN"}:
        return None
    if not re.match(r"^https?://", s, re.I):
        s = "https://" + s.lstrip("/")
    try:
        p = urlparse(s)
        if not p.netloc or "." not in p.netloc:
            return None
        host = p.netloc.lower()
        junk = (
            "reddit.com",
            "facebook.com",
            "archive.org",
            "genius.com",
            "zhihu.com",
            "tiktok.com",
            "google.com",
            "wikipedia.org",
        )
        if any(j in host for j in junk):
            return None
        return f"{p.scheme}://{p.netloc}".rstrip("/")
    except Exception:
        return None


def year_ok(y: object) -> int | None:
    if y is None:
        return None
    if isinstance(y, int):
        yi = y
    else:
        m = re.search(r"(18|19|20)\d{2}", str(y))
        yi = int(m.group(0)) if m else None
    if yi is None:
        return None
    if 1800 <= yi <= CURRENT_YEAR:
        return yi
    return None


def deepseek_client() -> OpenAI:
    return OpenAI(
        api_key=os.getenv("OPENAI_API_KEY") or "",
        base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/"),
    )


def extract_from_text(company: str, text: str, source: str) -> dict:
    """Use DeepSeek only as extractor from Google/DDGS text — not as Google replacement yet."""
    if not (text or "").strip():
        return {}
    client = deepseek_client()
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    prompt = f"""From this Google/search text, extract facts for ONE company.

Company: {company}

Text:
```
{text[:8000]}
```

Return JSON only:
{{
  "website": "https://..." or null,
  "founded_year": 1994 or null,
  "headquarters": "City, Country" or null,
  "confidence": "high"|"medium"|"low"
}}
Rules:
- Official company website only (not google/wiki/news unless unavoidable).
- founded_year 1800-{CURRENT_YEAR}; ignore copyright years.
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
    return {
        "website": clean_url(data.get("website")),
        "founded": year_ok(data.get("founded_year")),
        "hq": (str(data.get("headquarters") or "").strip() or None),
        "confidence": data.get("confidence") or "medium",
        "source": source,
    }


def ddg_snippets(company: str) -> str:
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS  # type: ignore
        except ImportError:
            return ""
    q = google_search_bar_query(company)
    bits: list[str] = []
    try:
        with DDGS() as ddgs:
            for r in ddgs.text(q, max_results=5):
                bits.append(
                    f"- {r.get('title','')}: {r.get('body') or r.get('description') or ''} | {r.get('href') or ''}"
                )
    except Exception as e:
        bits.append(f"(search error: {e})")
    return "\n".join(bits)


def deepseek_residual(c: dict) -> dict:
    snippets = ddg_snippets(c["company"])
    client = deepseek_client()
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    prompt = f"""Verify official website, founded year, headquarters.

Company: {c['company']}
Current website: {c.get('website') or 'empty'}
Current founded: {c.get('founded') or 'empty'}
Current HQ: {c.get('hq') or 'empty'}

Search query used: {google_search_bar_query(c['company'])}

Web snippets:
{snippets or '(none)'}

Return JSON only:
{{
  "website": "https://..." or null,
  "founded_year": 1994 or null,
  "headquarters": "City, Country" or null,
  "confidence": "high"|"medium"|"low"
}}
Do not invent.
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
    return {
        "website": clean_url(data.get("website")),
        "founded": year_ok(data.get("founded_year")),
        "hq": (str(data.get("headquarters") or "").strip() or None),
        "confidence": data.get("confidence") or "medium",
        "source": "deepseek",
    }


def decide_updates(old: dict, new: dict) -> dict:
    conf = str(new.get("confidence") or "medium").lower()
    changes: dict[str, str] = {}

    new_web = new.get("website")
    old_web = clean_url(old.get("website"))
    if new_web and conf in {"high", "medium"} and new_web != old_web:
        changes["Website"] = new_web

    new_y = new.get("founded")
    old_y = year_ok(old.get("founded"))
    if new_y is not None and conf in {"high", "medium"}:
        if old_y is None or str(old.get("founded") or "").upper() == "NPD":
            changes["Founded"] = str(new_y)
        elif new_y != old_y and (abs(new_y - old_y) <= 80 or conf == "high"):
            changes["Founded"] = str(new_y)

    new_hq = new.get("hq")
    old_hq = str(old.get("hq") or "").strip()
    if new_hq and conf in {"high", "medium"}:
        if (
            not old_hq
            or old_hq.upper() in {"NPD", "N/A", "GLOBAL", "UNKNOWN"}
            or (norm(new_hq) != norm(old_hq) and conf == "high")
            or (norm(new_hq) != norm(old_hq) and (not old_hq or len(old_hq) < 8))
        ):
            changes["Headquarters"] = new_hq

    if changes:
        return {"action": "fix", "changes": changes, "reason": ",".join(changes.keys()), "confidence": conf, "source": new.get("source")}
    if new_web or new_y is not None or new_hq:
        return {"action": "confirmed", "changes": {}, "reason": "already_ok", "source": new.get("source")}
    return {"action": "skip", "changes": {}, "reason": "no_data"}


async def google_one(company: str) -> tuple[str, str]:
    """Returns (query_used, markdown)."""
    q = google_search_bar_query(company)
    data = await google_ai_ask(q, llm_clean=True)
    md = str(data.get("markdown") or data.get("markdown_raw") or "").strip()
    return q, md


async def process_company(c: dict, sem: asyncio.Semaphore, google_ok: bool) -> dict:
    async with sem:
        parsed: dict = {}
        query = google_search_bar_query(c["company"])
        phase = "google_ai"
        if google_ok:
            try:
                query, md = await google_one(c["company"])
                # reject empty / "did not match" chrome
                low = md.lower()
                if md and "did not match any documents" not in low and "for each company, give" not in low:
                    parsed = await asyncio.to_thread(extract_from_text, c["company"], md, "google_ai")
            except Exception as e:
                parsed = {}
                print(f"  google error {c['company']}: {e}", flush=True)

        if not parsed or not any([parsed.get("website"), parsed.get("founded"), parsed.get("hq")]):
            phase = "deepseek"
            try:
                parsed = await asyncio.to_thread(deepseek_residual, c)
            except Exception as e:
                print(f"  deepseek error {c['company']}: {e}", flush=True)
                parsed = {}

        if not parsed:
            rec = {
                "key": c["key"],
                "file": c["file"],
                "company": c["company"],
                "query": query,
                "processed": True,
                "phase": phase,
                "action": "skip",
                "changes": {},
                "reason": "no_data",
            }
        else:
            decision = decide_updates(c, parsed)
            rec = {
                "key": c["key"],
                "file": c["file"],
                "company": c["company"],
                "query": query,
                "processed": True,
                "phase": phase,
                **decision,
                "parsed": {
                    "website": parsed.get("website"),
                    "founded": parsed.get("founded"),
                    "hq": parsed.get("hq"),
                    "confidence": parsed.get("confidence"),
                },
            }
        return rec


def apply_results(results: list[dict]) -> None:
    by_file: dict[str, dict[str, dict]] = {}
    for fx in results:
        if fx.get("action") != "fix" or not fx.get("changes"):
            continue
        by_file.setdefault(fx["file"], {})[norm(fx["company"])] = fx["changes"]

    for fpath, mapping in by_file.items():
        wb = load_workbook(fpath)
        ws = wb.active
        hr = header_row(ws)
        headers = [str(c.value or "").strip() for c in ws[hr]]
        ci = headers.index("Company")
        n = 0
        for r in range(hr + 1, ws.max_row + 1):
            company = ws.cell(r, ci + 1).value
            if not company:
                continue
            ch = mapping.get(norm(str(company)))
            if not ch:
                continue
            for col, val in ch.items():
                if col in headers:
                    ws.cell(r, headers.index(col) + 1).value = val
            n += 1
        wb.save(fpath)
        print(f"Applied fixes to {n} rows -> {Path(fpath).name}")


async def main() -> None:
    rows = collect_rows()
    done = load_done()
    todo = [r for r in rows if r["key"] not in done]
    print(f"Companies: {len(rows)} | done: {len(done)} | remaining: {len(todo)}")
    print("Google search bar query style:")
    print('  "{company} official website founded year headquarters"')
    print("NOT instruction prompts like 'For EACH company, give: ...'")

    google_ok = await google_ai_ping()
    print(f"Google AI scraper: {'connected' if google_ok else 'DOWN'}")

    # smoke one query so user can see correct search bar text
    if google_ok and todo:
        sample = todo[0]["company"]
        q = google_search_bar_query(sample)
        print(f"\nSmoke query in search bar: {q}")
        _, md = await google_one(sample)
        print(f"Smoke markdown chars: {len(md)} | preview: {md[:180]!r}")

    prev: list[dict] = []
    if PROGRESS.exists():
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            if line.strip():
                try:
                    prev.append(json.loads(line))
                except Exception:
                    pass

    sem = asyncio.Semaphore(CONCURRENCY)
    results: list[dict] = []
    PROGRESS.parent.mkdir(parents=True, exist_ok=True)

    with PROGRESS.open("a", encoding="utf-8") as prog:
        for i in range(0, len(todo), BATCH):
            batch = todo[i : i + BATCH]
            bnum = i // BATCH + 1
            btot = (len(todo) + BATCH - 1) // BATCH
            print(f"\n=== Batch {bnum}/{btot} ({len(batch)} companies, simple Google queries) ===", flush=True)
            for c in batch:
                print(f"  search: {google_search_bar_query(c['company'])}", flush=True)

            recs = await asyncio.gather(*[process_company(c, sem, google_ok) for c in batch])
            for rec in recs:
                prog.write(json.dumps(rec, ensure_ascii=False) + "\n")
                prog.flush()
                results.append(rec)
                print(
                    f"  -> {rec['action'].upper()} {rec['company']} "
                    f"[{rec.get('phase')}] {rec.get('changes') or rec.get('reason')}",
                    flush=True,
                )
            await asyncio.sleep(0.3)

    all_recs = prev + results
    by_key: dict[str, dict] = {}
    for r in all_recs:
        if r.get("key"):
            by_key[r["key"]] = r
    final = list(by_key.values())

    print("\nApplying Excel updates...")
    apply_results(final)

    fixed = sum(1 for r in final if r.get("action") == "fix")
    confirmed = sum(1 for r in final if r.get("action") == "confirmed")
    skipped = sum(1 for r in final if r.get("action") == "skip")
    gfix = sum(1 for r in final if r.get("action") == "fix" and r.get("phase") == "google_ai")
    dfix = sum(1 for r in final if r.get("action") == "fix" and r.get("phase") == "deepseek")

    summary = {
        "total_companies": len(rows),
        "processed": len(final),
        "fixed": fixed,
        "confirmed_already_ok": confirmed,
        "skipped": skipped,
        "fixed_via_google_ai": gfix,
        "fixed_via_deepseek": dfix,
        "search_bar_query": "{company} official website founded year headquarters",
    }
    SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print("\n=== SUMMARY ===")
    for k, v in summary.items():
        print(f"{k}: {v}")


if __name__ == "__main__":
    asyncio.run(main())
