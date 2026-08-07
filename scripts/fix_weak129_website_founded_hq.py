#!/usr/bin/env python3
"""Re-verify & fix the 129 weak/skipped companies (Victoria-type).

Search bar style:
  {company} official website founded year headquarters
  {company} founded year?

Google AI first → DDGS snippets → DeepSeek extract.
If solid evidence: update Excel.
If Google says not found / no solid evidence: set Website/Founded/HQ to NPD.
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
os.environ["GOOGLE_AI_SCRAPER_FORCE_OVERVIEW"] = "false"
os.environ["GOOGLE_AI_SCRAPER_OVERVIEW_RETRIES"] = "1"

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping  # noqa: E402

TODO = ROOT / "scripts/_weak129_todo.json"
OUT = ROOT / "scripts/_weak129_fix_progress.jsonl"
SUMMARY = ROOT / "scripts/_weak129_fix_summary.json"
CURRENT_YEAR = 2026
CONCURRENCY = 4


def norm(s: object) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def header_row(ws) -> int:
    for r in range(1, min(6, ws.max_row + 1)):
        vals = [norm(c.value) for c in ws[r] if c.value not in (None, "")]
        if "company" in vals:
            return r
    return 1


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
            "google.com",
            "wikipedia.org",
            "tiktok.com",
            "linkedin.com",
            "crunchbase.com",
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
    return yi if 1800 <= yi <= CURRENT_YEAR else None


def company_tokens(name: str) -> list[str]:
    stop = {
        "ltd", "llc", "inc", "gmbh", "co", "company", "corp", "the", "and",
        "group", "sa", "ag", "bv", "plc", "sdn", "bhd", "distribution",
        "distributor", "pharmacy", "specialty", "international", "global",
    }
    toks = re.findall(r"[a-z0-9]+", norm(name))
    return [t for t in toks if len(t) >= 3 and t not in stop]


def mentioned(text: str, company: str) -> bool:
    low = (text or "").lower()
    toks = company_tokens(company)
    if not toks:
        return True
    hits = sum(1 for t in toks[:4] if t in low)
    return hits >= (1 if len(toks) == 1 else min(2, len(toks)))


def not_found_signal(text: str) -> bool:
    low = (text or "").lower()
    patterns = [
        r"did not match any documents",
        r"no (?:major |official |widely documented )?(?:company|entity|website)",
        r"not (?:publicly )?(?:indexed|documented|available|found)",
        r"could not find",
        r"no results",
        r"unable to (?:find|locate)",
        r"there is no .{0,40}(?:official website|founding)",
    ]
    return any(re.search(p, low) for p in patterns)


def deepseek() -> OpenAI:
    return OpenAI(
        api_key=os.getenv("OPENAI_API_KEY") or "",
        base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/"),
    )


def extract(company: str, text: str, source: str) -> dict:
    if not (text or "").strip():
        return {"status": "empty", "source": source}
    if not_found_signal(text) and not mentioned(text, company):
        return {"status": "not_found", "source": source}
    if "for each company, give" in text.lower():
        return {"status": "empty", "source": source}

    client = deepseek()
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    prompt = f"""Extract firmographics ONLY if clearly about this exact company.

Company: {company}

Search/AI text:
```
{text[:10000]}
```

Return JSON only:
{{
  "status": "found"|"not_found"|"unclear",
  "website": "https://..." or null,
  "founded_year": 1994 or null,
  "headquarters": "City, Country" or null,
  "confidence": "high"|"medium"|"low",
  "reason": "short"
}}
Rules:
- If text says no official entity / not documented → status=not_found.
- If wrong company / ambiguous → status=unclear, null fields.
- Official website only (not LinkedIn/Wikipedia/news aggregators).
- founded_year 1800-{CURRENT_YEAR}; never copyright year alone.
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
        return {"status": "unclear", "source": source}
    if not isinstance(data, dict):
        return {"status": "unclear", "source": source}
    status = str(data.get("status") or "unclear").lower()
    conf = str(data.get("confidence") or "low").lower()
    out = {
        "status": status,
        "website": clean_url(data.get("website")),
        "founded": year_ok(data.get("founded_year")),
        "hq": (str(data.get("headquarters") or "").strip() or None),
        "confidence": conf,
        "reason": data.get("reason") or "",
        "source": source,
    }
    # downgrade invents
    if conf == "low":
        out["status"] = "unclear" if status == "found" else status
    if status == "found" and not any([out["website"], out["founded"], out["hq"]]):
        out["status"] = "unclear"
    return out


def ddg_text(company: str) -> str:
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS  # type: ignore
        except ImportError:
            return ""
    queries = [
        f"{company} official website founded year headquarters",
        f"{company} founded year?",
        f"{company} headquarters company",
    ]
    bits = []
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


def load_excel_row(filepath: str, company: str) -> dict:
    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    hr = header_row(ws)
    headers = [str(c.value or "").strip() for c in ws[hr]]
    ci = headers.index("Company")
    out = {"website": "", "founded": "", "hq": ""}
    for row in ws.iter_rows(min_row=hr + 1, values_only=True):
        if row and row[ci] and norm(row[ci]) == norm(company):
            out = {
                "website": str(row[headers.index("Website")] or "").strip() if "Website" in headers else "",
                "founded": str(row[headers.index("Founded")] or "").strip() if "Founded" in headers else "",
                "hq": str(row[headers.index("Headquarters")] or "").strip() if "Headquarters" in headers else "",
            }
            break
    wb.close()
    return out


async def google_bundle(company: str) -> str:
    texts = []
    for q in [
        f"{company} official website founded year headquarters",
        f"{company} founded year?",
    ]:
        try:
            data = await google_ai_ask(q, llm_clean=True)
            md = str(data.get("markdown") or data.get("markdown_raw") or "").strip()
            if md:
                texts.append(f"[query] {q}\n{md}")
        except Exception as e:
            texts.append(f"[query] {q}\n(error {e})")
        await asyncio.sleep(0.25)
    return "\n\n".join(texts)


async def process_one(item: dict, sem: asyncio.Semaphore, google_ok: bool) -> dict:
    async with sem:
        company = item["company"]
        filepath = item["filepath"]
        old = await asyncio.to_thread(load_excel_row, filepath, company)
        print(f"→ {company}", flush=True)

        parsed = {"status": "empty"}
        if google_ok:
            gtext = await google_bundle(company)
            parsed = await asyncio.to_thread(extract, company, gtext, "google_ai")
            print(f"  google status={parsed.get('status')} conf={parsed.get('confidence')}", flush=True)

        if parsed.get("status") not in {"found"}:
            dtext = await asyncio.to_thread(ddg_text, company)
            # also ask deepseek with ddg + google text remnant
            bundle = f"GOOGLE:\n{gtext if google_ok else '(none)'}\n\nDDG:\n{dtext}"
            parsed2 = await asyncio.to_thread(extract, company, bundle, "web+deepseek")
            print(f"  residual status={parsed2.get('status')} conf={parsed2.get('confidence')}", flush=True)
            # prefer found over not_found/unclear
            if parsed2.get("status") == "found":
                parsed = parsed2
            elif parsed.get("status") != "not_found" and parsed2.get("status") == "not_found":
                parsed = parsed2
            elif parsed.get("status") in {"empty", "unclear"} and parsed2.get("status") in {"found", "not_found", "unclear"}:
                # keep richer
                if parsed2.get("status") == "found" or parsed.get("status") == "empty":
                    parsed = parsed2

        # Decide Excel write
        changes: dict[str, str] = {}
        action = "keep"
        status = parsed.get("status")
        conf = parsed.get("confidence")

        if status == "found" and conf in {"high", "medium"}:
            if parsed.get("website"):
                changes["Website"] = parsed["website"]
            if parsed.get("founded") is not None:
                changes["Founded"] = str(parsed["founded"])
            if parsed.get("hq"):
                changes["Headquarters"] = parsed["hq"]
            action = "fix" if changes else "keep_ok"
        elif status == "not_found":
            changes = {"Website": "NPD", "Founded": "NPD", "Headquarters": "NPD"}
            action = "npd"
        else:
            # unclear/empty: keep existing if it looks real; else NPD
            ow = clean_url(old.get("website"))
            oy = year_ok(old.get("founded"))
            oh = str(old.get("hq") or "").strip()
            looks_ok = bool(ow) or (oy is not None) or (oh and oh.upper() not in {"NPD", "N/A", "GLOBAL", "UNKNOWN"} and len(oh) >= 5)
            if looks_ok:
                action = "keep_existing"
                changes = {}
            else:
                changes = {"Website": "NPD", "Founded": "NPD", "Headquarters": "NPD"}
                action = "npd"

        return {
            "company": company,
            "file": filepath,
            "old": old,
            "parsed": parsed,
            "action": action,
            "changes": changes,
        }


def apply(results: list[dict]) -> None:
    by_file: dict[str, dict[str, dict]] = {}
    for r in results:
        if not r.get("changes"):
            continue
        by_file.setdefault(r["file"], {})[norm(r["company"])] = r["changes"]

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
        print(f"Updated {n} rows -> {Path(fpath).name}")


async def main() -> None:
    items = json.loads(TODO.read_text(encoding="utf-8"))
    print(f"Fixing {len(items)} weak/skipped companies")
    google_ok = await google_ai_ping()
    print(f"Google AI scraper: {'connected' if google_ok else 'DOWN'}")

    if OUT.exists():
        OUT.unlink()

    sem = asyncio.Semaphore(CONCURRENCY)
    results: list[dict] = []
    with OUT.open("a", encoding="utf-8") as prog:
        # process in batches of 20 for logging
        for i in range(0, len(items), 20):
            batch = items[i : i + 20]
            print(f"\n=== Batch {i//20+1}/{(len(items)+19)//20} ===", flush=True)
            recs = await asyncio.gather(*[process_one(x, sem, google_ok) for x in batch])
            for rec in recs:
                prog.write(json.dumps(rec, ensure_ascii=False) + "\n")
                prog.flush()
                results.append(rec)
                print(
                    f"  {rec['action'].upper()} {rec['company']}: {rec.get('changes')}",
                    flush=True,
                )

    print("\nApplying to Excel...")
    apply(results)

    summary = {
        "total": len(results),
        "fixed_with_data": sum(1 for r in results if r["action"] == "fix"),
        "set_to_npd": sum(1 for r in results if r["action"] == "npd"),
        "keep_existing": sum(1 for r in results if r["action"] == "keep_existing"),
        "keep_ok": sum(1 for r in results if r["action"] == "keep_ok"),
    }
    SUMMARY.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print("\n=== SUMMARY ===")
    for k, v in summary.items():
        print(f"{k}: {v}")


if __name__ == "__main__":
    asyncio.run(main())
