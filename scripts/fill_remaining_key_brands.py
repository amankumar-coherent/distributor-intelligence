#!/usr/bin/env python3
"""Second-pass attempt at filling the remaining Key Brands Represented gaps.

Same stack as before -- Google AI Overview scraper + ddgs, no LLM, no
subagents -- but with more targeted brand-specific queries and a broader
brand-list trigger regex. Only overwrites a row when a real brand-list match
is found; otherwise leaves "Not publicly disclosed" untouched.

Resumable: one JSON checkpoint per company under
output/web_expand/global_wearable_medical_devices_market_global/verify_checkpoints_brands2/
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "scripts"))

import openpyxl
from dotenv import load_dotenv

load_dotenv(ROOT / ".env")

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping
from vendor_intel.enrichment.gap_fill.gaps import identity_text_matches

import verify_wearable_categories_brands as base  # reuse filters/regex

XLSX = ROOT / "output/web_expand/global_wearable_medical_devices_market_global/global_wearable_medical_devices_market_global_FINAL.xlsx"
SHEET = "Wearable Medical Companies"
CKPT_DIR = ROOT / "output/web_expand/global_wearable_medical_devices_market_global/verify_checkpoints_brands2"
CKPT_DIR.mkdir(parents=True, exist_ok=True)

CONCURRENCY = int(os.getenv("BRANDS2_CONCURRENCY", "3"))
NOT_DISCLOSED = "Not publicly disclosed"

_BRAND_TRIGGER_RE2 = re.compile(
    r"(?:brands?|principals?|partners?|manufacturers?|distributors? of|dealer of|"
    r"reseller of|products? (?:from|of)|represents?|carries|distributes?)"
    r"[\s—–\-:,]*(?:such as|include[sd]?|including|represented|distributed|carried|"
    r"sold|are|:)?[\s—–\-:,]*"
    r"((?:[A-Z][\w&.'\-]+(?:\s+[A-Z][\w&.'\-]+){0,3})"
    r"(?:\s*,\s*(?:and\s+)?[A-Z][\w&.'\-]+(?:\s+[A-Z][\w&.'\-]+){0,3}){1,14})",
)

# Second signal: many company profiles list subsidiary/brand names as bullet
# labels under a "Brands"/"Subsidiary Brands"/"Brand Partners" heading, e.g.
#   Major Subsidiary Brands
#   Triolab Group: Focuses on clinical diagnostics...
#   Mediplast: Develops and distributes...
# rather than prose "brands such as X, Y, Z" -- catch that structure too.
# Must be a standalone heading LINE (nothing else on that line) -- matching
# the bare word anywhere mid-sentence (e.g. "...over 500 proprietary brands
# with partnerships...") produces false positives on whatever bullet-like
# text happens to follow, unrelated to any real brand list.
_BRAND_HEADING_RE = re.compile(
    r"^[ \t]*(?:major\s+|key\s+|principal\s+|represented\s+)?"
    r"(?:subsidiary\s+)?brands?[ \t]*$|"
    r"^[ \t]*brand\s+partners?[ \t]*$",
    re.I | re.M,
)
_BULLET_LABEL_RE = re.compile(r"([A-Z][A-Za-z0-9&/.'\- ]{1,40}):")
_COMPANY_STOP_TOKENS = {
    "inc", "ltd", "llc", "corp", "corporation", "company", "co", "group",
    "the", "and", "ab", "gmbh", "plc", "holdings", "international",
}


def _company_tokens(company: str) -> set[str]:
    toks = re.sub(r"[^a-z0-9\s]", " ", company.lower()).split()
    return {t for t in toks if len(t) >= 3 and t not in _COMPANY_STOP_TOKENS}


_BULLET_JUNK = {
    "recent additions", "note", "source", "overview", "summary", "business focus",
    "primary products", "operational status", "incorporation date", "registered office",
    "industry", "stock listing", "corporate identification number", "basic information",
    "major subsidiary brands", "key facts", "company overview", "core business",
}


def _extract_bullet_brands(text: str, company_tokens: set[str]) -> list[str]:
    m = _BRAND_HEADING_RE.search(text)
    if not m:
        return []
    window = text[m.end(): m.end() + 1200]
    out: list[str] = []
    for lm in _BULLET_LABEL_RE.finditer(window):
        label = lm.group(1).strip()
        low = label.lower()
        if low in _BULLET_JUNK or len(label) < 2:
            continue
        if any(t and t in low for t in company_tokens):
            continue
        out.append(label)
    return out[:10]


def extract_key_brands2(text: str, company: str) -> str:
    tokens = _company_tokens(company)
    best: list[str] = []
    for m in _BRAND_TRIGGER_RE2.finditer(text):
        raw = m.group(1)
        parts = [p.strip(" .,;:") for p in re.split(r",|\band\b", raw) if p.strip(" .,;:")]
        parts = [
            p for p in parts
            if p
            and p.lower() not in base._STOPWORDS_IN_BRANDLIST
            and len(p) > 1
            and "linkedin" not in p.lower()
            and not base._NON_LATIN_RUN_RE.search(p)
        ]
        if len(parts) > len(best):
            best = parts

    bullet_hits = _extract_bullet_brands(text, tokens)

    combined = best + [b for b in bullet_hits if b not in best]
    if not combined:
        return ""
    seen: set[str] = set()
    out: list[str] = []
    for p in combined:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    out = out[:12]
    return "; ".join(out) + ("; others" if len(combined) > len(out) else "")


async def research_company(row: dict) -> dict:
    company = row["company"]
    website = row["website"] or ""
    ckpt_path = CKPT_DIR / f"{row['row']:04d}.json"
    if ckpt_path.exists():
        try:
            return json.loads(ckpt_path.read_text(encoding="utf-8"))
        except Exception:
            pass

    chunks: list[str] = []
    # Natural "what does X do" phrasing reliably triggers a real Google AI
    # Overview (this is what worked in the first research pass); narrow
    # keyword-stuffed brand queries mostly returned no overview at all.
    queries = [
        f"what brands or manufacturers does {company} sell or distribute",
        f"{company} product brands and manufacturer partners",
    ]
    for q in queries:
        try:
            data = await google_ai_ask(q, llm_clean=False)
            md = str(data.get("markdown") or "").strip()
            if (
                data.get("ai_overview_missing")
                or base._BLOCKED_PAGE_RE.search(md)
                or base._AMBIGUOUS_ENTITY_RE.search(md)
            ):
                continue
            if md and identity_text_matches(company, website, md):
                chunks.append(md)
        except Exception as exc:
            print(f"  [google_ai] error for {company} ({q!r}): {exc}", flush=True)

    ddgs_hits = base.ddgs_search(f'"{company}" brand partners OR principals OR "we represent"', max_results=6)
    lines = []
    for h in ddgs_hits:
        title = base._SITE_SUFFIX_RE.sub("", h.get("title", "")).strip()
        line = f"{title}: {h.get('body', '')}"
        if base._AMBIGUOUS_ENTITY_RE.search(line):
            continue
        if identity_text_matches(company, website, line):
            lines.append(line)
    if lines:
        chunks.append("\n".join(lines))

    context = "\n\n".join(chunks)
    key_brands = extract_key_brands2(context, company) if context.strip() else ""

    result = {
        "row": row["row"],
        "company": company,
        "key_brands": key_brands,
        "had_research_text": bool(context.strip()),
        "raw_context": context[:4000],
    }
    ckpt_path.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")
    return result


async def main() -> None:
    if not await google_ai_ping():
        print("ERROR: google-ai-scraper backend not reachable at :15551 — aborting.", flush=True)
        sys.exit(1)

    wb = openpyxl.load_workbook(XLSX)
    ws = wb[SHEET]
    header = [c.value for c in ws[1]]
    idx = {h: i + 1 for i, h in enumerate(header)}

    rows: list[dict] = []
    for r in range(2, ws.max_row + 1):
        company = ws.cell(row=r, column=idx["Company"]).value
        if not company:
            continue
        current = str(ws.cell(row=r, column=idx["Key Brands Represented"]).value or "").strip()
        if current and current.lower() != NOT_DISCLOSED.lower():
            continue
        rows.append(
            {
                "row": r,
                "company": str(company).strip(),
                "website": str(ws.cell(row=r, column=idx["Website"]).value or "").strip(),
            }
        )

    limit = int(os.getenv("BRANDS2_LIMIT", "0") or "0")
    if limit > 0:
        rows = rows[:limit]

    print(f"Companies to research: {len(rows)}", flush=True)

    sem = asyncio.Semaphore(CONCURRENCY)
    completed = 0

    async def _one(row: dict) -> None:
        nonlocal completed
        async with sem:
            result = await research_company(row)
            completed += 1
            print(
                f"[{completed}/{len(rows)}] row {row['row']} {row['company']!r} "
                f"found={'yes' if result.get('key_brands') else 'no'}",
                flush=True,
            )

    await asyncio.gather(*[_one(row) for row in rows])

    hits = 0
    for row in rows:
        ckpt_path = CKPT_DIR / f"{row['row']:04d}.json"
        if not ckpt_path.exists():
            continue
        result = json.loads(ckpt_path.read_text(encoding="utf-8"))
        if result.get("key_brands"):
            hits += 1
    print(
        f"Research complete: {hits} rows have a candidate brand list in checkpoints. "
        "NOT auto-writing to the workbook -- run a manual review pass before committing.",
        flush=True,
    )


if __name__ == "__main__":
    asyncio.run(main())
