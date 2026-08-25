#!/usr/bin/env python3
"""Verify/rewrite Core Categories, Specialty Focus, Key Brands Represented for the
Global Wearable Medical Devices Market FINAL sheet.

Research sources ONLY — no LLM anywhere in this script:
  - Google AI Overview scraper (vendor_intel.clients.google_ai_scraper, :15551),
    heuristic-cleaned only (llm_clean=False — no DeepSeek pass inside the client)
  - ddgs (DuckDuckGo) as a backup/supplement search

Field extraction is pure regex/keyword heuristics over the scraped text — no LLM
call, no fabrication: a field is only written when the source text actually
contains a matching signal; otherwise the existing cell value is left untouched.

Resumable: writes one JSON checkpoint file per company under
output/web_expand/global_wearable_medical_devices_market_global/verify_checkpoints/
so a re-run skips companies already done.
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

import openpyxl
from dotenv import load_dotenv

load_dotenv(ROOT / ".env")

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping
from vendor_intel.enrichment.gap_fill.gaps import identity_text_matches

XLSX = ROOT / "output/web_expand/global_wearable_medical_devices_market_global/global_wearable_medical_devices_market_global_FINAL.xlsx"
SHEET = "Wearable Medical Companies"
CKPT_DIR = ROOT / "output/web_expand/global_wearable_medical_devices_market_global/verify_checkpoints"
CKPT_DIR.mkdir(parents=True, exist_ok=True)

CONCURRENCY = int(os.getenv("VERIFY_CONCURRENCY", "3"))
NOT_DISCLOSED = "Not publicly disclosed"

# Google/Bing bot-check block pages sometimes leak through as a "result" body —
# never treat these as real research content.
_BLOCKED_PAGE_RE = re.compile(
    r"unusual traffic|solving the above captcha|automated queries|"
    r"detected unusual traffic|systems have detected|verify you are a human|"
    r"are you a robot|enable javascript and cookies",
    re.I,
)

# Text discussing several same-named companies at once — any single sentence
# pulled from it could be describing the wrong entity, so the whole chunk is
# unusable rather than salvageable.
_AMBIGUOUS_ENTITY_RE = re.compile(
    r"can refer to (?:multiple|several|different)|distinct entities|"
    r"also known or referred to as|not to be confused with|"
    r"may refer to (?:multiple|several|a)|depend(?:s)? on the specific company",
    re.I,
)

# Non-Latin-script runs — in this dataset these are always a bleed-through
# citation/source name (e.g. an Arabic regulator name, a Japanese rating
# agency, a katakana rendering of the company name), never real content.
# Output must be English-only, so strip these runs entirely.
_NON_LATIN_CLASS = (
    "؀-ۿݐ-ݿ"  # Arabic
    "ऀ-ॿ"  # Devanagari
    "一-鿿"  # CJK unified ideographs
    "぀-ヿ"  # Hiragana + Katakana
    "가-힣"  # Hangul syllables
    "Ѐ-ӿ"  # Cyrillic
)
_NON_LATIN_RUN_RE = re.compile(
    f"[{_NON_LATIN_CLASS}][{_NON_LATIN_CLASS}\\s]*"
)

# Search-result title suffixes ("... - Crunchbase Company Profile") — noise,
# not company description.
_SITE_SUFFIX_RE = re.compile(
    r"\s*[-|]\s*(?:crunchbase(?:\.com)?(?:\s+company\s+profile)?|zoominfo|"
    r"glassdoor|dun\s*&\s*bradstreet|d&b|owler(?:\s+company\s+profile)?|"
    r"wikipedia|bloomberg|linkedin)\b.*$",
    re.I,
)

# Raw page-section labels ("Product Categories:", "Core Activities:",
# "What it does:") that Google's Overview renders as headings — strip the
# label itself but keep the sentence content that follows it.
_SECTION_LABEL_RE = re.compile(
    r"\b(?:product\s+categories|core\s+activities|what\s+it\s+does|"
    r"what\s+they\s+do|specialty\s+focus|primary\s+focus|service\s+model|"
    r"business\s+model|key\s+facts|company\s+overview)\s*:\s*",
    re.I,
)
# Same section-header phrases with NO trailing colon (Google's Overview often
# runs two bolded sub-headers together with no punctuation between them) —
# just delete the phrase itself as noise.
_BARE_LABEL_RE = re.compile(
    r"\b(?:(?:main|key|core|primary)\s+)?product\s+categories\b|"
    r"\bcore\s+activities\b|\bkey\s+facts\b|\bcompany\s+overview\b|"
    r"\bcore\s+medical\s+devices\s+and\s+uses\b|\bkey\s+products\s+and\s+uses\b|"
    r"\bbrands\s+and\s+products\b|\bshared\s+research\b|\bglobal\s+data\b|"
    r"\bcore\s+business\s*(?:&|and)?\s*operations\b|\bcore\s+business\s+divisions\b",
    re.I,
)
_BARE_URL_RE = re.compile(r"\bwww\.[a-z0-9.\-]+\.[a-z]{2,}\b", re.I)
_MD_LINK_RE = re.compile(r"\[[^\]]*\]\(https?://[^)]+\)")


def ddgs_search(query: str, max_results: int = 5) -> list[dict]:
    try:
        from ddgs import DDGS

        hits = DDGS().text(query, max_results=max_results) or []
        return [
            h for h in hits
            if not _BLOCKED_PAGE_RE.search(f"{h.get('title', '')} {h.get('body', '')}")
        ]
    except Exception as exc:
        print(f"  [ddgs] error for {query!r}: {exc}", flush=True)
        return []


# ---------------------------------------------------------------------------
# Rule-based field extraction (no LLM) -- keyword vocabulary + regex only.
# ---------------------------------------------------------------------------

# (keyword-in-text, category label) — wearable-specific checked first.
_CATEGORY_RULES: list[tuple[str, str]] = [
    (r"continuous glucose|\bcgm\b", "Continuous Glucose Monitors (CGM)"),
    (r"\becg\b|electrocardiogram|holter monitor", "Wearable ECG Monitors"),
    (r"smart\s?watch", "Smartwatches"),
    (r"fitness tracker|activity tracker", "Fitness Trackers"),
    (r"pulse oximeter", "Pulse Oximeters"),
    (r"blood pressure monitor", "Blood Pressure Monitors"),
    (r"sleep tracker", "Sleep Trackers"),
    (r"hearing aid|hearable", "Hearables / Hearing Aids"),
    (r"smart ring", "Smart Rings"),
    (r"remote patient monitoring|\brpm\b", "Remote Patient Monitoring (RPM) Devices"),
    (r"patient monitor|vital sign", "Patient Monitors"),
    (r"wearable", "Wearable Devices (general)"),
    (r"pharmaceutical", "Pharmaceuticals"),
    (r"\botc\b|over.the.counter", "OTC Healthcare"),
    (r"medical device", "Medical Devices"),
    (r"hospital suppl|hospital product", "Hospital Supplies"),
    (r"retail pharmacy|pharmacy chain|drugstore|drug store", "Retail Pharmacy"),
    (r"personal care|health\s*&\s*beauty|health and beauty", "Personal Care & Wellness"),
    (r"durable medical equipment|\bdme\b|home\s?healthcare|homecare equipment", "Home Healthcare / DME"),
    (r"laboratory|diagnostic", "Laboratory & Diagnostics"),
    (r"surgical instrument", "Surgical Instruments"),
    (r"\bdental\b", "Dental Equipment"),
    (r"generic medicine|\bgenerics\b", "Generics"),
    (r"veterinary", "Veterinary Products"),
    (r"nursing staff|homecare service|home care service", "Nursing / Homecare Services"),
    (r"electronic component|semiconductor|cybersecurity|it system|cloud service", "IT / Industrial Distribution"),
]

_BRAND_TRIGGER_RE = re.compile(
    r"(?:brands?|principals?|partners?|manufacturers?)[\s—–\-:,]*"
    r"(?:such as|include[sd]?|including|represented|distributed|carried|sold|are|:)?[\s—–\-:,]*"
    r"((?:[A-Z][\w&.'\-]+(?:\s+[A-Z][\w&.'\-]+){0,3})"
    r"(?:\s*,\s*(?:and\s+)?[A-Z][\w&.'\-]+(?:\s+[A-Z][\w&.'\-]+){0,3}){1,14})",
)

_STOPWORDS_IN_BRANDLIST = {
    "the", "and", "its", "our", "their", "other", "others", "products",
    "medical", "healthcare", "devices", "company", "group",
}

_SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+")


def extract_key_brands(text: str) -> str:
    """Regex-only brand-list extraction. No LLM, no invention."""
    best: list[str] = []
    for m in _BRAND_TRIGGER_RE.finditer(text):
        raw = m.group(1)
        parts = [p.strip(" .,;:") for p in re.split(r",|\band\b", raw) if p.strip(" .,;:")]
        parts = [
            p for p in parts
            if p
            and p.lower() not in _STOPWORDS_IN_BRANDLIST
            and len(p) > 1
            and "linkedin" not in p.lower()
            and not _NON_LATIN_RUN_RE.search(p)
        ]
        if len(parts) > len(best):
            best = parts
    if not best:
        return NOT_DISCLOSED
    seen: set[str] = set()
    out: list[str] = []
    for p in best:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    out = out[:12]
    return ", ".join(out) + (", others" if len(best) > len(out) else "")


def extract_core_categories(text: str) -> str:
    """Keyword-vocabulary tagging. Only labels a category if its keyword literally
    appears in the scraped text — never guesses a category that isn't grounded."""
    low = text.lower()
    hits: list[str] = []
    for pattern, label in _CATEGORY_RULES:
        if label in hits:
            continue
        if re.search(pattern, low, re.I):
            hits.append(label)
        if len(hits) >= 6:
            break
    return "; ".join(hits)


def extract_specialty_focus(text: str, company: str) -> str:
    """First 1-3 substantive sentences of the cleaned scrape text. No synthesis —
    lightly trims sentences already present in the source, same as a human
    highlighting the relevant lines."""
    clean_text = re.sub(r"\s+", " ", text).strip()
    sentences = [s.strip() for s in _SENTENCE_SPLIT_RE.split(clean_text) if s.strip()]
    keep: list[str] = []
    for s in sentences:
        if len(s) < 25:
            continue
        low = s.lower()
        if any(bad in low for bad in ("skip to", "sign in", "cookie", "privacy policy", "all rights reserved")):
            continue
        if _BLOCKED_PAGE_RE.search(s):
            continue
        s = _MD_LINK_RE.sub("", s).strip()
        s = _SECTION_LABEL_RE.sub("", s).strip()
        s = _BARE_LABEL_RE.sub("", s).strip()
        s = _BARE_URL_RE.sub("", s).strip()
        s = _NON_LATIN_RUN_RE.sub("", s).strip()
        s = re.sub(r"\bLinkedIn\b\s*", "", s).strip()
        s = re.sub(r"\*\*", "", s).strip()
        s = re.sub(r"\s+", " ", s).strip()
        if len(s) < 25:
            continue
        keep.append(s)
        if len(keep) >= 3 or sum(len(k) for k in keep) > 420:
            break
    return re.sub(r"\s+", " ", " ".join(keep)).strip()[:500]


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

    q1 = f"what does {company} do products categories specialty focus"
    q2 = f"{company} brands distributed partners principals"

    for q in (q1, q2):
        try:
            data = await google_ai_ask(q, llm_clean=False)
            md = str(data.get("markdown") or "").strip()
            if (
                data.get("ai_overview_missing")
                or _BLOCKED_PAGE_RE.search(md)
                or _AMBIGUOUS_ENTITY_RE.search(md)
            ):
                continue
            if md and identity_text_matches(company, website, md):
                chunks.append(md)
        except Exception as exc:
            print(f"  [google_ai] error for {company} ({q!r}): {exc}", flush=True)

    def _identity_filtered_lines(hits: list[dict]) -> list[str]:
        lines = []
        for h in hits:
            title = _SITE_SUFFIX_RE.sub("", h.get("title", "")).strip()
            line = f"{title}: {h.get('body', '')}"
            if _AMBIGUOUS_ENTITY_RE.search(line):
                continue
            if identity_text_matches(company, website, line):
                lines.append(line)
        return lines

    ddgs_hits = ddgs_search(f"{company} brands products distributor", max_results=5)
    lines = _identity_filtered_lines(ddgs_hits)
    if lines:
        chunks.append("\n".join(lines))

    if website:
        ddgs_site_hits = ddgs_search(f"site:{website} brands products", max_results=4)
        lines = _identity_filtered_lines(ddgs_site_hits)
        if lines:
            chunks.append("\n".join(lines))

    context = "\n\n".join(chunks)

    result = {
        "row": row["row"],
        "company": company,
        "core_categories": extract_core_categories(context) if context.strip() else "",
        "specialty_focus": extract_specialty_focus(context, company) if context.strip() else "",
        "key_brands": extract_key_brands(context) if context.strip() else "",
        "had_research_text": bool(context.strip()),
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
        rows.append(
            {
                "row": r,
                "company": str(company).strip(),
                "website": str(ws.cell(row=r, column=idx["Website"]).value or "").strip(),
            }
        )

    limit = int(os.getenv("VERIFY_LIMIT", "0") or "0")
    if limit > 0:
        rows = rows[:limit]

    print(f"Total companies to verify: {len(rows)}", flush=True)
    done_ckpts = len(list(CKPT_DIR.glob("*.json")))
    print(f"Already checkpointed: {done_ckpts}", flush=True)

    sem = asyncio.Semaphore(CONCURRENCY)
    completed = 0

    async def _one(row: dict) -> None:
        nonlocal completed
        async with sem:
            result = await research_company(row)
            completed += 1
            print(
                f"[{completed}/{len(rows)}] row {row['row']} {row['company']!r} "
                f"brands={'yes' if result.get('key_brands') and result['key_brands'] != NOT_DISCLOSED else 'no'} "
                f"categories={'yes' if result.get('core_categories') else 'no'}",
                flush=True,
            )

    await asyncio.gather(*[_one(row) for row in rows])

    print("Research complete. Writing results back into workbook...", flush=True)
    written = 0
    for row in rows:
        ckpt_path = CKPT_DIR / f"{row['row']:04d}.json"
        if not ckpt_path.exists():
            continue
        result = json.loads(ckpt_path.read_text(encoding="utf-8"))
        r = row["row"]
        if result.get("core_categories"):
            ws.cell(row=r, column=idx["Core Categories"], value=result["core_categories"])
        if result.get("specialty_focus"):
            ws.cell(row=r, column=idx["Specialty Focus"], value=result["specialty_focus"])
        if result.get("key_brands"):
            ws.cell(row=r, column=idx["Key Brands Represented"], value=result["key_brands"])
        written += 1

    wb.save(XLSX)
    print(f"Saved {XLSX} — {written} rows updated.", flush=True)


if __name__ == "__main__":
    asyncio.run(main())
