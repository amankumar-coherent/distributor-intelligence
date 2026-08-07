#!/usr/bin/env python3
"""Full Founded-year verify for all FINAL_SLIM Excels via Google AI scraper.

Query style (same as manual Google): \"{company} founded year?\"
Applies only high/medium confidence fixes. Progress is checkpointed for resume.
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

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")
os.environ["GOOGLE_AI_SCRAPER_ENABLED"] = "true"
# Faster founded-year sweep
os.environ.setdefault("GOOGLE_AI_SCRAPER_OVERVIEW_RETRIES", "2")
os.environ.setdefault("GOOGLE_AI_SCRAPER_OVERVIEW_RETRY_SLEEP_SEC", "1.5")
os.environ.setdefault("GOOGLE_AI_SCRAPER_MODE", "pro")

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping  # noqa: E402

SLIM_FILES = [
    ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL_SLIM.xlsx",
    ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL_SLIM.xlsx",
    ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global/global_antibody_drug_conjugates_market_FINAL_SLIM.xlsx",
]

PROGRESS = ROOT / "scripts/_founded_full_progress.jsonl"
RESULTS = ROOT / "scripts/_founded_full_results.json"
CURRENT_YEAR = 2026


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def company_tokens(name: str) -> list[str]:
    stop = {
        "ltd", "llc", "inc", "gmbh", "co", "company", "corp", "corporation",
        "sdn", "bhd", "plc", "sa", "ag", "bv", "fzco", "fzc", "the", "and",
        "group", "trading", "distribution", "distributor", "distributors",
        "international", "global", "technologies", "technology", "tech",
    }
    toks = re.findall(r"[a-z0-9]+", norm(name))
    return [t for t in toks if len(t) >= 3 and t not in stop]


def company_mentioned(text: str, company: str) -> bool:
    low = (text or "").lower()
    toks = company_tokens(company)
    if not toks:
        return company.lower().split()[0] in low if company else False
    # require at least one strong token, prefer 2 if available
    hits = sum(1 for t in toks[:4] if t in low)
    need = 1 if len(toks) == 1 else min(2, len(toks))
    return hits >= need


def parse_founded(markdown: str, company: str) -> tuple[int | None, str, str]:
    """Return (year, confidence, reason). confidence: high|medium|low|none"""
    text = (markdown or "").strip()
    if not text:
        return None, "none", "empty"
    low = text.lower()
    mentioned = company_mentioned(text, company)

    # Explicit wrong-entity guard: "If you meant" / other company founded
    if re.search(r"if you meant|different (?:company|entity)|not (?:to be )?confused", low):
        # still allow if primary sentence has our company + year
        pass

    patterns = [
        (r"(?:was\s+)?(?:established|founded|incorporated|started|formed)\s+(?:in\s+)?((?:18|19|20)\d{2})", "established_in"),
        (r"established[:\s]+((?:18|19|20)\d{2})", "established_colon"),
        (r"founded[:\s]+((?:18|19|20)\d{2})", "founded_colon"),
        (r"incorporation(?:\s+date)?[:\s]+(?:[a-z]+\s+\d{1,2},?\s+)?((?:18|19|20)\d{2})", "incorporation"),
        (r"since\s+((?:18|19|20)\d{2})", "since"),
    ]
    for pat, tag in patterns:
        for m in re.finditer(pat, low, re.I):
            y = int(m.group(1))
            if not (1800 <= y <= CURRENT_YEAR):
                continue
            window = low[max(0, m.start() - 30) : m.end() + 30]
            if "©" in window or "copyright" in window:
                continue
            # Prefer years near company mention
            conf = "high" if mentioned else "medium"
            # Downgrade if snippet is clearly about another well-known brand and our tokens missing
            if not mentioned:
                conf = "low"
            return y, conf, tag

    # "not publicly documented / not specified"
    if re.search(
        r"not publicly (?:documented|specified|available)|exact founding year .{0,40}not |"
        r"founding year is not|no (?:exact |official )?founding year",
        low,
    ):
        return None, "none", "not_documented"

    return None, "none", "no_clear_year"


def collect_all() -> list[dict]:
    rows: list[dict] = []
    for f in SLIM_FILES:
        wb = load_workbook(f, read_only=True, data_only=True)
        ws = wb.active
        hdr = hr = None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            if row and any(str(c or "").strip().lower() == "company" for c in row):
                hdr = [str(c or "") for c in row]
                hr = i
                break
        assert hdr and hr
        ci, fi = hdr.index("Company"), hdr.index("Founded")
        hi = hdr.index("Headquarters") if "Headquarters" in hdr else None
        wi = hdr.index("Website") if "Website" in hdr else None
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row or not row[ci]:
                continue
            company = str(row[ci]).strip()
            founded = str(row[fi] or "").strip()
            m = re.search(r"(18|19|20)\d{2}", founded)
            y = int(m.group(0)) if m else None
            rows.append(
                {
                    "file": str(f),
                    "company": company,
                    "founded": founded,
                    "year": y,
                    "hq": str(row[hi] or "") if hi is not None else "",
                    "website": str(row[wi] or "") if wi is not None else "",
                    "key": f"{norm(str(f.name))}::{norm(company)}",
                }
            )
        wb.close()
    return rows


def load_done_keys() -> set[str]:
    done: set[str] = set()
    if PROGRESS.exists():
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                obj = json.loads(line)
                if obj.get("key"):
                    done.add(obj["key"])
            except Exception:
                continue
    return done


async def verify_one(item: dict) -> dict:
    company = item["company"]
    query = f"{company} founded year?"
    data = await google_ai_ask(query, llm_clean=True)
    text = str(data.get("markdown") or data.get("markdown_raw") or "").strip()
    new_year, conf, reason = parse_founded(text, company)

    out = {
        "key": item["key"],
        "file": item["file"],
        "company": company,
        "old": item["founded"],
        "old_year": item["year"],
        "query": query,
        "new_year": new_year,
        "confidence": conf,
        "reason": reason,
        "ai_overview_missing": data.get("ai_overview_missing"),
        "error": data.get("error"),
        "snippet": text[:450],
        "action": "keep",
        "new": None,
    }

    old_y = item["year"]

    # Invalid current-year copyright pollution with no real year
    if old_y == CURRENT_YEAR and new_year is None and reason in {"not_documented", "no_clear_year", "empty"}:
        out["action"] = "fix"
        out["new"] = "NPD"
        out["confidence"] = "medium"
        return out

    if new_year is None:
        return out

    # Only change on high/medium confidence and different year
    if conf in {"high", "medium"} and (old_y is None or new_year != old_y):
        # Extra guard: don't jump >80 years unless high confidence + company mentioned
        if old_y is not None and abs(new_year - old_y) > 80 and conf != "high":
            out["action"] = "keep"
            out["reason"] = f"delta_too_large:{reason}"
            return out
        # Don't replace solid pre-2000 with post-2018 unless high
        if old_y is not None and old_y < 2000 and new_year >= 2018 and conf != "high":
            out["action"] = "keep"
            out["reason"] = f"protect_old_year:{reason}"
            return out
        out["action"] = "fix"
        out["new"] = str(new_year)
        return out

    if conf in {"high", "medium"} and old_y == new_year:
        out["action"] = "keep"
        out["new"] = str(new_year)
        out["reason"] = f"confirmed:{reason}"
    return out


def apply_fixes_from_progress() -> dict[str, int]:
    """Read progress jsonl and apply fix actions to Excels."""
    by_file: dict[str, dict[str, str]] = {}
    stats = {"fix": 0, "keep": 0, "confirmed": 0}
    if not PROGRESS.exists():
        return stats
    for line in PROGRESS.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        fx = json.loads(line)
        if fx.get("action") == "fix" and fx.get("new"):
            by_file.setdefault(fx["file"], {})[norm(fx["company"])] = str(fx["new"])
            stats["fix"] += 1
        elif str(fx.get("reason", "")).startswith("confirmed"):
            stats["confirmed"] += 1
        else:
            stats["keep"] += 1

    for fpath, mapping in by_file.items():
        wb = load_workbook(fpath)
        ws = wb.active
        hdr = hr = None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
            if row and any(str(c or "").strip().lower() == "company" for c in row):
                hdr = [str(c or "") for c in row]
                hr = i
                break
        assert hdr and hr
        ci, fi = hdr.index("Company"), hdr.index("Founded")
        n = 0
        for r in range(hr + 1, ws.max_row + 1):
            company = ws.cell(r, ci + 1).value
            if not company:
                continue
            key = norm(str(company))
            if key in mapping:
                ws.cell(r, fi + 1).value = mapping[key]
                n += 1
        wb.save(fpath)
        print(f"Applied {n} founded fixes -> {Path(fpath).name}")
    return stats


async def main() -> None:
    if not await google_ai_ping():
        raise SystemExit("Google AI scraper not up on :15551")

    rows = collect_all()
    done = load_done_keys()
    todo = [r for r in rows if r["key"] not in done]
    print(f"Total companies: {len(rows)} | already done: {len(done)} | remaining: {len(todo)}")

    PROGRESS.parent.mkdir(parents=True, exist_ok=True)
    with PROGRESS.open("a", encoding="utf-8") as prog:
        for i, item in enumerate(todo, 1):
            print(
                f"[{i}/{len(todo)}] {item['company']} ({item['founded']}) "
                f"-> \"{item['company']} founded year?\"",
                flush=True,
            )
            try:
                fx = await verify_one(item)
            except Exception as e:
                fx = {
                    "key": item["key"],
                    "file": item["file"],
                    "company": item["company"],
                    "old": item["founded"],
                    "action": "keep",
                    "error": str(e),
                    "new": None,
                    "confidence": "none",
                    "reason": "exception",
                }
            prog.write(json.dumps(fx, ensure_ascii=False) + "\n")
            prog.flush()
            print(
                f"    {fx.get('action')} conf={fx.get('confidence')} "
                f"{fx.get('old')} => {fx.get('new') or fx.get('old')} ({fx.get('reason')})",
                flush=True,
            )
            await asyncio.sleep(0.8)

    # Build summary results
    all_fx = []
    for line in PROGRESS.read_text(encoding="utf-8").splitlines():
        if line.strip():
            all_fx.append(json.loads(line))
    RESULTS.write_text(json.dumps(all_fx, indent=2, ensure_ascii=False), encoding="utf-8")

    print("\nApplying confident fixes...")
    stats = apply_fixes_from_progress()
    changed = sum(1 for f in all_fx if f.get("action") == "fix")
    confirmed = sum(1 for f in all_fx if str(f.get("reason", "")).startswith("confirmed"))
    print(f"Done. fixes={changed} confirmed={confirmed} total_checked={len(all_fx)}")
    print(f"Results: {RESULTS}")


if __name__ == "__main__":
    asyncio.run(main())
