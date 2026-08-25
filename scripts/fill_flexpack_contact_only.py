# -*- coding: utf-8 -*-
"""Fill ONLY Contact Person / Role / LinkedIn for Global Flexible Packaging FINAL
via the Google AI scraper — a single targeted query per company, nothing else
(no revenue, no offices, no email, no fast-fill combined-ask noise).

Sales / procurement leadership only. Never invents a name — if Google AI has
nothing usable, the row is left as-is (still NPD) and can be retried later.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv
from openpyxl import load_workbook
from openai import OpenAI

load_dotenv(ROOT / ".env")

from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_enabled, google_ai_ping

XLSX = ROOT / (
    "output/web_expand/global_flexible_packaging_market_global"
    "/global_flexible_packaging_market_global_FINAL.xlsx"
)
CKPT = XLSX.parent / "flexpack_contact_only_ckpt.json"

NPD = {"", "n/a", "na", "none", "-", "not publicly disclosed", "unknown"}

ROLE_OK = re.compile(
    r"(head of sales|sales head|\bvp\b.{0,25}sales|vice president.{0,25}sales|"
    r"sales director|director of sales|director.{0,20}sales|"
    r"head of procurement|procurement head|\bvp\b.{0,25}procurement|"
    r"vice president.{0,25}procurement|procurement director|director of procurement|"
    r"head of purchasing|purchasing director|chief (?:revenue|commercial|procurement)|"
    r"\bcro\b|\bcpo\b|commercial director|head of (?:business development|channel|partner)|"
    r"national sales)",
    re.I,
)
ROLE_BAD = re.compile(r"\b(ceo|cfo|cto|coo|founder|chairman|owner|hr|marketing)\b", re.I)


def _empty(v: object) -> bool:
    return str(v or "").strip().lower() in NPD


def _openai_client() -> tuple[OpenAI, str]:
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    return OpenAI(api_key=key, base_url=base, timeout=60.0), model


def _extract(client: OpenAI, model: str, company: str, text: str) -> dict:
    """Ask DeepSeek to pull ONLY {contact_person, role, linkedin} out of the
    Google AI Overview text. Returns {} if nothing usable."""
    if not text.strip():
        return {}
    prompt = (
        f"From this Google search result text about the company \"{company}\", extract ONLY a "
        "sales or procurement leadership contact (e.g. VP Sales, Head of Procurement, Sales "
        "Director) if one is explicitly named. Do NOT return CEOs/founders/HR/marketing. "
        "Do NOT invent anything not present in the text. Return strict JSON with keys "
        "contact_person, role, linkedin (each \"\" if not found; linkedin must be a "
        "linkedin.com/in/ personal profile URL, not a company page).\n\nTEXT:\n" + text[:4000]
    )
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        return json.loads(resp.choices[0].message.content or "{}")
    except Exception:
        return {}


def _load_ckpt() -> dict:
    if CKPT.exists():
        try:
            return json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_ckpt(done: dict) -> None:
    CKPT.write_text(json.dumps(done, indent=2, ensure_ascii=False), encoding="utf-8")


async def _process_one(row, idx, client, model, sem, lock, done, n, total):
    async with sem:
        company = str(row[idx["Company"]].value or "").strip()
        website = str(row[idx["Website"]].value or "").strip()
        if not company:
            return
        domain = website.replace("https://", "").replace("http://", "").split("/")[0]
        query = (
            f'sales or procurement contact person name, their job title, and LinkedIn '
            f'profile URL at "{company}" ({domain})'
        )
        print(f"[{n}/{total}] {company} — asking Google AI…", flush=True)
        try:
            data = await google_ai_ask(query)
        except Exception as e:
            print(f"  [{n}/{total}] {company} error: {e}", flush=True)
            async with lock:
                done[company.lower()] = {"skipped": "ask_error"}
                _save_ckpt(done)
            return

        text = str(data.get("markdown") or "")
        extracted = _extract(client, model, company, text)
        person = str(extracted.get("contact_person") or "").strip()
        role = str(extracted.get("role") or "").strip()
        linkedin = str(extracted.get("linkedin") or "").strip()

        if role and ROLE_BAD.search(role) and not ROLE_OK.search(role):
            person, role, linkedin = "", "", ""

        async with lock:
            wrote = []
            if person and _empty(row[idx["Contact Person"]].value):
                row[idx["Contact Person"]].value = person
                wrote.append("Contact Person")
            if role and _empty(row[idx["Role"]].value):
                row[idx["Role"]].value = role
                wrote.append("Role")
            if linkedin and "linkedin.com/in/" in linkedin.lower() and _empty(row[idx["LinkedIn"]].value):
                row[idx["LinkedIn"]].value = linkedin.split("?")[0].rstrip("/")
                wrote.append("LinkedIn")
            done[company.lower()] = {"wrote": wrote}
            _save_ckpt(done)
        print(f"  [{n}/{total}] {company} -> {', '.join(wrote) if wrote else 'nothing found'}", flush=True)


async def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--concurrency", type=int, default=20)
    args = ap.parse_args()

    if not google_ai_enabled():
        print("GOOGLE_AI_SCRAPER_ENABLED is not true — aborting.")
        return
    if not await google_ai_ping():
        print("Google AI scraper backend is not reachable/healthy — aborting.")
        return

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    done = _load_ckpt()
    client, model = _openai_client()

    rows = list(ws.iter_rows(min_row=2))
    targets = [
        r for r in rows
        if _empty(r[idx["Contact Person"]].value)
        and str(r[idx["Company"]].value or "").strip().lower() not in done
    ]
    if args.limit:
        targets = targets[: args.limit]
    total = len(targets)
    print(f"total rows: {len(rows)} | still missing Contact Person: {total} | concurrency: {args.concurrency}", flush=True)

    sem = asyncio.Semaphore(args.concurrency)
    lock = asyncio.Lock()

    # Process in chunks of `concurrency` so we can checkpoint-save the workbook between batches
    for start in range(0, total, args.concurrency):
        batch = targets[start : start + args.concurrency]
        await asyncio.gather(
            *[
                _process_one(row, idx, client, model, sem, lock, done, start + i + 1, total)
                for i, row in enumerate(batch)
            ]
        )
        wb.save(XLSX)
        print(f"  [checkpoint save at {min(start + args.concurrency, total)}/{total}]", flush=True)

    wb.save(XLSX)
    print(f"done. saved: {XLSX}", flush=True)


if __name__ == "__main__":
    asyncio.run(main())
