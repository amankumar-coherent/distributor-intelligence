#!/usr/bin/env python3
"""Fill empty Founded (and Role if contact exists) via OpenAI web_search."""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv(ROOT / ".env")

from clean_verify_five_finals import (  # noqa: E402
    MARKETS,
    _empty,
    load_sheet_rows,
    write_company_sheet,
)
from vendor_intel.pipeline.chatgpt_expand import _web_search_json  # noqa: E402

YEAR_RE = re.compile(r"\b(18|19|20)\d{2}\b")


async def fill_founded(company: str, website: str, client: OpenAI, model: str) -> str:
    try:
        data = await asyncio.to_thread(
            _web_search_json,
            client,
            model,
            (
                "Find the company founding year. Return JSON only: "
                '{"founded":"YYYY or empty","confidence":"high|medium|low"} '
                "Do NOT invent. Prefer official About / LinkedIn / Wikipedia."
            ),
            f'Company: {company}\nWebsite: {website}\nSearch: "{company}" founded OR founded in OR established',
            label="founded-web",
            max_tokens=200,
        )
    except Exception:
        return ""
    if not isinstance(data, dict):
        return ""
    raw = str(data.get("founded") or "").strip()
    m = YEAR_RE.search(raw)
    return m.group(0) if m else ""


async def main() -> None:
    key = os.getenv("FILL_FOUNDED_MARKET", "smartphone")
    markets = [m for m in MARKETS if m["key"] == key]
    if not markets:
        raise SystemExit(f"unknown market {key}")
    m = markets[0]
    path = m["path"]
    headers, rows, _ = load_sheet_rows(path, m["sheet"])
    todo = [i for i, r in enumerate(rows) if _empty(r.get("Founded"))]
    limit = os.getenv("FOUNDED_FILL_LIMIT")
    if limit and limit.isdigit():
        todo = todo[: int(limit)]
    print(f"Founded fill queue: {len(todo)} / empty={sum(1 for r in rows if _empty(r.get('Founded')))}", flush=True)

    client = OpenAI(
        api_key=(os.getenv("OPENAI_API_KEY") or "").strip(),
        base_url=(os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").rstrip("/"),
    )
    model = (os.getenv("OPENAI_MODEL") or "gpt-4o-mini").strip()
    sem = asyncio.Semaphore(int(os.getenv("FOUNDED_FILL_CONCURRENT") or "3"))
    filled = 0
    ckpt = path.parent / f"founded_fill_checkpoint_{key}.json"
    done: dict[str, str] = {}
    if ckpt.exists():
        done = json.loads(ckpt.read_text(encoding="utf-8"))
    for r in rows:
        c = r.get("Company") or ""
        if c in done and _empty(r.get("Founded")):
            r["Founded"] = done[c]

    todo = [i for i, r in enumerate(rows) if _empty(r.get("Founded"))]
    if limit and limit.isdigit():
        todo = todo[: int(limit)]

    async def one(i: int) -> None:
        nonlocal filled
        row = rows[i]
        company = row.get("Company") or ""
        async with sem:
            print(f"  [founded] {company[:50].encode('ascii','replace').decode('ascii')}", flush=True)
            year = await fill_founded(company, row.get("Website") or "", client, model)
            if year:
                row["Founded"] = year
                done[company] = year
                filled += 1
                ckpt.write_text(json.dumps(done, indent=2), encoding="utf-8")

    await asyncio.gather(*[one(i) for i in todo])
    ckpt.write_text(json.dumps(done, indent=2), encoding="utf-8")

    out = path.with_name(path.stem + "_CLEANED.xlsx")
    write_company_sheet(
        out,
        sheet_name=m["sheet"],
        headers=[h for h in headers if h],
        rows=rows,
        audit={"founded_filled": filled, "empty_founded_after": sum(1 for r in rows if _empty(r.get("Founded")))},
        title_banner=f"{m['query']} — cleaned + founded fill",
    )
    # sync to FINAL
    import shutil

    shutil.copy2(out, path)
    print(
        f"done filled=+{filled} empty_founded={sum(1 for r in rows if _empty(r.get('Founded')))} -> {path.name}",
        flush=True,
    )


if __name__ == "__main__":
    asyncio.run(main())
