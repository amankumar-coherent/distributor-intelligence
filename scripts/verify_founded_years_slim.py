#!/usr/bin/env python3
"""Verify suspicious Founded years in FINAL_SLIM Excels via web search + DeepSeek.

Does not modify original FINAL files. Updates *_FINAL_SLIM.xlsx in place.
"""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook
from openai import OpenAI

load_dotenv(Path(r"D:\vendor-intel\.env"))

SUSPECTS_PATH = Path(r"D:\vendor-intel\scripts\_founded_suspects.json")
FIXES_PATH = Path(r"D:\vendor-intel\scripts\_founded_fixes.json")
CURRENT_YEAR = 2026

# High-confidence fixes already verified via Google / company sites
KNOWN_FIXES: dict[str, str] = {
    "switch": "1995",  # CG Computers / Switch Malaysia since 1995
    "machines": "2006",  # Machines Sdn Bhd Malaysia — already correct, confirm
    "rptechindia": "1989",  # Rashi Peripherals / RP tech India
    "bkm distribuidora": "2002",  # CNPJ abertura 16/05/2002
    "asar general trading fzc": "NPD",  # ©2026 copyright contamination; year not published
    "distribuidora de celulares brasil": "NPD",  # 2026 invalid; no reliable founding year found
    "canadian bargain": "NPD",  # 2026 invalid; site has no published founding year
}


def norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def extract_year(text: str) -> int | None:
    m = re.search(r"(18|19|20)\d{2}", str(text or ""))
    return int(m.group(0)) if m else None


def ddg_snippets(query: str, max_results: int = 5) -> str:
    try:
        from ddgs import DDGS
    except ImportError:
        from duckduckgo_search import DDGS  # type: ignore
    bits: list[str] = []
    try:
        with DDGS() as ddgs:
            for r in ddgs.text(query, max_results=max_results):
                title = r.get("title") or ""
                body = r.get("body") or r.get("description") or ""
                href = r.get("href") or r.get("link") or ""
                bits.append(f"- {title}\n  {body}\n  {href}")
    except Exception as e:
        bits.append(f"(search error: {e})")
    return "\n".join(bits) if bits else "(no results)"


def deepseek_year(company: str, hq: str, website: str, current: str, snippets: str) -> dict:
    api_key = os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY")
    base = os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1"
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    client = OpenAI(api_key=api_key, base_url=base)
    prompt = f"""You verify company founding years from web search snippets.

Company: {company}
HQ: {hq}
Website: {website}
Current Excel founded value: {current}

Web search snippets:
{snippets}

Rules:
- Prefer explicit "founded/established/incorporated/since YEAR" statements.
- Do NOT use copyright footer years (e.g. © 2026) as founding year.
- Do NOT invent a year. If unclear, return null.
- If current value is {CURRENT_YEAR} and only copyright shows that year, treat as wrong.
- Return JSON only: {{"year": 1995 or null, "confidence": "high"|"medium"|"low", "reason": "short"}}
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )
    text = (resp.choices[0].message.content or "").strip()
    # strip fences
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.M).strip()
    try:
        return json.loads(text)
    except Exception:
        y = extract_year(text)
        return {"year": y, "confidence": "low", "reason": text[:200]}


def verify_one(item: dict) -> dict:
    company = item["company"]
    key = norm_name(company)
    out = {
        "file": item["file"],
        "company": company,
        "old": item["founded"],
        "new": None,
        "action": "keep",
        "source": "",
        "reason": "",
    }
    if key in KNOWN_FIXES:
        new = KNOWN_FIXES[key]
        if str(new) != str(item["founded"]).strip():
            out["new"] = new
            out["action"] = "fix"
            out["source"] = "known_web_verified"
            out["reason"] = "pre-verified via Google / company registry / about page"
        else:
            out["action"] = "keep"
            out["source"] = "known_web_verified"
            out["reason"] = "already correct"
        return out

    y = item.get("year")
    # Always re-check Malaysia + year>=2020
    query = f"{company} {item.get('hq','')} founded OR established OR incorporated year"
    if item.get("website"):
        query += f" {item['website']}"
    snippets = ddg_snippets(query)
    time.sleep(0.8)
    try:
        parsed = deepseek_year(company, item.get("hq", ""), item.get("website", ""), item["founded"], snippets)
    except Exception as e:
        parsed = {"year": None, "confidence": "low", "reason": f"llm error: {e}"}

    new_year = parsed.get("year")
    conf = parsed.get("confidence") or "low"
    reason = parsed.get("reason") or ""

    # Guardrails
    if isinstance(new_year, str) and new_year.isdigit():
        new_year = int(new_year)
    if not isinstance(new_year, int):
        new_year = None
    if new_year is not None and (new_year < 1600 or new_year > CURRENT_YEAR):
        new_year = None

    if y == CURRENT_YEAR and new_year is None:
        out["new"] = "NPD"
        out["action"] = "fix"
        out["source"] = "web+deepseek"
        out["reason"] = f"2026 invalid / copyright contamination; {reason}"
        return out

    if new_year is not None and new_year != y and conf in {"high", "medium"}:
        out["new"] = str(new_year)
        out["action"] = "fix"
        out["source"] = "web+deepseek"
        out["reason"] = reason
        return out

    if new_year is not None and new_year == y:
        out["action"] = "keep"
        out["source"] = "web+deepseek"
        out["reason"] = reason or "confirmed"
        return out

    out["action"] = "keep"
    out["source"] = "web+deepseek"
    out["reason"] = reason or "insufficient evidence to change"
    return out


def apply_fixes(fixes: list[dict]) -> None:
    by_file: dict[str, dict[str, str]] = {}
    for fx in fixes:
        if fx.get("action") != "fix" or not fx.get("new"):
            continue
        by_file.setdefault(fx["file"], {})[norm_name(fx["company"])] = str(fx["new"])

    for fpath, mapping in by_file.items():
        wb = load_workbook(fpath)
        ws = wb.active
        hdr = None
        hr = None
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
            key = norm_name(str(company))
            if key in mapping:
                ws.cell(r, fi + 1).value = mapping[key]
                n += 1
        wb.save(fpath)
        print(f"Applied {n} founded fixes -> {Path(fpath).name}")


def main() -> None:
    suspects = json.loads(SUSPECTS_PATH.read_text(encoding="utf-8"))
    print(f"Verifying {len(suspects)} suspect founded years...\n")
    fixes: list[dict] = []
    for i, item in enumerate(suspects, 1):
        print(f"[{i}/{len(suspects)}] {item['company']} ({item['founded']}) ...", flush=True)
        fx = verify_one(item)
        fixes.append(fx)
        print(f"    -> {fx['action']}: {fx['old']} => {fx.get('new') or fx['old']} ({fx['source']}) {fx['reason'][:80]}")
    FIXES_PATH.write_text(json.dumps(fixes, indent=2), encoding="utf-8")
    changed = [f for f in fixes if f["action"] == "fix"]
    print(f"\nFixes to apply: {len(changed)}")
    apply_fixes(fixes)
    print(f"Wrote {FIXES_PATH}")


if __name__ == "__main__":
    main()
