"""Faster Founded-year verify for all flexible packaging FINAL rows.

Uses DDGS + strong founded-context regex. DeepSeek only for mismatch candidates.
Checkpoint every company. Optional APPLY_FOUNDED_FIXES=1 to write high-confidence fixes.
"""
from __future__ import annotations

import json
import os
import re
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
CKPT = FP.parent / "founded_verify_all_ckpt.json"
AUDIT = FP.parent / "founded_verify_all_audit.json"

API_KEY = (os.getenv("OPENAI_API_KEY") or "").strip()
BASE_URL = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()

STRONG = re.compile(
    r"(?:founded|established|est\.?|since|incorporated|gegr[uü]ndet|seit|"
    r"opgericht|fond[eée]|in business since|family[- ]owned since|"
    r"serving.{0,40}since|commenced trading|incorporation)\D{0,40}\b((?:18|19|20)\d{2})\b|"
    r"\b((?:18|19|20)\d{2})\b\D{0,20}(?:founded|established|gegr[uü]ndet|seit)",
    re.I,
)


def ddgs_text(query: str) -> str:
    try:
        from ddgs import DDGS
    except Exception as e:
        return f"[ddgs_error] {e}"
    lines = []
    try:
        with DDGS() as ddg:
            for r in ddg.text(query, max_results=5):
                lines.append(f"{r.get('title','')} | {r.get('body','')} | {r.get('href','')}")
    except Exception as e:
        return f"[ddgs_error] {e}"
    return "\n".join(lines)[:5000]


def strong_years(text: str) -> list[str]:
    out, seen = [], set()
    for m in STRONG.finditer(text or ""):
        y = m.group(1) or m.group(2)
        if not y or y in seen:
            continue
        yi = int(y)
        if 1800 <= yi <= 2026:
            seen.add(y)
            out.append(y)
    return out


def deepseek_gate(company: str, excel_year: str, evidence: str, candidates: list[str]) -> dict:
    if not API_KEY or not candidates:
        return {"year": None, "confidence": "none", "reason": "skip"}
    from openai import OpenAI

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    prompt = f"""Verify founding year. ONLY return a year from CANDIDATES or null.
Company: {company}
Excel: {excel_year}
Candidates: {candidates}
Evidence:
{evidence[:6000]}
Ignore copyright © years. JSON: {{"year":"1995"|null,"confidence":"high"|"medium"|"low","reason":"short"}}"""
    try:
        resp = client.chat.completions.create(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0,
        )
        text = (resp.choices[0].message.content or "").strip()
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.M).strip()
        data = json.loads(text)
        y = data.get("year")
        if y is not None:
            y = str(y)
            if y not in candidates:
                return {"year": None, "confidence": "low", "reason": "not_in_candidates"}
        return {"year": y, "confidence": data.get("confidence") or "low", "reason": data.get("reason") or ""}
    except Exception as e:
        return {"year": None, "confidence": "error", "reason": str(e)[:120]}


def main() -> None:
    apply_fixes = os.environ.get("APPLY_FOUNDED_FIXES", "0") == "1"
    wb = load_workbook(FP)
    ws = wb.active
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=2), start=2):
        rows.append(
            {
                "excel_row": i,
                "company": str(r[0].value or "").strip(),
                "website": str(r[1].value or "").strip(),
                "founded": str(r[2].value or "").strip(),
            }
        )

    done = {}
    if CKPT.exists():
        done = json.loads(CKPT.read_text(encoding="utf-8")).get("done") or {}

    counts = Counter()
    mismatches = []

    for idx, row in enumerate(rows, 1):
        key = row["company"]
        if key in done:
            counts[done[key].get("status", "?")] += 1
            continue

        excel_y = row["founded"]
        print(f"[{idx}/{len(rows)}] {key} excel={excel_y}", flush=True)

        if not re.fullmatch(r"\d{4}", excel_y):
            item = {**row, "status": "invalid_excel", "candidates": []}
            done[key] = item
            counts["invalid_excel"] += 1
        else:
            q = f'"{key}" founded OR established OR gegründet OR "seit" OR "est."'
            evidence = ddgs_text(q)
            # second query with domain if useful
            host = re.sub(r"^https?://(www\.)?", "", row["website"]).split("/")[0]
            if host:
                evidence += "\n" + ddgs_text(f"{key} {host} founded OR established")
            cands = strong_years(evidence)

            if excel_y in cands:
                item = {**row, "status": "confirmed", "candidates": cands, "reason": "in_evidence"}
                counts["confirmed"] += 1
            elif not cands:
                item = {**row, "status": "no_evidence", "candidates": [], "reason": "no_strong_year"}
                counts["no_evidence"] += 1
            else:
                # majority candidate
                most = Counter(cands).most_common(1)[0][0]
                if most == excel_y:
                    item = {**row, "status": "confirmed", "candidates": cands, "reason": "majority"}
                    counts["confirmed"] += 1
                else:
                    gate = deepseek_gate(key, excel_y, evidence, cands)
                    gy = gate.get("year")
                    conf = gate.get("confidence")
                    if gy == excel_y and conf in ("high", "medium"):
                        item = {**row, "status": "confirmed", "candidates": cands, "gate": gate}
                        counts["confirmed"] += 1
                    elif gy and gy != excel_y and conf == "high":
                        item = {
                            **row,
                            "status": "mismatch",
                            "candidates": cands,
                            "suggested": gy,
                            "gate": gate,
                        }
                        counts["mismatch"] += 1
                        mismatches.append(item)
                        if apply_fixes:
                            ws.cell(row["excel_row"], 3).value = gy
                            item["applied"] = True
                    else:
                        item = {
                            **row,
                            "status": "uncertain",
                            "candidates": cands,
                            "gate": gate,
                            "reason": "low_confidence",
                        }
                        counts["uncertain"] += 1

            done[key] = item

        CKPT.write_text(
            json.dumps({"updated": datetime.now().isoformat(timespec="seconds"), "done": done}, indent=2),
            encoding="utf-8",
        )
        time.sleep(0.2)

    if apply_fixes:
        wb.save(FP)

    # recount from done for final accuracy
    counts = Counter(v.get("status") for v in done.values())
    mismatches = [v for v in done.values() if v.get("status") == "mismatch"]
    audit = {
        "updated": datetime.now().isoformat(timespec="seconds"),
        "total": len(rows),
        "counts": dict(counts),
        "confirmed": counts.get("confirmed", 0),
        "mismatch": counts.get("mismatch", 0),
        "uncertain": counts.get("uncertain", 0),
        "no_evidence": counts.get("no_evidence", 0),
        "mismatches": mismatches,
        "apply_fixes": apply_fixes,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(json.dumps(audit["counts"], indent=2))
    print("MISMATCHES:")
    for m in mismatches:
        print(f" - {m['company']}: excel={m['founded']} -> {m.get('suggested')} | {m.get('candidates')}")
    print("wrote", AUDIT)


if __name__ == "__main__":
    main()
