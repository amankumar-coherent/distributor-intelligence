#!/usr/bin/env python3
"""Audit electrical FINAL via DeepSeek only: relevance, Summary, Key Brands.

- DROP companies not primarily electrical distribution / wholesale / channel
- Rewrite Summary for every kept row
- Fill Key Brands when NPD (or leave NPD if unknown — never invent)
Uses OPENAI_* from .env pointed at api.deepseek.com (deepseek-chat).
"""
from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import dotenv_values, load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Font

load_dotenv(ROOT / ".env")
_cmi = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env")
if _cmi.exists():
    for k, v in (dotenv_values(_cmi) or {}).items():
        if v and k.startswith(("OPENAI_", "LLM_", "COMPILER_", "CLASSIFIER_", "MARKET_MAP_")):
            os.environ[k] = v

XLSX = Path(
    r"D:\vendor-intel\output\web_expand\global_electrical_distribution_market_global"
    r"\global_electrical_distribution_market_global_FINAL.xlsx"
)
MARKET = "Global Electrical Distribution Market"
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip() or "deepseek-chat"
BASE = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
BATCH = int(os.getenv("ELEC_DS_BATCH") or "12")
SLEEP = float(os.getenv("ELEC_DS_SLEEP") or "0.4")
LIMIT = int(os.getenv("ELEC_DS_LIMIT") or "0")  # 0 = all
DRY = os.getenv("ELEC_DS_DRY", "").strip() in {"1", "true", "yes"}

NPD = "Not publicly disclosed"
NPD_SET = {
    "",
    "n/a",
    "na",
    "npd",
    "unknown",
    "not publicly disclosed",
    "-",
    "—",
}

# Clear off-market / non-electrical-channel (deterministic)
HARD_DROP: dict[str, str] = {
    "Schmidt + Bender GmbH & Co. KG": "riflescope OEM — not electrical distribution",
    "Mantrac (Egypt)": "Caterpillar heavy equipment dealer — not electrical distribution",
    "Grupo Bafar": "food / protein conglomerate — not electrical distribution",
    "HSS Hire": "tool/equipment hire — not electrical product distribution",
    "Bunzl": "cleaning/safety/packaging distribution — not electrical channel",
    "Bunzl Australia": "cleaning/safety/packaging — not electrical channel",
    "Ferguson Enterprises": "plumbing/PVF primary — not electrical distribution market",
    "Wolseley UK": "plumbing/heating merchant — not electrical distribution market",
    "Wolseley Canada": "plumbing/HVAC merchant — not electrical distribution market",
    "Plumb Center (Wolseley UK)": "plumbing merchant brand — not electrical distribution",
    "City Plumbing": "plumbing merchant — not electrical distribution",
    "Al-Futtaim Real Estate": "real estate — not electrical distribution",
    "Al-Futtaim Construction": "construction — not electrical distribution",
    "Al-Futtaim Infrastructure": "infrastructure — not electrical distribution",
    "Al-Futtaim Import": "general import arm — not electrical channel primary",
    "Al-Futtaim Engineering": "engineering services — not electrical wholesale primary",
    "Maritime Electric": "electric utility — not electrical product distributor",
    "Kalpataru Power Transmission Ltd.": "EPC / transmission contractor — OEM/project, not channel",
    "City Plumbing": "plumbing merchant — not electrical distribution",
}

# Always keep (known electrical channel) — still rewrite summary/brands via LLM
KNOWN_KEEP = {
    "Rexel",
    "Sonepar",
    "WESCO International",
    "Anixter (now WESCO)",
    "Graybar Electric",
    "Graybar Canada",
    "W.W. Grainger",
    "Acklands-Grainger",
    "Fastenal",
    "Digi-Key Electronics",
    "Digi-Key Electronics (Asia Pacific)",
    "Mouser Electronics",
    "Mouser Electronics (Asia Pacific)",
    "City Electrical Factors (CEF)",
    "City Electric Supply (CES)",
    "Consolidated Electrical Distributors (CED)",
    "Edmundson Electrical",
    "Border States Electric",
    "Crescent Electric Supply Company",
    "Elliott Electric Supply",
    "Mayer Electric Supply",
    "Gexpro",
    "Platt Electric Supply",
    "RS Components (now RS Group)",
    "RS Components (Asia Pacific)",
    "Allied Electronics & Automation",
    "Farnell (Newark)",
    "Newark",
    "Element14",
    "Arrow Electronics (Asia Pacific)",
    "Avnet (Asia Pacific)",
    "Voltex",
    "Middy's",
    "L&H Group",
    "Blackwoods",
    "Bunnings Group",
    "B&Q",
    "Screwfix",
}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _client():
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not key:
        raise RuntimeError("OPENAI_API_KEY missing (DeepSeek)")
    if "deepseek" not in BASE.lower():
        raise RuntimeError(f"Refusing non-DeepSeek base URL: {BASE}")
    return OpenAI(api_key=key, base_url=BASE)


def _parse_json(text: str) -> Any:
    raw = (text or "").strip()
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.I)
    if m:
        raw = m.group(1).strip()
    start = raw.find("{")
    if start == -1:
        return {}
    raw = raw[start:]
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # trim trailing junk
        end = raw.rfind("}")
        if end != -1:
            try:
                return json.loads(raw[: end + 1])
            except json.JSONDecodeError:
                return {}
        return {}


def _is_npd(val: str) -> bool:
    return (val or "").strip().lower() in NPD_SET


def _clean_brands(val: str) -> str:
    s = (val or "").strip()
    if _is_npd(s):
        return NPD
    if re.search(r"need to verify|invent|unknown brand|n/?a", s, re.I):
        return NPD
    # normalize commas to semicolons
    if ";" not in s and "," in s:
        parts = [p.strip() for p in s.split(",") if p.strip()]
        s = "; ".join(parts)
    s = re.sub(r"\s+", " ", s).strip(" .;")
    return s or NPD


def _clean_summary(val: str, company: str) -> str:
    s = re.sub(r"\s+", " ", (val or "").strip())
    s = re.sub(
        r"^" + re.escape(company) + r"\s*[—\-–:]\s*channel participant.*?(?:Verify:)?\s*",
        "",
        s,
        flags=re.I,
    )
    if len(s) > 420:
        s = s[:417].rstrip() + "…"
    return s


def llm_batch(client: Any, rows: list[dict[str, str]]) -> dict[str, dict[str, Any]]:
    """Classify + summary + brands for a batch. Keyed by company name."""
    payload = []
    for r in rows:
        payload.append(
            {
                "Company": r.get("Company") or "",
                "Website": (r.get("Website") or "")[:100],
                "Headquarters": (r.get("Headquarters") or "")[:80],
                "Ownership": (r.get("Ownership") or "")[:80],
                "Core Categories": (r.get("Core Categories") or "")[:140],
                "Specialty Focus": (r.get("Specialty Focus") or "")[:140],
                "Distribution Type": (r.get("Distribution Type") or "")[:60],
                "Key Brands Represented": (r.get("Key Brands Represented") or "")[:160],
                "Summary": (r.get("Summary") or "")[:200],
            }
        )
    system = (
        "You are a strict vendor-intel analyst for the Global Electrical Distribution Market. "
        "KEEP only if PRIMARY business is: electrical wholesaler/distributor, industrial MRO "
        "distributor that materially sells electrical, electronic-components catalog distributor, "
        "authorized electrical dealer, trade electrical retailer (e.g. CEF/Screwfix/Bunnings with "
        "electrical), or trading company with a real electrical distribution arm. "
        "DROP if PRIMARY is: riflescopes/optics, heavy equipment (Caterpillar), food conglomerate, "
        "pure plumbing/HVAC merchant, pure DIY big-box without pro electrical wholesale role, "
        "tool hire, electric utility/IPP, EPC/transmission contractor, pure OEM manufacturer of "
        "transformers/motors/cables with no distribution channel role, real estate, unrelated retail, "
        "or unclear/unrelated. "
        "For KEEP rows: write a fresh factual Summary (2 sentences, no boilerplate like "
        "'channel participant' or 'openai_web_search'). "
        "For Key Brands Represented: list manufacturer brands THIS company distributes/sells "
        "(semicolon-separated). If unknown, use exactly 'Not publicly disclosed'. "
        "Never invent brands. Own OEM product names only if the company is a manufacturer-distributor "
        "selling its own lines. "
        "Return ONLY JSON: "
        '{"results":[{"Company":"...","keep":true|false,"reason":"short",'
        '"summary":"...","key_brands":"..."}]}'
    )
    user = (
        f"MARKET: {MARKET}\n"
        f"Classify and fill each company.\n"
        f"Companies:\n{json.dumps(payload, ensure_ascii=False)}"
    )
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0.15,
        max_tokens=4000,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    data = _parse_json(resp.choices[0].message.content or "{}")
    out: dict[str, dict[str, Any]] = {}
    for item in data.get("results") or []:
        if not isinstance(item, dict):
            continue
        name = str(item.get("Company") or "").strip()
        if name:
            out[_norm(name)] = item
    return out


def main() -> None:
    client = _client()
    # ping
    ping = client.chat.completions.create(
        model=MODEL,
        temperature=0,
        max_tokens=20,
        messages=[{"role": "user", "content": "Reply with OK"}],
    )
    print(f"deepseek ping: {(ping.choices[0].message.content or '')[:40]}", flush=True)
    print(f"base={BASE} model={MODEL}", flush=True)

    backup = XLSX.with_name(XLSX.stem + "_BEFORE_DEEPSEEK_AUDIT.xlsx")
    if not backup.exists():
        backup.write_bytes(XLSX.read_bytes())
        print(f"backup {backup}", flush=True)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {h: i + 1 for i, h in enumerate(headers)}
    need = [
        "Company",
        "Website",
        "Headquarters",
        "Ownership",
        "Core Categories",
        "Specialty Focus",
        "Distribution Type",
        "Key Brands Represented",
        "Summary",
    ]
    for h in need:
        if h not in col:
            raise RuntimeError(f"missing column {h}")

    rows: list[dict[str, Any]] = []
    for ridx in range(2, ws.max_row + 1):
        name = str(ws.cell(ridx, col["Company"]).value or "").strip()
        if not name:
            continue
        d = {"_row": ridx, "Company": name}
        for h in need[1:]:
            d[h] = str(ws.cell(ridx, col[h]).value or "").strip()
        rows.append(d)

    if LIMIT > 0:
        rows = rows[:LIMIT]
    print(f"companies={len(rows)} batch={BATCH}", flush=True)

    decisions: list[dict[str, Any]] = []
    keep_rows: list[dict[str, Any]] = []

    # Phase 0: hard drop / known keep flags
    to_llm: list[dict[str, Any]] = []
    for r in rows:
        name = r["Company"]
        if name in HARD_DROP:
            decisions.append(
                {
                    "company": name,
                    "action": "remove",
                    "reason": HARD_DROP[name],
                    "source": "hard_drop",
                }
            )
            continue
        if name in KNOWN_KEEP:
            r["_force_keep"] = True
        to_llm.append(r)

    # Phase 1: DeepSeek batches
    for i in range(0, len(to_llm), BATCH):
        chunk = to_llm[i : i + BATCH]
        names = ", ".join(c["Company"][:28] for c in chunk)
        print(f"  batch {i // BATCH + 1}/{(len(to_llm) + BATCH - 1) // BATCH}: {names}", flush=True)
        try:
            result = llm_batch(client, chunk)
            batch_failed = False
        except Exception as exc:
            print(f"    ERROR: {type(exc).__name__}: {exc}", flush=True)
            result = {}
            batch_failed = True
        for r in chunk:
            name = r["Company"]
            hit = result.get(_norm(name)) or {}
            force = bool(r.get("_force_keep"))
            if batch_failed:
                keep = True
                reason = "batch_error_fail_open"
            elif not hit and not force:
                keep = False
                reason = "no_llm_result_strict_drop"
            else:
                keep = force or bool(hit.get("keep"))
                reason = str(hit.get("reason") or ("known_keep" if force else ""))
            if force:
                keep = True
                if not reason:
                    reason = "known_keep"
            summary = _clean_summary(str(hit.get("summary") or ""), name)
            brands_in = str(hit.get("key_brands") or "")
            brands = _clean_brands(brands_in)

            decisions.append(
                {
                    "company": name,
                    "action": "keep" if keep else "remove",
                    "reason": reason,
                    "source": "deepseek" if hit else ("known_keep" if force else "missing"),
                    "summary": summary,
                    "key_brands": brands,
                    "had_npd_brands": _is_npd(r.get("Key Brands Represented") or ""),
                }
            )
            if keep:
                r["_new_summary"] = summary
                r["_new_brands"] = brands
                keep_rows.append(r)
        time.sleep(SLEEP)

    removed = [d for d in decisions if d["action"] == "remove"]
    kept = [d for d in decisions if d["action"] == "keep"]
    print(f"keep={len(kept)} remove={len(removed)}", flush=True)

    if DRY:
        audit = XLSX.parent / "_deepseek_audit_DRY.json"
        audit.write_text(
            json.dumps(
                {"when": datetime.now().isoformat(timespec="seconds"), "decisions": decisions},
                indent=2,
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        print(f"DRY — wrote {audit}")
        return

    # Apply: update kept rows; delete removed rows (bottom-up)
    brands_filled = 0
    summaries_written = 0
    for r in keep_rows:
        ridx = r["_row"]
        new_sum = r.get("_new_summary") or ""
        if new_sum:
            ws.cell(ridx, col["Summary"]).value = new_sum
            summaries_written += 1
        elif r.get("Summary"):
            # strip boilerplate even if LLM empty
            cleaned = _clean_summary(r["Summary"], r["Company"])
            if cleaned and cleaned != r["Summary"]:
                ws.cell(ridx, col["Summary"]).value = cleaned
                summaries_written += 1

        new_br = r.get("_new_brands") or ""
        old_br = r.get("Key Brands Represented") or ""
        if new_br and new_br != NPD:
            if _is_npd(old_br) or old_br != new_br:
                ws.cell(ridx, col["Key Brands Represented"]).value = new_br
                if _is_npd(old_br):
                    brands_filled += 1
        # if still NPD and LLM returned NPD, leave NPD

    remove_row_idxs = sorted(
        {r["_row"] for r in rows if any(d["company"] == r["Company"] and d["action"] == "remove" for d in decisions)},
        reverse=True,
    )
    for ridx in remove_row_idxs:
        ws.delete_rows(ridx, 1)

    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(XLSX)

    audit = XLSX.parent / "_deepseek_audit_fill.json"
    audit.write_text(
        json.dumps(
            {
                "when": datetime.now().isoformat(timespec="seconds"),
                "model": MODEL,
                "base": BASE,
                "kept": len(kept),
                "removed": len(removed),
                "summaries_written": summaries_written,
                "brands_filled_from_npd": brands_filled,
                "removed_companies": removed,
                "decisions": decisions,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    print(f"summaries_written={summaries_written} brands_filled_from_npd={brands_filled}")
    print(f"removed {len(remove_row_idxs)} rows")
    for d in removed[:30]:
        print(f"  DROP {d['company']}: {d['reason'][:90]}")
    if len(removed) > 30:
        print(f"  ... +{len(removed) - 30} more")
    print(f"wrote {XLSX}")
    print(f"audit {audit}")


if __name__ == "__main__":
    main()
