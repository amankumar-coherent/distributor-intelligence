#!/usr/bin/env python3
"""Second pass: fill remaining NPD Key Brands via DeepSeek; fix leftover summaries."""
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
BACKUP = XLSX.with_name(XLSX.stem + "_BEFORE_DEEPSEEK_AUDIT.xlsx")
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip() or "deepseek-chat"
BASE = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
BATCH = 15
NPD = "Not publicly disclosed"

MULTI_ELEC = (
    "Multi-vendor electrical (Schneider Electric; Siemens; ABB; Eaton; Legrand "
    "& others — assortment varies / not fully public)"
)
MULTI_ECOM = (
    "Multi-vendor electronic components (TI; Analog Devices; TE Connectivity; "
    "Molex; Microchip & others — assortment varies / not fully public)"
)

# Restore wrongly dropped electronics-channel firms (same family as Digi-Key/Mouser/Arrow)
RESTORE = {
    "Sager Electronics (Asia Pacific)",
    "WPG Holdings (Asia Pacific)",
    "Würth Elektronik",
    "Excelpoint Technology",  # may already be kept
}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _client():
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not key:
        raise RuntimeError("OPENAI_API_KEY missing")
    if "deepseek" not in BASE.lower():
        raise RuntimeError(f"non-deepseek base: {BASE}")
    return OpenAI(api_key=key, base_url=BASE)


def _parse_json(text: str) -> Any:
    raw = (text or "").strip()
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.I)
    if m:
        raw = m.group(1).strip()
    i = raw.find("{")
    if i < 0:
        return {}
    raw = raw[i:]
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        j = raw.rfind("}")
        if j > 0:
            try:
                return json.loads(raw[: j + 1])
            except json.JSONDecodeError:
                return {}
        return {}


def llm_brands(client: Any, rows: list[dict[str, str]]) -> dict[str, str]:
    payload = [
        {
            "Company": r["Company"],
            "Website": r.get("Website", "")[:80],
            "Headquarters": r.get("Headquarters", "")[:60],
            "Distribution Type": r.get("Distribution Type", ""),
            "Core Categories": r.get("Core Categories", "")[:120],
            "Specialty Focus": r.get("Specialty Focus", "")[:120],
        }
        for r in rows
    ]
    system = (
        "You fill Key Brands Represented for electrical/electronics distributors. "
        "Return manufacturer brands the company sells/distributes (semicolon-separated). "
        "If it is a broadline electrical wholesaler and the exact line card is not public, "
        "return exactly: "
        f"'{MULTI_ELEC}'. "
        "If it is an electronics components catalog/VAD and line card is not public, return exactly: "
        f"'{MULTI_ECOM}'. "
        "If it is an OEM manufacturer-distributor, list its own product brand names only when known. "
        "If truly unknown conglomerate/trading with no electrical line card, return "
        "'Not publicly disclosed'. NEVER invent specific obscure local brands. "
        'JSON only: {"results":[{"Company":"...","key_brands":"..."}]}'
    )
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0.1,
        max_tokens=2500,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system},
            {
                "role": "user",
                "content": "Fill brands:\n" + json.dumps(payload, ensure_ascii=False),
            },
        ],
    )
    data = _parse_json(resp.choices[0].message.content or "{}")
    out: dict[str, str] = {}
    for it in data.get("results") or []:
        name = str(it.get("Company") or "").strip()
        br = str(it.get("key_brands") or "").strip()
        if name and br:
            out[_norm(name)] = br
    return out


def _load_sheet(path: Path):
    wb = load_workbook(path)
    ws = wb.active
    headers = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {h: i + 1 for i, h in enumerate(headers)}
    return wb, ws, headers, col


def main() -> None:
    client = _client()
    wb, ws, headers, col = _load_sheet(XLSX)
    # Fix B&Q boilerplate summary if present
    for ridx in range(2, ws.max_row + 1):
        name = str(ws.cell(ridx, col["Company"]).value or "").strip()
        if name != "B&Q":
            continue
        s = str(ws.cell(ridx, col["Summary"]).value or "")
        if "openai_web_search" in s.lower() or "channel participant" in s.lower():
            ws.cell(ridx, col["Summary"]).value = (
                "B&Q is Kingfisher's major UK DIY and home-improvement retailer with a "
                "substantial electrical, lighting and wiring assortment for consumers and trades."
            )
            print("fixed B&Q summary")

    # Restore dropped electronics channel rows from backup if missing
    present = {
        str(ws.cell(r, col["Company"]).value or "").strip()
        for r in range(2, ws.max_row + 1)
    }
    if BACKUP.exists():
        bwb = load_workbook(BACKUP, data_only=True)
        bws = bwb.active
        bh = [str(c.value or "").strip() for c in next(bws.iter_rows(min_row=1, max_row=1))]
        restored = 0
        for brow in bws.iter_rows(min_row=2, values_only=True):
            rowd = {
                bh[j]: ("" if j >= len(brow) or brow[j] is None else str(brow[j]).strip())
                for j in range(len(bh))
                if bh[j]
            }
            name = rowd.get("Company") or ""
            if name not in RESTORE or name in present:
                continue
            # append
            ws.append([rowd.get(h, "") for h in headers])
            # ensure summary/brands sensible
            ridx = ws.max_row
            if (rowd.get("Key Brands Represented") or "").lower() in {
                "",
                "not publicly disclosed",
            }:
                ws.cell(ridx, col["Key Brands Represented"]).value = (
                    "Würth Elektronik"
                    if name == "Würth Elektronik"
                    else MULTI_ECOM
                )
            if "openai_web_search" in (rowd.get("Summary") or "").lower():
                ws.cell(ridx, col["Summary"]).value = (
                    f"{name} is an electronics components distributor serving Asia Pacific "
                    "OEM and EMS customers with semiconductors, interconnect and passives."
                    if "Würth" not in name
                    else (
                        "Würth Elektronik manufactures and distributes electronic components "
                        "(EMC, power modules, passives) worldwide."
                    )
                )
            present.add(name)
            restored += 1
            print(f"restored {name}")
        print(f"restored_count={restored}")

    # Collect NPD brand rows
    targets: list[dict[str, Any]] = []
    for ridx in range(2, ws.max_row + 1):
        name = str(ws.cell(ridx, col["Company"]).value or "").strip()
        if not name:
            continue
        br = str(ws.cell(ridx, col["Key Brands Represented"]).value or "").strip()
        if br.lower() != "not publicly disclosed" and br:
            continue
        dist = str(ws.cell(ridx, col["Distribution Type"]).value or "").strip()
        # Heuristic fill for clear wholesalers without burning tokens
        if dist in {
            "Wholesaler",
            "Broadline Distributor",
            "Online Distributor",
            "Specialty Distributor",
        }:
            targets.append(
                {
                    "_row": ridx,
                    "Company": name,
                    "Website": str(ws.cell(ridx, col["Website"]).value or ""),
                    "Headquarters": str(ws.cell(ridx, col.get("Headquarters", 1)).value or "")
                    if "Headquarters" in col
                    else "",
                    "Distribution Type": dist,
                    "Core Categories": str(ws.cell(ridx, col["Core Categories"]).value or ""),
                    "Specialty Focus": str(ws.cell(ridx, col["Specialty Focus"]).value or ""),
                }
            )
        elif dist in {"Manufacturer / Distributor", "Trading Company", "Conglomerate / Group"}:
            targets.append(
                {
                    "_row": ridx,
                    "Company": name,
                    "Website": str(ws.cell(ridx, col["Website"]).value or ""),
                    "Headquarters": str(ws.cell(ridx, col["Headquarters"]).value or "")
                    if "Headquarters" in col
                    else "",
                    "Distribution Type": dist,
                    "Core Categories": str(ws.cell(ridx, col["Core Categories"]).value or ""),
                    "Specialty Focus": str(ws.cell(ridx, col["Specialty Focus"]).value or ""),
                }
            )

    print(f"npd_targets={len(targets)}")
    filled = 0
    still = 0
    changes: list[dict[str, str]] = []
    for i in range(0, len(targets), BATCH):
        chunk = targets[i : i + BATCH]
        print(
            f"  brands batch {i // BATCH + 1}: "
            + ", ".join(c["Company"][:24] for c in chunk),
            flush=True,
        )
        try:
            got = llm_brands(client, chunk)
        except Exception as exc:
            print(f"  ERROR {exc}")
            got = {}
        for r in chunk:
            name = r["Company"]
            br = got.get(_norm(name), "")
            if not br:
                # fallback for wholesalers
                if r["Distribution Type"] in {
                    "Wholesaler",
                    "Broadline Distributor",
                    "Online Distributor",
                    "Specialty Distributor",
                }:
                    br = MULTI_ELEC
                else:
                    br = NPD
            if br != NPD:
                ws.cell(r["_row"], col["Key Brands Represented"]).value = br
                filled += 1
                changes.append({"company": name, "brands": br})
            else:
                still += 1
        time.sleep(0.35)

    for cell in ws[1]:
        cell.font = Font(bold=True)
    wb.save(XLSX)
    audit = XLSX.parent / "_deepseek_brands_pass2.json"
    audit.write_text(
        json.dumps(
            {
                "when": datetime.now().isoformat(timespec="seconds"),
                "filled": filled,
                "still_npd": still,
                "changes": changes,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    print(f"filled={filled} still_npd≈{still}")
    print(f"wrote {XLSX}")
    print(f"audit {audit}")


if __name__ == "__main__":
    main()
