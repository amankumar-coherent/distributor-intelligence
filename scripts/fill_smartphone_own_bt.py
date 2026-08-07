#!/usr/bin/env python3
"""Verify/fix Ownership + Business Type on smartphone FINAL via web + DeepSeek."""
from __future__ import annotations

import asyncio
import json
import os
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import dotenv_values, load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

load_dotenv(ROOT / ".env")
_cmi = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env")
if _cmi.exists():
    for k, v in (dotenv_values(_cmi) or {}).items():
        if v and k.startswith(("OPENAI_", "LLM_", "COMPILER_", "CLASSIFIER_", "MARKET_MAP_")):
            os.environ[k] = v

FINAL = (
    ROOT
    / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL.xlsx"
)
SHEET = "Companies"
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip() or "deepseek-chat"
BASE = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
CONCURRENT = int(os.getenv("OWN_BT_CONCURRENT") or "4")
LIMIT = int(os.getenv("OWN_BT_LIMIT") or "0")

OWN_OK = {
    "public",
    "private",
    "independent",
    "subsidiary",
    "state-owned",
    "family-owned",
    "pe-backed",
    "cooperative",
}
BT_OK = {
    "distributor",
    "wholesaler",
    "retailer",
    "reseller",
    "vad",
    "var",
    "carrier",
    "dealer",
    "importer",
    "e-commerce",
    "omnichannel",
    "systems integrator",
}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


# Curated facts for well-known firms (wins unless LLM high-confidence contradicts carefully)
KNOWN: dict[str, dict[str, str]] = {
    "amazon": {"ownership": "Public", "business_type": "Retailer / E-commerce"},
    "best buy": {"ownership": "Public", "business_type": "Retailer"},
    "currys": {"ownership": "Public", "business_type": "Retailer"},
    "cdw": {"ownership": "Public", "business_type": "Reseller / Distributor"},
    "also holding": {"ownership": "Public", "business_type": "Distributor"},
    "ab sa": {"ownership": "Public", "business_type": "Distributor"},
    "ingram micro": {"ownership": "Private", "business_type": "Distributor"},  # Platinum Equity
    "td synnex": {"ownership": "Public", "business_type": "Distributor"},
    "d&h distributing": {"ownership": "Private", "business_type": "Distributor"},
    "esprinet": {"ownership": "Public", "business_type": "Distributor"},
    "computacenter": {"ownership": "Public", "business_type": "Reseller / Systems Integrator"},
    "bechtle": {"ownership": "Public", "business_type": "Reseller / Systems Integrator"},
    "dicker data": {"ownership": "Public", "business_type": "Distributor"},
    "elko group": {"ownership": "Public", "business_type": "Distributor"},
    "ban leong technologies": {"ownership": "Public", "business_type": "Distributor"},
    "brightstar": {"ownership": "Subsidiary", "business_type": "Distributor"},  # SoftBank
    "exertis": {"ownership": "Subsidiary", "business_type": "Distributor"},  # Currys / Dixons
    "exertis ireland": {"ownership": "Subsidiary", "business_type": "Distributor"},
    "shi international": {"ownership": "Private", "business_type": "Reseller / Distributor"},
    "insight enterprises": {"ownership": "Public", "business_type": "Reseller / Distributor"},
    "softcat": {"ownership": "Public", "business_type": "Reseller / Distributor"},
    "softchoice": {"ownership": "Public", "business_type": "Reseller / Distributor"},
    "zones": {"ownership": "Private", "business_type": "Reseller / Distributor"},
    "erajaya": {"ownership": "Public", "business_type": "Distributor / Retailer"},
    "reliance jio": {"ownership": "Subsidiary", "business_type": "Carrier / Retail"},
    "fido solutions inc.": {"ownership": "Subsidiary", "business_type": "Carrier / Retail"},
    "croma": {"ownership": "Subsidiary", "business_type": "Retailer"},
    "redington": {"ownership": "Public", "business_type": "Distributor"},
    "rashi peripherals": {"ownership": "Public", "business_type": "Distributor"},
    "asbis": {"ownership": "Public", "business_type": "Distributor"},
    "westcoast": {"ownership": "Private", "business_type": "Distributor"},
    "siewert & kau": {"ownership": "Private", "business_type": "Distributor"},
    "3g mobile": {"ownership": "Private", "business_type": "Distributor"},
}


def _load_rows(path: Path) -> tuple[list[str], list[dict[str, str]], int]:
    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = None
    headers: list[str] = []
    for i, r in enumerate(rows[:20]):
        vals = [str(c or "").strip() for c in r]
        if "Company" in vals:
            header_idx = i
            headers = vals
            break
    if header_idx is None:
        raise RuntimeError("No Company header")
    data: list[dict[str, str]] = []
    for r in rows[header_idx + 1 :]:
        d = {
            headers[j]: ("" if j >= len(r) or r[j] is None else str(r[j]).strip())
            for j in range(len(headers))
            if headers[j]
        }
        if d.get("Company"):
            data.append(d)
    return [h for h in headers if h], data, header_idx


def _needs(row: dict[str, str]) -> bool:
    o = _norm(row.get("Ownership") or "")
    b = _norm(row.get("Business Type") or "")
    # All Independent is a mass default — re-verify everyone
    if o in ("", "independent", "seed_file", "unknown", "n/a"):
        return True
    if not b or b in ("seed_file", "unknown", "company", "other"):
        return True
    return True  # verify all


async def _web_evidence(name: str, website: str) -> str:
    host = ""
    if website:
        host = re.sub(r"^https?://", "", website.strip(), flags=re.I).split("/")[0]
        host = host.replace("www.", "")
    queries = [
        f'"{name}" distributor OR retailer OR wholesaler ownership OR "publicly traded" OR subsidiary OR private',
        f'"{name}" "business type" OR wholesaler OR "authorized distributor" OR reseller OR carrier',
    ]
    if host:
        queries.insert(0, f'"{name}" {host} about company ownership OR listed OR subsidiary')
        queries.insert(1, f"site:{host} about OR company OR investors OR ownership")

    parts: list[str] = []
    try:
        from ddgs import DDGS

        def _one(q: str) -> list[str]:
            lines: list[str] = []
            try:
                hits = list(DDGS().text(q, max_results=5)) or []
            except Exception:
                hits = []
            for h in hits:
                lines.append(
                    f"- {h.get('title', '')}: {str(h.get('body', ''))[:220]} ({h.get('href', '')})"
                )
            return lines

        for q in queries[:3]:
            try:
                lines = await asyncio.wait_for(asyncio.to_thread(_one, q), timeout=22)
            except Exception:
                lines = []
            if lines:
                parts.append(f"Query: {q}\n" + "\n".join(lines))
    except Exception as err:
        parts.append(f"(ddgs: {err})")

    if not any(p.startswith("Query:") for p in parts):
        try:
            from vendor_intel.clients.search_router import FreeSearchRouter
            from vendor_intel.config import Settings
            from vendor_intel.enrichment.gap_fill.sources import _search_snippets

            router = FreeSearchRouter(Settings.load())
            for q in queries[:2]:
                try:
                    snip = await asyncio.wait_for(
                        _search_snippets(router, q, max_hits=5), timeout=25
                    )
                except Exception:
                    snip = ""
                if snip:
                    parts.append(f"Query: {q}\n{snip}")
        except Exception as err:
            parts.append(f"(router: {err})")
    return "\n\n".join(parts)[:6500] or "(no web hits)"


def _client():
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not key:
        raise RuntimeError("OPENAI_API_KEY missing")
    return OpenAI(api_key=key, base_url=BASE)


def _normalize_ownership(v: str) -> str:
    s = _norm(v)
    if not s:
        return ""
    if "public" in s or "listed" in s or "stock exchange" in s:
        return "Public"
    if "state" in s or "government" in s:
        return "State-owned"
    if "subsidiar" in s or "owned by" in s or "parent" in s:
        return "Subsidiary"
    if "private equity" in s or s in ("pe-backed", "pe backed") or "pe-backed" in s:
        return "PE-backed"
    if "family" in s:
        return "Family-owned"
    if "cooperative" in s or "co-op" in s:
        return "Cooperative"
    if "private" in s:
        return "Private"
    if "independent" in s:
        return "Independent"
    # title-case first token set
    for cand in ("Public", "Private", "Subsidiary", "Independent", "Family-owned", "PE-backed", "State-owned"):
        if s == _norm(cand):
            return cand
    return v.strip()[:40]


def _normalize_bt(v: str) -> str:
    s = (v or "").strip()
    if not s:
        return ""
    sl = s.lower()
    # collapse noisy variants
    mapping = [
        (r"telecom operator.*retail", "Carrier / Retail"),
        (r"carrier", "Carrier / Retail"),
        (r"e-?commerce.*retail|online retailer|retailer / e-commerce", "Retailer / E-commerce"),
        (r"retailer.*distributor|distributor.*retailer", "Distributor / Retailer"),
        (r"reseller.*distributor|distributor.*reseller|vad|value.?added", "Reseller / Distributor"),
        (r"systems?\s*integrator|integrator", "Reseller / Systems Integrator"),
        (r"wholesaler.*distributor|distributor.*wholesaler|wholesaler", "Distributor / Wholesaler"),
        (r"\bretailer\b", "Retailer"),
        (r"\bdealer\b", "Dealer"),
        (r"\bimporter\b", "Importer / Distributor"),
        (r"\bdistributor\b", "Distributor"),
        (r"\breseller\b", "Reseller"),
    ]
    for pat, label in mapping:
        if re.search(pat, sl):
            return label
    return s[:60]


def _llm(client: Any, row: dict[str, str], evidence: str) -> dict[str, Any]:
    name = row.get("Company") or ""
    system = (
        "Classify ownership and channel business type for a smartphone market landscape. "
        "Return JSON only. Do not invent. Prefer evidence.\n"
        "Ownership allowed: Public | Private | Subsidiary | Independent | Family-owned | "
        "PE-backed | State-owned | Cooperative.\n"
        "Use Public if listed on a stock exchange. Subsidiary if owned by a parent group. "
        "Private for privately held. Independent only if clearly independent private without "
        "known parent (prefer Private over Independent when unsure private).\n"
        "Business Type: Distributor | Wholesaler | Retailer | Reseller / Distributor | "
        "Distributor / Retailer | Carrier / Retail | Retailer / E-commerce | "
        "Reseller / Systems Integrator | Dealer | Importer / Distributor | "
        "Distributor / Wholesaler.\n"
        "Pick the PRIMARY role in the smartphone channel."
    )
    user = f"""Company: {name}
Website: {row.get('Website') or ''}
Current Ownership: {row.get('Ownership') or ''}
Current Business Type: {row.get('Business Type') or ''}
Specialty: {row.get('Specialty Focus') or ''}
Summary: {(row.get('Summary') or '')[:280]}

WEB EVIDENCE:
{evidence}

Return JSON:
{{
  "ownership": "Public|Private|Subsidiary|Independent|Family-owned|PE-backed|State-owned|Cooperative",
  "business_type": "one allowed business type",
  "confidence": "high|medium|low",
  "notes": "short"
}}
"""
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0.1,
        max_tokens=350,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    raw = resp.choices[0].message.content or "{}"
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", raw)
        data = json.loads(m.group(0)) if m else {}
    return data if isinstance(data, dict) else {}


def _apply(row: dict[str, str], data: dict[str, Any]) -> dict[str, str]:
    out = dict(row)
    known = KNOWN.get(_norm(row.get("Company") or ""), {})
    conf = str(data.get("confidence") or "").lower()

    own = _normalize_ownership(str(data.get("ownership") or ""))
    bt = _normalize_bt(str(data.get("business_type") or ""))

    if known.get("ownership") and (conf != "high" or not own):
        own = known["ownership"]
    elif known.get("ownership") and conf == "high":
        # still prefer curated for ownership of majors
        own = known["ownership"]

    if known.get("business_type") and (conf != "high" or not bt):
        bt = known["business_type"]
    elif known.get("business_type") and conf == "high":
        # if LLM agrees category-ish keep known canonical
        bt = known["business_type"]

    if not own:
        own = "Private" if _norm(out.get("Ownership") or "") in ("independent", "") else (
            out.get("Ownership") or "Private"
        )
    # mass Independent default → Private when still Independent after weak evidence
    if _norm(own) == "independent" and conf in ("", "low", "medium"):
        own = "Private"

    if not bt:
        bt = out.get("Business Type") or "Distributor"

    out["Ownership"] = own
    out["Business Type"] = bt
    src = out.get("Data Sources") or ""
    if "deepseek_own_bt" not in src:
        out["Data Sources"] = (src + "; deepseek_own_bt").strip("; ")
    return out


def _write(path: Path, headers: list[str], rows: list[dict[str, str]], title: str, audit: dict) -> None:
    wb = load_workbook(path)
    ws = wb[SHEET]
    header_row = 2
    for r in range(1, 6):
        if ws.cell(r, 1).value == "Company":
            header_row = r
            break
    if ws.max_row > header_row:
        ws.delete_rows(header_row + 1, ws.max_row - header_row)
    for j, h in enumerate(headers, 1):
        ws.cell(header_row, j).value = h
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    for i, row in enumerate(rows):
        rr = header_row + 1 + i
        for j, h in enumerate(headers, 1):
            cell = ws.cell(rr, j, row.get(h, ""))
            cell.alignment = wrap
            cell.border = thin
    ws.cell(1, 1).value = title

    name = "OwnBT Audit"
    if name in wb.sheetnames:
        del wb[name]
    aus = wb.create_sheet(name)
    aus.append(["field", "value"])
    for k, v in audit.items():
        if k == "samples":
            continue
        aus.append([k, str(v)])
    wb.save(path)


async def _one(sem, client, row, idx, total):
    name = row.get("Company") or ""
    async with sem:
        print(f"  [{idx}/{total}] {name}", flush=True)
        try:
            evidence = await _web_evidence(name, row.get("Website") or "")
            data = await asyncio.to_thread(_llm, client, row, evidence)
            fixed = _apply(row, data)
            print(
                f"      → Own={fixed.get('Ownership')} | BT={fixed.get('Business Type')} "
                f"(conf={data.get('confidence')})",
                flush=True,
            )
            return fixed, {
                "company": name,
                "ownership": fixed.get("Ownership"),
                "business_type": fixed.get("Business Type"),
                "confidence": data.get("confidence"),
                "notes": data.get("notes"),
            }
        except Exception as err:  # noqa: BLE001
            print(f"      ERROR {type(err).__name__}: {err}", flush=True)
            fixed = _apply(row, {})
            return fixed, {"company": name, "error": str(err)}


async def main() -> int:
    if not FINAL.exists():
        print(f"ERROR missing {FINAL}", file=sys.stderr)
        return 1
    headers, rows, _ = _load_rows(FINAL)
    todo = list(range(len(rows)))
    if LIMIT > 0:
        todo = todo[:LIMIT]

    print("=== Smartphone Ownership + Business Type verify ===", flush=True)
    print(f"  file : {FINAL}", flush=True)
    print(f"  rows : {len(todo)} / {len(rows)}", flush=True)
    print(f"  model: {MODEL} @ {BASE}", flush=True)

    client = _client()
    ping = client.chat.completions.create(
        model=MODEL,
        max_tokens=20,
        messages=[{"role": "user", "content": 'Reply JSON {"ok":true}'}],
        response_format={"type": "json_object"},
    )
    print(f"  ping: {(ping.choices[0].message.content or '')[:40]}", flush=True)

    sem = asyncio.Semaphore(CONCURRENT)
    out_rows = list(rows)
    audits: list[dict[str, Any]] = []
    ckpt_path = FINAL.parent / "smartphone_own_bt_checkpoint.json"
    done_map: dict[str, dict[str, Any]] = {}
    if ckpt_path.exists() and not (os.getenv("OWN_BT_FRESH") or "").strip():
        try:
            prior = json.loads(ckpt_path.read_text(encoding="utf-8"))
            for v in prior.get("audits") or []:
                if v.get("company") and "error" not in v:
                    done_map[_norm(v["company"])] = v
            print(f"  resume: {len(done_map)} prior", flush=True)
        except Exception:
            pass

    # apply prior if any
    for i, row in enumerate(rows):
        a = done_map.get(_norm(row.get("Company") or ""))
        if a and a.get("ownership") and a.get("business_type"):
            fixed = dict(row)
            fixed["Ownership"] = a["ownership"]
            fixed["Business Type"] = a["business_type"]
            out_rows[i] = fixed

    pending = [i for i in todo if _norm(rows[i].get("Company") or "") not in done_map]
    chunk = 12
    for start in range(0, len(pending), chunk):
        batch = pending[start : start + chunk]
        print(f"--- batch {start // chunk + 1}: {len(batch)} (left ~{len(pending)-start}) ---", flush=True)
        tasks = [
            _one(sem, client, rows[i], start + j + 1 + len(done_map), len(todo))
            for j, i in enumerate(batch)
        ]
        results = await asyncio.gather(*tasks)
        for i, (fixed, audit) in zip(batch, results):
            out_rows[i] = fixed
            audits.append(audit)
            done_map[_norm(fixed.get("Company") or "")] = audit
        ckpt_path.write_text(
            json.dumps({"audits": list(done_map.values()), "done": len(done_map)}, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )

    # rebuild audits in order
    audits = [
        done_map.get(
            _norm(r.get("Company") or ""),
            {"company": r.get("Company"), "ownership": r.get("Ownership"), "business_type": r.get("Business Type")},
        )
        for r in out_rows
    ]

    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    if LIMIT > 0:
        p = FINAL.parent / f"smartphone_own_bt_DRY_{ts}.json"
        p.write_text(json.dumps(audits, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"DRY wrote {p}", flush=True)
        from collections import Counter
        print("Own", Counter(a.get("ownership") for a in audits))
        print("BT", Counter(a.get("business_type") for a in audits))
        return 0

    bak = FINAL.with_suffix(FINAL.suffix + f".bak_own_bt_{ts}")
    shutil.copy2(FINAL, bak)
    audit_path = FINAL.parent / f"smartphone_own_bt_{ts}.json"
    from collections import Counter

    own_c = Counter(r.get("Ownership") for r in out_rows)
    bt_c = Counter(r.get("Business Type") for r in out_rows)
    payload = {
        "ts": ts,
        "model": MODEL,
        "rows": len(out_rows),
        "ownership_counts": dict(own_c),
        "business_type_counts": dict(bt_c),
        "audits": audits,
    }
    audit_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    title = f"Global Smartphone Market — channel partners (ownership/BT verified {ts})"
    _write(
        FINAL,
        headers,
        out_rows,
        title,
        {
            "ts": ts,
            "model": MODEL,
            "rows": len(out_rows),
            "ownership_counts": dict(own_c),
            "business_type_counts": dict(bt_c),
            "audit_json": str(audit_path),
            "backup": str(bak),
        },
    )
    print("=== DONE ===", flush=True)
    print(f"  backup: {bak.name}", flush=True)
    print(f"  audit : {audit_path}", flush=True)
    print("  Ownership:", dict(own_c), flush=True)
    print("  Business Type:", dict(bt_c.most_common(12)), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
