#!/usr/bin/env python3
"""DeepSeek + heuristics: verify wearable medical devices rows are channel partners."""
from __future__ import annotations

import json
import os
import re
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

import httpx
import openpyxl
from dotenv import load_dotenv
from openai import OpenAI

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")
OUT = ROOT / "output/web_expand/global_wearable_medical_devices_market_global"
XLSX = OUT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
CACHE = OUT / "_wearable_channel_relevance_cache.json"
AUDIT = OUT / "_wearable_channel_relevance_audit.json"

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0.0.0 Safari/537.36"

# Hard OEM / brand manufacturers (not pure channel for this landscape)
OEM_RE = re.compile(
    r"\b(medtronic|abbott|dexcom|masimo|garmin|fitbit|apple|samsung electronics|"
    r"philips|siemens healthineers|roche|boston scientific|withings|oura|whoop|"
    r"polar electro|suunto|huawei|xiaomi|omron healthcare|ge healthcare|resmed|"
    r"biotronik|alivecor|iRhythm|irhythm|senseonics|tandem diabetes|insulet|"
    r"edward lifesciences|stryker|zimmer biomet)\b",
    re.I,
)

TEMPLATED_RE = re.compile(
    r"^(Africa|African) (Health|Medical) (Distributors|Retail|Supplies|Equipment|"
    r"Distributor|Wholesale|Trading)|^Al [A-Za-z\-]+ Medical( Company)?$",
    re.I,
)


def emptyish(v: object) -> bool:
    return str(v or "").strip().lower() in {"", "n/a", "na", "none", "-", "not publicly disclosed"}


def load_rows() -> list[dict]:
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hr = next(i for i, r in enumerate(rows[:5]) if r and str(r[0] or "").strip().lower() == "company")
    hdr = [str(c or "").strip() for c in rows[hr]]
    ic = {h: i for i, h in enumerate(hdr) if h}
    out = []
    for r in rows[hr + 1 :]:
        if not r or not str(r[0] or "").strip():
            continue
        out.append({h: ("" if r[ic[h]] is None else str(r[ic[h]]).strip()) for h in ic})
    wb.close()
    return out


def probe_site(website: str) -> dict:
    url = (website or "").strip()
    if emptyish(url):
        return {"ok": False, "status": 0, "error": "no_website", "title": "", "final_url": ""}
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    try:
        with httpx.Client(
            headers={"User-Agent": UA},
            verify=False,
            timeout=httpx.Timeout(8.0, connect=4.0),
            follow_redirects=True,
        ) as client:
            r = client.get(url)
            text = r.text or ""
            title = ""
            m = re.search(r"<title[^>]*>([^<]+)</title>", text, re.I)
            if m:
                title = re.sub(r"\s+", " ", m.group(1)).strip()[:120]
            low = text.lower()
            parked = any(
                x in low
                for x in (
                    "domain is for sale",
                    "buy this domain",
                    "godaddy",
                    "sedo.com",
                    "parked free",
                    "this domain may be for sale",
                    "coming soon",
                )
            )
            channel_kw = any(
                k in low
                for k in (
                    "distributor",
                    "distribution",
                    "wholesale",
                    "wholesaler",
                    "reseller",
                    "channel partner",
                    "authorized dealer",
                    "medical supply",
                    "pharmacy wholesale",
                    "hospital supplier",
                )
            )
            oem_kw = any(
                k in low
                for k in (
                    "we manufacture",
                    "our manufacturing",
                    "oem manufacturer",
                    "medical device manufacturer",
                    "we design and manufacture",
                )
            )
            wearable_kw = any(
                k in low
                for k in (
                    "wearable",
                    "cgm",
                    "continuous glucose",
                    "pulse oxim",
                    "remote patient",
                    "rpm",
                    "smartwatch",
                    "fitness tracker",
                    "holter",
                    "ambulatory monitor",
                    "patch monitor",
                )
            )
            return {
                "ok": r.status_code < 400 and not parked,
                "status": r.status_code,
                "title": title,
                "final_url": str(r.url)[:160],
                "parked": parked,
                "channel_kw": channel_kw,
                "oem_kw": oem_kw,
                "wearable_kw": wearable_kw,
                "error": "",
            }
    except Exception as ex:  # noqa: BLE001
        return {
            "ok": False,
            "status": 0,
            "title": "",
            "final_url": url,
            "parked": False,
            "channel_kw": False,
            "oem_kw": False,
            "wearable_kw": False,
            "error": str(ex)[:120],
        }


def ds_client() -> tuple[OpenAI, str]:
    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not key:
        raise SystemExit("OPENAI_API_KEY missing")
    return OpenAI(api_key=key, base_url=base, timeout=60.0), model


def classify(cli: OpenAI, model: str, facts: dict) -> dict:
    prompt = f"""Classify this company for Global Wearable Medical Devices Market CHANNEL landscape.

KEEP (relevant=true) only if they are a distributor / wholesaler / importer / reseller / hospital supplier /
medical device distributor / pharmacy wholesaler / authorized channel partner that can sell wearable
medical devices or adjacent patient-monitoring / digital health devices (CGM, pulse ox, RPM wearables, holter, patches).

DROP (relevant=false) if primarily: OEM/manufacturer brand, pure drug manufacturer with no device channel,
fake/templated company, parked domain, directory, hospital/clinic operator only, telecom/IT only,
consumer electronics retailer with no medical device channel evidence, or unrelated.

Return STRICT JSON:
{{
  "relevant": true/false,
  "role": "medical_device_distributor|pharma_wholesaler_adjacent|hospital_supplier|reseller|oem_manufacturer|fake_or_parked|off_market|unclear",
  "wearable_medical_link": "direct|adjacent_patient_monitoring|pharma_channel_weak|none",
  "confidence": 0.0-1.0,
  "action": "keep|drop|weak",
  "reason": "one short sentence"
}}

Facts:
{json.dumps(facts, ensure_ascii=False)}
"""
    resp = cli.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "Strict channel-map QA. JSON only. Never invent."},
            {"role": "user", "content": prompt},
        ],
        temperature=0,
        max_tokens=250,
        response_format={"type": "json_object"},
    )
    text = (resp.choices[0].message.content or "").strip()
    try:
        data = json.loads(text)
    except Exception:
        m = re.search(r"\{.*\}", text, re.S)
        data = json.loads(m.group(0)) if m else {}
    return {
        "relevant": bool(data.get("relevant")),
        "role": str(data.get("role") or "unclear"),
        "wearable_medical_link": str(data.get("wearable_medical_link") or "none"),
        "confidence": float(data.get("confidence") or 0),
        "action": str(data.get("action") or ("keep" if data.get("relevant") else "drop")),
        "reason": str(data.get("reason") or "")[:240],
    }


def main() -> None:
    rows = load_rows()
    print(f"companies={len(rows)}", flush=True)
    cache = {}
    if CACHE.exists():
        try:
            cache = json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    # Phase 1: website probes
    probes: dict[str, dict] = cache.get("_probes") or {}
    todo = []
    for r in rows:
        web = r.get("Website") or ""
        key = web.strip().lower()
        if key and key not in probes:
            todo.append(web)
    print(f"probing {len(todo)} new websites (cached {len(probes)}) ...", flush=True)
    with ThreadPoolExecutor(max_workers=16) as ex:
        futs = {ex.submit(probe_site, w): w for w in todo}
        done = 0
        for fut in as_completed(futs):
            w = futs[fut]
            probes[w.strip().lower()] = fut.result()
            done += 1
            if done % 50 == 0 or done == len(todo):
                print(f"  probed {done}/{len(todo)}", flush=True)
    cache["_probes"] = probes
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    # Phase 2: DeepSeek classify
    cli, model = ds_client()
    print(f"DeepSeek model={model}", flush=True)
    results = []
    for i, r in enumerate(rows, 1):
        co = r["Company"]
        web = r.get("Website") or ""
        ck = f"ds::{co}::{web.lower()}"
        probe = probes.get(web.strip().lower(), {})

        # hard heuristic preflags
        hard_drop = None
        if TEMPLATED_RE.match(co) and (not probe.get("ok") or probe.get("parked")):
            hard_drop = "templated_name_unreachable_or_parked"
        if OEM_RE.search(co) or OEM_RE.search(web):
            hard_drop = "oem_manufacturer_name"
        if probe.get("parked"):
            hard_drop = "parked_domain"
        # smartwatch-only core with no medical signal in row text
        core = r.get("Core Categories") or ""
        blob = " ".join([co, core, r.get("Specialty Focus") or "", r.get("Summary") or ""]).lower()
        if "smartwatch" in core.lower() and not any(
            k in blob for k in ("medical", "cgm", "patient", "hospital", "clinical", "oximetry", "rpm")
        ):
            # still ask DS, but note leak
            pass

        if hard_drop and ck not in cache:
            hit = {
                "relevant": False,
                "role": "fake_or_parked" if "parked" in hard_drop or "templated" in hard_drop else "oem_manufacturer",
                "wearable_medical_link": "none",
                "confidence": 0.9,
                "action": "drop",
                "reason": hard_drop,
                "source": "heuristic",
            }
            cache[ck] = hit
        elif ck in cache and isinstance(cache[ck], dict) and cache[ck].get("action"):
            hit = cache[ck]
        else:
            facts = {
                "company": co,
                "website": web,
                "hq": r.get("Headquarters"),
                "core": core,
                "specialty": r.get("Specialty Focus"),
                "distribution_type": r.get("Distribution Type"),
                "summary": (r.get("Summary") or "")[:500],
                "site_probe": {
                    "ok": probe.get("ok"),
                    "status": probe.get("status"),
                    "title": probe.get("title"),
                    "parked": probe.get("parked"),
                    "channel_kw": probe.get("channel_kw"),
                    "oem_kw": probe.get("oem_kw"),
                    "wearable_kw": probe.get("wearable_kw"),
                    "error": probe.get("error"),
                },
            }
            print(f"[{i}/{len(rows)}] DS {co[:45]}", flush=True)
            try:
                hit = classify(cli, model, facts)
                hit["source"] = "deepseek"
            except Exception as ex:  # noqa: BLE001
                hit = {
                    "relevant": False,
                    "role": "unclear",
                    "wearable_medical_link": "none",
                    "confidence": 0.0,
                    "action": "weak",
                    "reason": f"classify_error:{ex}"[:200],
                    "source": "error",
                }
            cache[ck] = hit
            if i % 25 == 0:
                CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
            time.sleep(0.08)

        results.append(
            {
                "company": co,
                "website": web,
                "probe_ok": probe.get("ok"),
                "probe_title": probe.get("title"),
                **{k: hit.get(k) for k in ("relevant", "role", "wearable_medical_link", "confidence", "action", "reason", "source")},
            }
        )

    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    actions = Counter(r["action"] for r in results)
    roles = Counter(r["role"] for r in results)
    drop = [r for r in results if r["action"] == "drop"]
    weak = [r for r in results if r["action"] == "weak"]
    keep = [r for r in results if r["action"] == "keep"]
    probe_fail = sum(1 for r in results if not r.get("probe_ok"))

    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "file": str(XLSX),
        "total": len(results),
        "actions": dict(actions),
        "roles": dict(roles),
        "probe_fail_or_bad": probe_fail,
        "keep": len(keep),
        "drop": len(drop),
        "weak": len(weak),
        "verdict": "NOT_ALL_CHANNEL"
        if len(drop) + len(weak) > 0
        else "ALL_LOOK_CHANNEL",
        "drop_samples": drop[:60],
        "weak_samples": weak[:40],
        "keep_samples": keep[:20],
        "model": model,
        "rule": "DeepSeek + site probe + OEM/template heuristics; wearable medical channel only",
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: audit[k] for k in audit if k not in ("drop_samples", "weak_samples", "keep_samples")}, indent=2))
    print("DROP examples:")
    for r in drop[:25]:
        print(" -", r["company"], "|", r["role"], "|", r["reason"][:90])


if __name__ == "__main__":
    main()
