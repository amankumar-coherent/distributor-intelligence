# -*- coding: utf-8 -*-
"""
Full re-verify pass for global_flexible_packaging_market_global_FINAL.xlsx:
- Core Categories: main product categories the distributor sells/distributes
- Specialty Focus: specific market/technology/segment/customer focus the distributor
  is particularly known for
- Key Brands Represented: manufacturer/brand names the distributor officially
  distributes, resells, represents, or carries

For every row: fetch website evidence (home/about/products/brands/partners pages),
then send current values + evidence to DeepSeek for verification/correction.
Only overwrite when DeepSeek returns high/medium confidence; never invent brands
without evidence. Resumable via checkpoint; backs up the workbook before writing.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import httpx
from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel")
load_dotenv(ROOT / ".env")

WORKDIR = ROOT / "output" / "web_expand" / "global_flexible_packaging_market_global"
XLSX = WORKDIR / "global_flexible_packaging_market_global_FINAL.xlsx"
AUDIT = WORKDIR / "_flexpack_verify_core_spec_brand_audit.json"
CKPT = WORKDIR / "_flexpack_verify_core_spec_brand_ckpt.json"
NPD = "Not publicly disclosed"

UA = {"User-Agent": "Mozilla/5.0"}

CANDIDATE_PATHS = [
    "",
    "/about",
    "/about-us",
    "/products",
    "/brands",
    "/our-brands",
    "/partners",
    "/manufacturers",
    "/principals",
    "/company",
]


def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
    html = re.sub(r"(?is)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", html).strip()


def fetch(url: str) -> str:
    try:
        with httpx.Client(follow_redirects=True, timeout=8.0, headers=UA) as c:
            r = c.get(url)
            if r.status_code >= 400:
                return ""
            return strip_html(r.text)[:6000]
    except Exception:
        return ""


def candidate_urls(website: str) -> list[str]:
    base = (website or "").strip().rstrip("/")
    if not base:
        return []
    if not base.startswith("http"):
        base = "https://" + base
    out = []
    for p in CANDIDATE_PATHS:
        u = base if p == "" else urljoin(base + "/", p.lstrip("/"))
        if u not in out:
            out.append(u)
    return out


def gather_evidence(website: str, max_chars: int = 3500, max_urls: int = 5) -> str:
    snippets = []
    total = 0
    for u in candidate_urls(website)[:max_urls]:
        t = fetch(u)
        if len(t) > 200:
            snippets.append(f"URL:{u}\n{t[:2500]}")
            total += len(t)
        time.sleep(0.1)
        if total > max_chars:
            break
    return "\n\n---\n\n".join(snippets)


def clean_core(text: str) -> str:
    t = re.sub(r"\s+", " ", (text or "").strip())
    t = t.replace("|", ";").replace(",", ";")
    parts = [p.strip(" .;") for p in re.split(r"[;•/]+", t) if p.strip(" .;")]
    out, seen = [], set()
    for p in parts:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return "; ".join(out[:8])


def clean_specialty(text: str) -> str:
    t = re.sub(r"\s+", " ", (text or "").strip())
    if not t:
        return ""
    t = re.sub(r"[,;]?\s*others\.?$", "", t, flags=re.I).strip()
    if len(t) > 420:
        t = t[:417].rsplit(" ", 1)[0] + "..."
    if t and t[-1] not in ".!?":
        t += "."
    return t


def clean_brands(raw: str) -> str:
    if not raw:
        return ""
    t = re.sub(r"\s+", " ", raw.strip())
    if t.lower() in {"npd", "n/a", "none", "unknown", "", NPD.lower()}:
        return ""
    parts = re.split(r"[,;/|]+", t)
    out, seen = [], set()
    junk = {"others", "various", "multiple", "n/a", "none", "unknown", "brands", "brand", "etc"}
    for p in parts:
        b = re.sub(r"\s+", " ", p).strip(" .")
        if not b or len(b) < 2 or b.lower() in junk or len(b) > 60:
            continue
        key = b.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(b)
    return "; ".join(out[:14])


def parse_json(text: str) -> dict:
    text = (text or "").strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"\{[\s\S]*\}", text)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                return {}
    return {}


SYSTEM_PROMPT = (
    "You are verifying three columns of a flexible-packaging CHANNEL DISTRIBUTOR "
    "spreadsheet row. The spreadsheet already has a CURRENT value for each field, "
    "produced by earlier deep research (it may be based on more/better sources than "
    "the single evidence snippet you are given now). Your job is narrow: flag ONLY "
    "clear errors, not incompleteness.\n\n"
    "Definitions:\n"
    "1. core_categories: the MAIN PRODUCT CATEGORIES this distributor sells/distributes.\n"
    "2. specialty_focus: the SPECIFIC MARKET, TECHNOLOGY, PRODUCT SEGMENT, or CUSTOMER "
    "SEGMENT this distributor is particularly known for or focused on.\n"
    "3. key_brands: the manufacturer/brand names whose products this distributor officially "
    "distributes, resells, represents, or carries.\n\n"
    "CRITICAL RULE for core_categories and specialty_focus: the evidence snippet you see is "
    "PARTIAL (often just a homepage, sometimes in another language, sometimes fetch failed). "
    "The absence of a topic in the evidence does NOT mean the current value is wrong — do "
    "NOT trim or narrow the current value just because this snippet doesn't happen to mention "
    "every item. Only mark current_accurate=false when the evidence ACTIVELY CONTRADICTS the "
    "current value (wrong industry entirely, factually wrong claim) OR the current value is "
    "obviously broken (empty, boilerplate like 'is a distributor of', echoes the company name, "
    "a single vague word, or an unrelated template). When in doubt, current_accurate=true.\n\n"
    "For key_brands: this field is about finding brand/manufacturer names this distributor "
    "carries. Look for brand names in the evidence (product pages, 'brands we carry', "
    "'partners', 'principals', 'manufacturers' sections, product listings). Report ONLY "
    "brand names you actually see in the evidence text — never invent or guess. If none are "
    "visible in the evidence, return an empty string for key_brands_found (do not treat this "
    "as an error; most homepages don't list brands).\n\n"
    'Return strict JSON: {'
    '"core_categories_current_accurate": true|false, '
    '"core_categories_corrected": "only fill if false, else empty string", '
    '"specialty_focus_current_accurate": true|false, '
    '"specialty_focus_corrected": "only fill if false, else empty string", '
    '"key_brands_found": "brand names seen in evidence, separated by \'; \', or empty string", '
    '"confidence": "high|medium|low|none", '
    '"notes": "short reason"}'
)


def deepseek_verify(client_key: str, base: str, model: str, company: str, website: str,
                     cur_core: str, cur_spec: str, cur_brand: str, summary: str, evidence: str) -> dict:
    user = (
        f"Company: {company}\n"
        f"Website: {website}\n\n"
        f"Current Core Categories: {cur_core or '(empty)'}\n"
        f"Current Specialty Focus: {cur_spec or '(empty)'}\n"
        f"Current Key Brands Represented: {cur_brand or '(empty)'}\n\n"
        f"Existing Summary (from earlier deep research — treat as reliable context): "
        f"{(summary or '')[:900]}\n\n"
        f"Fresh Website Evidence (partial, this pass only):\n"
        f"{evidence[:6500] if evidence else '(no website text retrieved this pass)'}"
    )
    try:
        with httpx.Client(timeout=60.0) as c:
            r = c.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {client_key}", "Content-Type": "application/json"},
                json={
                    "model": model,
                    "temperature": 0,
                    "max_tokens": 500,
                    "messages": [
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {"role": "user", "content": user},
                    ],
                    "response_format": {"type": "json_object"},
                },
            )
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            return parse_json(content)
    except Exception as e:
        return {"confidence": "none", "notes": f"error: {e}"}


def main() -> None:
    api_key = (os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY") or "").strip()
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    if not api_key:
        raise SystemExit("Missing DeepSeek/OpenAI API key in .env")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = XLSX.with_name(XLSX.stem + f"_BEFORE_VERIFY_CORE_SPEC_BRAND_{stamp}" + XLSX.suffix)
    shutil.copy2(XLSX, bak)
    print(f"backup -> {bak.name}", flush=True)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    ckpt = {"done": {}}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    rows = list(range(2, ws.max_row + 1))
    print(f"total rows={len(rows)}", flush=True)

    core_changed = spec_changed = brand_changed = 0
    processed = 0
    failed = []

    for i, r in enumerate(rows, 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        if co in ckpt.get("done", {}):
            continue

        website = str(ws.cell(r, ic["Website"]).value or "").strip()
        cur_core = str(ws.cell(r, ic["Core Categories"]).value or "").strip()
        cur_spec = str(ws.cell(r, ic["Specialty Focus"]).value or "").strip()
        cur_brand = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
        cur_brand_norm = "" if cur_brand.lower() == NPD.lower() else cur_brand
        summary = str(ws.cell(r, ic.get("Summary", 0)).value or "") if "Summary" in ic else ""

        evidence = gather_evidence(website)
        ds = deepseek_verify(api_key, base, model, co, website, cur_core, cur_spec, cur_brand_norm, summary, evidence)
        conf = str(ds.get("confidence") or "").lower()

        core_ok = bool(ds.get("core_categories_current_accurate", True))
        spec_ok = bool(ds.get("specialty_focus_current_accurate", True))
        new_core = clean_core(str(ds.get("core_categories_corrected") or ""))
        new_spec = clean_specialty(str(ds.get("specialty_focus_corrected") or ""))
        found_brands = clean_brands(str(ds.get("key_brands_found") or ""))

        applied = {"core": False, "spec": False, "brand": False}

        if conf in {"high", "medium"} and not core_ok and new_core and new_core.lower() != cur_core.lower():
            ws.cell(r, ic["Core Categories"]).value = new_core
            core_changed += 1
            applied["core"] = True
        if conf in {"high", "medium"} and not spec_ok and new_spec and new_spec.lower() != cur_spec.lower():
            ws.cell(r, ic["Specialty Focus"]).value = new_spec
            spec_changed += 1
            applied["spec"] = True

        if conf in {"high", "medium"} and found_brands:
            if cur_brand_norm:
                existing = [b.strip() for b in re.split(r"[,;]+", cur_brand_norm) if b.strip()]
                existing_lower = {b.lower() for b in existing}
                additions = [b for b in found_brands.split("; ") if b.lower() not in existing_lower]
                if additions:
                    merged = "; ".join(existing + additions)
                    ws.cell(r, ic["Key Brands Represented"]).value = merged
                    brand_changed += 1
                    applied["brand"] = True
            else:
                ws.cell(r, ic["Key Brands Represented"]).value = found_brands
                brand_changed += 1
                applied["brand"] = True
        elif not cur_brand_norm and not str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip():
            ws.cell(r, ic["Key Brands Represented"]).value = NPD

        if conf == "none" and not evidence:
            failed.append(co)

        processed += 1
        ckpt.setdefault("done", {})[co] = {
            "applied": applied,
            "confidence": conf,
            "core": ws.cell(r, ic["Core Categories"]).value,
            "spec": ws.cell(r, ic["Specialty Focus"]).value,
            "brand": ws.cell(r, ic["Key Brands Represented"]).value,
        }

        if processed % 5 == 0 or i <= 5:
            print(
                f"[{i}/{len(rows)}] {co} conf={conf} core={applied['core']} "
                f"spec={applied['spec']} brand={applied['brand']}",
                flush=True,
            )
        if processed % 10 == 0:
            CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
            wb.save(XLSX)
            print(f"checkpoint saved at processed={processed}", flush=True)

    wb.save(XLSX)
    CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    npd_final = 0
    for r in rows:
        v = str(ws.cell(r, ic["Key Brands Represented"]).value or "").strip()
        if not v or v.lower() == NPD.lower():
            npd_final += 1

    audit = {
        "backup": bak.name,
        "total_rows": len(rows),
        "processed_this_run": processed,
        "core_changed": core_changed,
        "spec_changed": spec_changed,
        "brand_changed": brand_changed,
        "no_evidence_no_confidence": failed,
        "npd_brand_final": npd_final,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(audit, indent=2, ensure_ascii=False))
    print("DONE", flush=True)


if __name__ == "__main__":
    main()
