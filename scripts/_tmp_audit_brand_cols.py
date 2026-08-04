"""Audit Core Categories, Specialty Focus, Price Segment, Key Brands."""
from __future__ import annotations

import json
import re
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
COLS = ["Core Categories", "Specialty Focus", "Price Segment", "Key Brands Represented"]
JUNK_RE = re.compile(
    r"professional source|the professional|like something|lorem|placeholder|"
    r"\btodo\b|\btbd\b|see website|various brands|n/?a\b|unknown|"
    r"industry manufacturer|service provider|^yes$|^no$|"
    r"^\{|^\[|certificat|iso\s*\d|dict_keys",
    re.I,
)


def find_header(ws):
    for r in range(1, 8):
        vals = [c.value for c in ws[r]]
        if any(v and str(v).strip().lower() == "company" for v in vals):
            return r, vals
    return 1, [c.value for c in ws[1]]


def audit_xlsx(path: Path) -> None:
    if not path.exists():
        print("MISSING", path)
        return
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    hr, headers = find_header(ws)
    hmap = {str(h).strip(): i for i, h in enumerate(headers) if h}
    print(f"\n=== {path.name} ===")
    ci = hmap.get("Company")
    for col in COLS:
        idx = hmap.get(col)
        if idx is None:
            print(f"  {col}: NOT FOUND")
            continue
        blank = 0
        junk = []
        samples = []
        total = 0
        for row in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not row or ci is None or not row[ci]:
                continue
            total += 1
            val = row[idx]
            s = "" if val is None else str(val).strip()
            if not s:
                blank += 1
            else:
                if JUNK_RE.search(s) or len(s) > 180:
                    junk.append((str(row[ci])[:40], s[:140]))
                if len(samples) < 6:
                    samples.append((str(row[ci])[:35], s[:100]))
        print(f"  {col}: total={total} blank={blank} junkish={len(junk)}")
        for co, s in junk[:12]:
            print(f"    JUNK {co} | {s}")
        for co, s in samples:
            print(f"    eg   {co} | {s}")
    wb.close()


def audit_json(mid: str) -> None:
    path = ROOT / mid / f"{mid}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    rows = data.get("relevant_companies") or []
    print(f"\n=== JSON {mid} n={len(rows)} crawl field samples ===")
    keys = [
        "key_brands_represented",
        "key_brands_source",
        "core_categories",
        "specialty_focus",
        "price_segment",
        "products",
        "industries",
        "certifications",
        "company_type",
    ]
    blank = {k: 0 for k in keys}
    junk_brands = []
    for r in rows:
        c = r.get("_crawl") if isinstance(r.get("_crawl"), dict) else {}
        co = str(r.get("company") or "")
        for k in keys:
            v = str(c.get(k) or "").strip()
            if not v:
                blank[k] += 1
        kb = str(c.get("key_brands_represented") or "").strip()
        if kb and JUNK_RE.search(kb):
            junk_brands.append((co, kb[:140]))
        # also check company-level fields
        for alt in ("core_categories", "specialty_focus", "price_segment"):
            if r.get(alt):
                pass
    print("  blank crawl counts:", blank)
    print("  junk key_brands:", len(junk_brands))
    for co, s in junk_brands[:20]:
        print(f"    {co[:40]} | {s}")
    # show one full crawl keys of first row with brands
    for r in rows[:3]:
        c = r.get("_crawl") or {}
        print("  sample company", r.get("company"))
        for k in keys:
            if c.get(k):
                print(f"    {k}={str(c.get(k))[:120]}")


def main() -> None:
    for mid in [
        "global_smartwatch_market_global",
        "global_green_chemical_market_global",
        "global_firewall_market_global",
    ]:
        short = re.sub(r"_global$", "", mid)
        for suffix in ["_FINAL.xlsx", "_founded_and_hq_v2_cleaned.xlsx"]:
            audit_xlsx(ROOT / mid / f"{short}{suffix}")
        audit_json(mid)


if __name__ == "__main__":
    main()
