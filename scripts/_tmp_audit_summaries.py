"""Audit Summary column quality across run4 FINAL Excels."""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")


def find_header(ws):
    for r in range(1, 8):
        vals = [c.value for c in ws[r]]
        if any(v and str(v).strip().lower() == "company" for v in vals):
            return r, vals
    return 1, [c.value for c in ws[1]]


def looks_bad(summary: str) -> list[str]:
    reasons = []
    s = (summary or "").strip()
    if not s:
        reasons.append("blank")
        return reasons
    if len(s) < 40:
        reasons.append("too_short")
    if len(s) > 600:
        reasons.append("too_long")
    # non-latin heavy
    non_ascii = sum(1 for c in s if ord(c) > 127)
    if non_ascii / max(len(s), 1) > 0.15:
        reasons.append("non_english_heavy")
    # keyword salad / nav junk
    low = s.lower()
    if any(x in low for x in ("skip to content", "cookie", "open menu", "add to cart", "privacy policy")):
        reasons.append("nav_junk")
    if s.count("|") >= 3 or s.count("•") >= 5:
        reasons.append("bullet_salad")
    # single run-on with no spaces weirdness
    if "  " in s and s.count("  ") > 8:
        reasons.append("weird_spacing")
    # looks like raw crawl dump
    if low.startswith("http") or low.startswith("www."):
        reasons.append("url_only")
    # Cyrillic / CJK / Arabic blocks
    if re.search(r"[\u0400-\u04FF]{8,}", s):
        reasons.append("cyrillic")
    if re.search(r"[\u4e00-\u9fff]{6,}", s):
        reasons.append("cjk")
    if re.search(r"[\u0600-\u06FF]{6,}", s):
        reasons.append("arabic")
    # all one line and very long without sentence breaks
    if "\n" not in s and len(s) > 220 and s.count(".") + s.count("!") + s.count("?") < 2:
        reasons.append("unreadable_runon")
    return reasons


def main() -> None:
    for mid in [
        "global_smartwatch_market_global",
        "global_green_chemical_market_global",
        "global_firewall_market_global",
    ]:
        short = re.sub(r"_global$", "", mid)
        xp = ROOT / mid / f"{short}_FINAL.xlsx"
        wb = load_workbook(xp, data_only=True)
        ws = wb.active
        hr, headers = find_header(ws)
        hmap = {str(h).strip(): i for i, h in enumerate(headers) if h}
        ci = hmap["Company"]
        si = hmap.get("Summary")
        ri = hmap.get("Regions Served")
        print(f"\n=== {mid} ===")
        rows = []
        for ridx in range(hr + 1, ws.max_row + 1):
            vals = [c.value for c in ws[ridx]]
            if not vals or not vals[ci]:
                continue
            rows.append((ridx, vals))
        print(f"n={len(rows)}")
        # last 3 companies
        print("-- last 3 companies --")
        for ridx, vals in rows[-3:]:
            co = vals[ci]
            sm = vals[si] if si is not None else ""
            rg = vals[ri] if ri is not None else ""
            print(f"  row {ridx} | {co}")
            print(f"    regions={rg!r}")
            print(f"    summary[{len(str(sm or ''))}]={str(sm or '')[:300]!r}")
            print(f"    bad={looks_bad(str(sm or ''))}")
        # all bad
        bad_rows = []
        for ridx, vals in rows:
            sm = str(vals[si] or "") if si is not None else ""
            reasons = looks_bad(sm)
            if reasons:
                bad_rows.append((str(vals[ci]), reasons, sm[:180]))
        print(f"-- bad summaries: {len(bad_rows)} --")
        for co, reasons, sm in bad_rows[:25]:
            print(f"  {co[:40]:40} | {reasons} | {sm!r}")
        wb.close()


if __name__ == "__main__":
    main()
