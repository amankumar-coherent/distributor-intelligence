"""Final polish: merge acquired brands, fix remaining verified URLs, drop dead DNS without replacement."""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "unique_website_cleanup_pass2_audit.json"

# Rename / re-home exact companies into existing canonical rows
REHOME = {
    # acquired / same entity
    "service paper": "bunzl",  # Bunzl acquired Service Paper Co.
    "central paper products": "imperial dade",  # Imperial Dade acquired Central Paper
    "packaging suppliers": "wholesale packaging",
    "packaging suppliers uk": "wholesale packaging",
}

# Create or update canonical rows
CANONICAL = {
    "bunzl": ("Bunzl", "https://www.bunzl.com"),
    "imperial dade": ("Imperial Dade", "https://www.imperialdade.com"),
    "wholesale packaging": ("Wholesale Packaging", "https://www.wholesalepackaging.co.uk"),
}

# Exact website/name fixes (verified)
FIX = {
    "u.s. plastic corp.": ("U.S. Plastic Corp.", "https://usplastic.com"),
    "inapa": ("Inapa", "https://www.inapashapingideas.pt"),
    "verpackungsdiscount": ("Verpackungsdiscount", "https://verpackungsdiscount.de"),
    "officeworks au": ("Officeworks", "https://www.officeworks.com.au"),
    "blackwoods au": ("Blackwoods", "https://www.blackwoods.com.au"),
    "the cary company": ("The Cary Company", "https://www.thecarycompany.com"),
    "reliable corporation": ("Reliable", "https://www.reliable.com"),
}


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def domain_of(url: str) -> str:
    if not url:
        return ""
    u = url if "://" in url else "https://" + url
    try:
        return urlparse(u).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def uniq_key(name: str, web: str) -> str:
    n = norm(name).lower()
    n = re.sub(r"\s+(inc\.?|ltd\.?|llc|corp\.?|corporation|company|co\.)$", "", n, flags=re.I)
    return f"{n}|{domain_of(web)}"


def main():
    backup = FP.with_name(
        FP.stem + f"_pre_unique_web_pass2_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    shutil.copy2(FP, backup)
    print("backup", backup)

    wb = load_workbook(FP)
    ws = wb.active
    header = [c.value for c in ws[1]]
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]

    by_name = {norm(str(r[0] or "")).lower(): r for r in rows}
    actions = []

    # Apply FIX first
    for r in rows:
        low = norm(str(r[0] or "")).lower()
        if low in FIX:
            nn, nw = FIX[low]
            actions.append({"action": "fix", "from": r[0], "to": nn, "website": nw})
            r[0] = nn
            r[1] = nw

    # Rebuild name map
    by_name = {norm(str(r[0] or "")).lower(): r for r in rows}

    # Rehome / merge
    keep = []
    seen_canon = set()
    for r in rows:
        low = norm(str(r[0] or "")).lower()
        if low in REHOME:
            target = REHOME[low]
            cname, cweb = CANONICAL[target]
            # if target already exists, drop this row into it
            if cname.lower() in by_name or target in seen_canon:
                actions.append(
                    {
                        "action": "merge_into",
                        "from": r[0],
                        "into": cname,
                        "website": cweb,
                        "reason": "same_entity_or_acquisition",
                    }
                )
                # enrich presence on existing if present later
                continue
            # transform this row into canonical
            actions.append(
                {
                    "action": "rehome",
                    "from": r[0],
                    "to": cname,
                    "website": cweb,
                }
            )
            r[0] = cname
            r[1] = cweb
            seen_canon.add(target)
            keep.append(r)
            continue
        keep.append(r)

    # Ensure Imperial Dade / Wholesale Packaging / Bunzl exist with right URLs
    # (Bunzl already exists)
    have = {norm(str(r[0] or "")).lower() for r in keep}
    for key, (cname, cweb) in CANONICAL.items():
        if cname.lower() not in have:
            # only create if we merged something into it and lost the row
            if any(REHOME.get(norm(str(r[0] or "")).lower()) == key for r in rows):
                # create minimal row from header length
                newr = [None] * len(header)
                newr[0] = cname
                newr[1] = cweb
                newr[12] = "Specialty Distributor"
                newr[20] = f"{cname} packaging/supplies distributor channel (verified website)."
                keep.append(newr)
                actions.append({"action": "create_canonical", "company": cname, "website": cweb})
                have.add(cname.lower())

    # Force update existing Bunzl / Imperial Dade / Wholesale Packaging websites
    for r in keep:
        low = norm(str(r[0] or "")).lower()
        for cname, cweb in CANONICAL.values():
            if low == cname.lower():
                if norm(str(r[1] or "")) != cweb:
                    actions.append(
                        {
                            "action": "force_website",
                            "company": r[0],
                            "from": r[1],
                            "to": cweb,
                        }
                    )
                r[0] = cname
                r[1] = cweb

    # Final uniqueness
    final = []
    seen = set()
    dups = []
    for r in keep:
        r[0] = norm(str(r[0] or ""))
        web = norm(str(r[1] or ""))
        if web and not web.startswith("http"):
            web = "https://" + web
        if web.endswith("/") and web.count("/") > 2:
            web = web.rstrip("/")
        r[1] = web
        if len(r) > 12:
            r[12] = "Specialty Distributor"
        uk = uniq_key(r[0], r[1])
        # also unique by company name alone
        nk = norm(r[0]).lower()
        if uk in seen or nk in {norm(x[0]).lower() for x in final}:
            dups.append({"company": r[0], "website": r[1]})
            continue
        seen.add(uk)
        final.append(r)

    final.sort(key=lambda x: str(x[0] or "").lower())

    out = Workbook()
    ws2 = out.active
    ws2.title = "Channel Landscape"
    ws2.append(header)
    for c in ws2[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in final:
        row = list(r) + [None] * (len(header) - len(r))
        ws2.append(row[: len(header)])
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=len(header)):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    widths = {1: 34, 2: 38, 3: 12, 4: 22, 5: 22, 6: 36, 13: 20, 21: 55}
    for i, w in widths.items():
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A2"
    out.save(FP)

    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "before": len(rows),
        "after": len(final),
        "actions": actions,
        "dups_dropped": dups,
        "file": str(FP),
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print("DONE before", len(rows), "after", len(final), "actions", len(actions), "dups", len(dups))
    print("audit", AUDIT)


if __name__ == "__main__":
    main()
