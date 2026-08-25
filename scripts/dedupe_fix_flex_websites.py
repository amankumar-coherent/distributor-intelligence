"""Deduplicate multi-location companies; fix verified websites; remove bad/unverifiable rows.

Rules:
- Never invent websites; only apply verified URLs from research.
- One row per unique company (merge country clones).
- Remove manufacturers / dead DNS with no verified replacement / wrong market.
"""
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
AUDIT = FP.parent / "unique_website_cleanup_audit.json"

# Canonical group key -> (canonical_name, corporate_website, presence_note)
MERGE_GROUPS = {
    "manutan": (
        "Manutan",
        "https://www.manutan.com",
        "Europe (17 countries / 25 subsidiaries)",
    ),
    "raja": (
        "RAJA",
        "https://raja-group.com",
        "Europe (19 countries; RAJAPACK network)",
    ),
    "ratioform": (
        "Ratioform",
        "https://www.ratioform.de",
        "Europe (multi-country Ratioform/RAJA network)",
    ),
    "viking direct": (
        "Viking Direct",
        "https://www.vikingdirect.co.uk",
        "Europe (multi-country Viking Direct / Lyreco network)",
    ),
    "lyreco": (
        "Lyreco",
        "https://www.lyreco.com",
        "Europe / Global (multi-country Lyreco)",
    ),
    "bunzl": (
        "Bunzl",
        "https://www.bunzl.com",
        "Global (incl. AU, Ireland, Canada operations)",
    ),
    "uline": (
        "Uline",
        "https://www.uline.com",
        "North America & Europe (Uline.com / .ca / .eu / .co.uk)",
    ),
    "transpak": (
        "Transpak",
        "https://www.transpak.de",
        "Europe (DE/AT/CH/NL Transpak)",
    ),
    "ahlsell": (
        "Ahlsell",
        "https://www.ahlsell.com",
        "Nordics (SE/DK/NO and group footprint)",
    ),
    "office depot": (
        "Office Depot",
        "https://www.officedepot.com",
        "North America & Europe / MX (Office Depot network)",
    ),
    "antalis": (
        "Antalis",
        "https://www.antalis.com",
        "Europe & Americas (Antalis papers/packaging distribution)",
    ),
    "staples": (
        "Staples",
        "https://www.staples.com",
        "North America & Europe (Staples network)",
    ),
    "fastenal": (
        "Fastenal",
        "https://www.fastenal.com",
        "North America (US/Canada)",
    ),
    "veritiv": (
        "Veritiv",
        "https://www.veritiv.com",
        "North America (incl. former Unisource Canada)",
    ),
    "unisource": (
        "Veritiv",
        "https://www.veritiv.com",
        "North America (Unisource Canada now Veritiv)",
    ),
    "supplyone": (
        "SupplyOne",
        "https://www.supplyone.com",
        "North America (incl. SupplyOne Canada / former Crownhill)",
    ),
    "crownhill": (
        "SupplyOne",
        "https://www.supplyone.com",
        "North America (Crownhill = SupplyOne Canada)",
    ),
    "grainger": (
        "Grainger",
        "https://www.grainger.com",
        "North America / MX (W.W. Grainger)",
    ),
    "w. w. grainger": (
        "Grainger",
        "https://www.grainger.com",
        "North America / MX (W.W. Grainger)",
    ),
    "jpg": (
        "JPG",
        "https://www.jpg.fr",
        "Europe (France/Belgium JPG / Manutan family)",
    ),
}

# Exact company name (lower) -> remove reason
REMOVE_EXACT = {
    "agenda boxes": "manufacturer (agendapackaging.co.uk) — not pure channel",
    "ideal packaging canada": "hybrid manufacturer+distributor (idealpkg.com)",
    "pride container": "no verified packaging-distributor website; DNS dead",
    "national packaging company": "natpkg.com DNS dead; entity ambiguous — not verified",
    "clayton packaging solutions": "claytonpackaging.com DNS dead; no verified match",
    "all size packaging canada": "allsizepackaging.com DNS dead; not verified as All Packaging Canada",
    "lone star containers": "wrong market / manufacturer ambiguity; DNS dead for listed URL",
    "cheap packaging materials": "placeholder domain DNS dead; not verified",
    "online packaging supplies uk": "placeholder domain DNS dead; not verified",
    "img packing": "listed domain DNS dead; not verified",
    "jaymar packaging": "listed domain DNS dead; not verified",
    "neptune packaging": "listed domain DNS dead; not verified",
    "emballages service france": "listed domain DNS dead; not verified",
}

# Exact company name (lower) -> (new_name, new_website) verified only
FIX_EXACT = {
    "the bottle crew": ("The Bottle Crew", "https://www.bottlecrew.com"),
    "uk packaging supplies london": ("UK Packaging Supplies", "https://www.ukplc.co.uk"),
    "the packaging centre ireland": ("The Packaging Centre", "https://www.thepackagingcentre.ie"),
    "signet au": ("Signet", "https://www.signet.net.au"),
    "us plastic corp": ("U.S. Plastic Corp.", "https://www.usplastic.com"),
    "w. w. grainger inc.": ("Grainger", "https://www.grainger.com"),
    "fastenal company": ("Fastenal", "https://www.fastenal.com"),
    "veritiv corporation": ("Veritiv", "https://www.veritiv.com"),
    "msc industrial supply co.": ("MSC Industrial Supply", "https://www.mscdirect.com"),
    "cospak pty ltd.": ("Cospak", "https://www.cospak.com.au"),
    "chase packaging ltd": ("Chase Packaging", "https://www.chase-packaging.co.uk"),
    "multiplastics europe ltd": ("Multiplastics Europe", "https://www.multiplastics.co.uk"),
    "super poly ltd": ("Super Poly", "https://superpoly.ca"),
    "o.berk company": ("O.Berk", "https://www.oberk.com"),
    "the cary company": ("The Cary Company", "https://www.thecarycompany.com"),
    "schermerhorn bros. co.": ("Schermerhorn Bros.", "https://www.schermerhornbrosco.com"),
    "crown packaging corp": ("Crown Packaging", "https://crownpack.com"),
    "excel pac llc": ("Excel Pac", "https://excelpacllc.com"),
    "reliable corporation": ("Reliable Corporation", "https://www.reliable.com"),
    "durable packaging corp": ("Durable Packaging", "https://www.durablepackaging.com"),
    "service paper company": ("Service Paper", "https://www.servicepaper.com"),
    "u.s. packaging & wrapping llc": (
        "U.S. Packaging & Wrapping",
        "https://uspackagingandwrapping.com",
    ),
    "seal packaging uk merchant": ("Seal Packaging", "https://www.sealpackaging.co.uk"),
    "transpak uk merchant": ("Transpack UK", "https://www.transpack.co.uk"),
    "raja group corporate": ("RAJA", "https://raja-group.com"),
    "office depot europe": ("Office Depot", "https://www.officedepot.com"),
    "staples europe": ("Staples", "https://www.staples.com"),
    "bunzl canada": ("Bunzl", "https://www.bunzl.com"),
    "unisource canada": ("Veritiv", "https://www.veritiv.com"),
    "supplyone canada (crownhill packaging)": ("SupplyOne", "https://www.supplyone.com"),
}


def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def group_key(name: str) -> str | None:
    n = norm_space(name).lower()
    n = re.sub(r"\s*\([^)]*\)\s*", " ", n)
    n = re.sub(r"\s+", " ", n).strip()

    # strip common location / legal suffixes for grouping
    for tok in [
        " group corporate",
        " corporate",
        " merchant",
        " canada (crownhill packaging)",
        " canada",
        " europe",
        " france",
        " germany",
        " italy",
        " spain",
        " netherlands",
        " belgium",
        " switzerland",
        " sweden",
        " norway",
        " denmark",
        " ireland",
        " portugal",
        " austria",
        " hungary",
        " romania",
        " slovakia",
        " poland",
        " luxembourg",
        " finland",
        " czech republic",
        " czech",
        " uk",
        " au",
        " mx",
        " za",
        " usa",
        " us",
        " pty ltd.",
        " pty ltd",
        " ltd.",
        " ltd",
        " llc",
        " inc.",
        " inc",
        " corp.",
        " corp",
        " corporation",
        " company",
        " co.",
        " limited",
    ]:
        if n.endswith(tok):
            n = n[: -len(tok)].strip()

    # special aliases
    if n.startswith("w. w. grainger") or n == "grainger":
        return "grainger"
    if n.startswith("viking"):
        return "viking direct"
    if n.startswith("raja"):
        return "raja"
    if n.startswith("manutan"):
        return "manutan"
    if n.startswith("ratioform"):
        return "ratioform"
    if n.startswith("lyreco"):
        return "lyreco"
    if n.startswith("bunzl"):
        return "bunzl"
    if n.startswith("uline"):
        return "uline"
    if n == "transpak" or n.startswith("transpak "):
        return "transpak"
    if n.startswith("ahlsell"):
        return "ahlsell"
    if n.startswith("office depot"):
        return "office depot"
    if n.startswith("antalis"):
        return "antalis"
    if n.startswith("staples"):
        return "staples"
    if n.startswith("fastenal"):
        return "fastenal"
    if n.startswith("veritiv") or n.startswith("unisource"):
        return "veritiv"
    if n.startswith("supplyone") or "crownhill" in n:
        return "supplyone"
    if n.startswith("jpg"):
        return "jpg"
    if n in MERGE_GROUPS:
        return n
    return None


def domain_of(url: str) -> str:
    if not url:
        return ""
    u = url if "://" in url else "https://" + url
    try:
        return urlparse(u).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def pick_better_row(a: list, b: list) -> list:
    """Prefer row with richer filled fields / corporate-looking website."""
    def score(r: list) -> int:
        s = 0
        for i, v in enumerate(r):
            if v is None or str(v).strip() == "" or str(v).strip().lower() in {
                "not publicly disclosed",
                "n/a",
                "na",
            }:
                continue
            s += 1
            if i == 2:  # founded year numeric-ish
                if re.search(r"\d{4}", str(v)):
                    s += 3
        web = str(r[1] or "").lower()
        if any(x in web for x in (".com/", ".com", "group", "corporate")):
            s += 1
        return s

    return a if score(a) >= score(b) else b


def main():
    backup = FP.with_name(
        FP.stem + f"_pre_unique_web_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    shutil.copy2(FP, backup)
    print("backup", backup)

    wb = load_workbook(FP)
    ws = wb.active
    header = [c.value for c in ws[1]]
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]

    removed = []
    fixed = []
    merged_log = []

    # Pass 1: remove + exact fixes
    staged = []
    for r in rows:
        name = norm_space(str(r[0] or ""))
        web = norm_space(str(r[1] or ""))
        low = name.lower()

        if low in REMOVE_EXACT:
            removed.append({"company": name, "website": web, "reason": REMOVE_EXACT[low]})
            continue

        if low in FIX_EXACT:
            new_name, new_web = FIX_EXACT[low]
            if new_name != name or new_web.rstrip("/") != web.rstrip("/"):
                fixed.append(
                    {
                        "from_company": name,
                        "to_company": new_name,
                        "from_website": web,
                        "to_website": new_web,
                        "reason": "verified_fix",
                    }
                )
            r[0] = new_name
            r[1] = new_web
            name = new_name
            web = new_web

        staged.append(r)

    # Pass 2: merge groups
    groups: dict[str, list] = {}
    singles: list[list] = []
    for r in staged:
        name = norm_space(str(r[0] or ""))
        gk = group_key(name)
        if gk and gk in MERGE_GROUPS:
            groups.setdefault(gk, []).append(r)
        else:
            singles.append(r)

    merged_rows = []
    for gk, glist in groups.items():
        canon_name, canon_web, presence = MERGE_GROUPS[gk]
        # merge
        best = glist[0]
        for other in glist[1:]:
            best = pick_better_row(best, other)
        best = list(best)
        old_names = [norm_space(str(x[0] or "")) for x in glist]
        old_webs = [norm_space(str(x[1] or "")) for x in glist]
        best[0] = canon_name
        best[1] = canon_web
        # Operational Presence col index 5
        if len(best) > 5:
            best[5] = presence
        # Continent / Geography col 4 — keep broad
        if len(best) > 4:
            if gk in {"uline", "bunzl", "veritiv", "supplyone", "grainger", "fastenal", "staples", "office depot"}:
                best[4] = "Global / Multi-region"
            else:
                best[4] = "Europe" if gk not in {"veritiv", "supplyone", "grainger", "fastenal"} else best[4]
        # Distribution type
        if len(best) > 12:
            best[12] = "Specialty Distributor"
        # Summary append uniqueness note
        if len(best) > 20:
            summ = str(best[20] or "").strip()
            note = f" Consolidated from {len(glist)} location rows into one unique company record."
            if note.strip() not in summ:
                best[20] = (summ + note).strip()
        merged_rows.append(best)
        merged_log.append(
            {
                "group": gk,
                "canonical_name": canon_name,
                "canonical_website": canon_web,
                "merged_from": old_names,
                "old_websites": old_webs,
                "count": len(glist),
            }
        )
        print(f"MERGE {gk}: {len(glist)} -> {canon_name} | {canon_web}")

    # Pass 3: normalize remaining singles websites (strip trailing slash consistency)
    final_rows = []
    seen_keys = set()
    dup_dropped = []

    def uniq_key(r: list) -> str:
        name = norm_space(str(r[0] or "")).lower()
        d = domain_of(str(r[1] or ""))
        # also collapse near-identical names
        name2 = re.sub(
            r"\s+(inc\.?|ltd\.?|llc|corp\.?|corporation|company|co\.)$",
            "",
            name,
            flags=re.I,
        )
        return f"{name2}|{d}"

    all_candidates = merged_rows + singles
    for r in all_candidates:
        r[0] = norm_space(str(r[0] or ""))
        web = norm_space(str(r[1] or ""))
        if web and not web.startswith("http"):
            web = "https://" + web
        # normalize trailing slash off for consistency (except keep as-is if only domain)
        if web.endswith("/") and web.count("/") > 2:
            web = web.rstrip("/")
        r[1] = web
        if len(r) > 12:
            r[12] = "Specialty Distributor"

        uk = uniq_key(r)
        if uk in seen_keys:
            dup_dropped.append({"company": r[0], "website": r[1], "reason": "duplicate_after_merge"})
            continue
        seen_keys.add(uk)
        final_rows.append(r)

    # sort by company name
    final_rows.sort(key=lambda x: str(x[0] or "").lower())

    # Write workbook
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "Channel Landscape"
    out_ws.append(header)
    for cell in out_ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for r in final_rows:
        # pad
        row = list(r) + [None] * (len(header) - len(r))
        out_ws.append(row[: len(header)])
    for row in out_ws.iter_rows(min_row=2, max_row=out_ws.max_row, max_col=len(header)):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    widths = {
        1: 34,
        2: 36,
        3: 12,
        4: 22,
        5: 22,
        6: 36,
        7: 14,
        8: 12,
        9: 28,
        10: 28,
        11: 28,
        12: 16,
        13: 20,
        14: 18,
        15: 14,
        16: 24,
        17: 28,
        18: 16,
        19: 12,
        20: 12,
        21: 55,
    }
    for i, w in widths.items():
        out_ws.column_dimensions[get_column_letter(i)].width = w
    out_ws.freeze_panes = "A2"
    out_wb.save(FP)

    audit = {
        "ts": datetime.now().isoformat(),
        "backup": str(backup),
        "before": len(rows),
        "after": len(final_rows),
        "removed_count": len(removed),
        "merged_groups": len(merged_log),
        "rows_collapsed_via_merge": sum(m["count"] for m in merged_log) - len(merged_log),
        "fixed_count": len(fixed),
        "dup_dropped_after_merge": dup_dropped,
        "removed": removed,
        "fixed": fixed,
        "merged": merged_log,
        "file": str(FP),
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(
        "DONE before",
        audit["before"],
        "after",
        audit["after"],
        "removed",
        audit["removed_count"],
        "collapsed",
        audit["rows_collapsed_via_merge"],
    )
    print("audit", AUDIT)


if __name__ == "__main__":
    main()
