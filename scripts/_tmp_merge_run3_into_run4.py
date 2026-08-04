"""Merge eligible run3-only companies into run4, then verify/fill columns and re-export."""
from __future__ import annotations

import json
import re
import shutil
import sys
from copy import deepcopy
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

# Reuse column fixer
sys.path.insert(0, str(Path(__file__).resolve().parent))
from _tmp_fix_brand_cols import (  # noqa: E402
    MARKETS as MARKET_TYPES,
    apply_excel,
    apply_json as apply_brand_json,
    mirror_dir,
)

ROOT3 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run3")
ROOT4 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
MIRROR = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\output\run4")

SKIP = {
    "garmin",
    "oemwatchco",
    "sunlitesystems",
    "sunlite systems",
    "xiaomistores",
    "getdistributors",
    "dafitapp",
    "wearfitpro",
    "zdnet",
    "testequity",
    "accenture",
    "accenture plc",
    "evonik industries",
    "evonik",
    "sabic",
    "solvay",
    "trishulent india",
    "trilliumchemicals",
    "greenchemicaldesign",
    "skytechgreen",
    "smartchemical",
    "chemidyes solvents",
    "chemidyes & solvents",
    "green chemical technologies",
    "icc industries",
    "gocomet",
    "iaam",
    "our partners",
    "fortinet, inc.",
    "fortinet",
    "mcafee",
    "qualys",
    "trellix",
    "zonealarm",
    "huntress",
    "redseal",
    "nordlayer",
    "comodo",
    "ipfire",
    "expertinsights",
    "firewalldealer",
    "itmonteur",
    "wroffy",
    "firewall",
    "firewalla",
    "next generation firewall service",
    "software firewall platinum support",
    "partner",
    "verizon communications",
    "encse",
    "ariacybersecurity",
    "ivanti",
    "deloitte touche tohmatsu limited",
    "deloitte",
    "support",
    "offsec",
    "fastwebhost",
    "mindhives llc",
    "mindhives",
    "tradees",
    "hik partner",
}

ALIASES = {
    "xiaomistores": "cover company",
    "synnex": "td synnex",
    "oqema": "oqema group",
    "thegreenchemicalstore": "green chemical store",
    "tech data": "td synnex",
}


def norm(name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", " ", (name or "").lower()).strip()
    return re.sub(r"\s+", " ", s)


def should_skip(name: str) -> bool:
    k = norm(name)
    if k in SKIP:
        return True
    for s in SKIP:
        if len(s) >= 5 and s in k:
            return True
    return False


def already_in_run4(name: str, keys4: set[str]) -> bool:
    k = norm(name)
    if k in keys4:
        return True
    alias = ALIASES.get(k)
    if alias and alias in keys4:
        return True
    # soft contains match for short aliases
    for a, b in ALIASES.items():
        if k == a and b in keys4:
            return True
        if k == b and a in keys4:
            return True
    return False


def merge_market(mid: str) -> dict:
    p3 = ROOT3 / mid / f"{mid}.json"
    p4 = ROOT4 / mid / f"{mid}.json"
    d3 = json.loads(p3.read_text(encoding="utf-8"))
    d4 = json.loads(p4.read_text(encoding="utf-8"))
    rows3 = list(d3.get("relevant_companies") or [])
    rows4 = list(d4.get("relevant_companies") or [])
    keys4 = {norm(r.get("company") or r.get("brand") or "") for r in rows4}

    added = []
    skipped = []
    for r in rows3:
        name = str(r.get("company") or r.get("brand") or "").strip()
        if not name:
            continue
        if already_in_run4(name, keys4):
            continue
        if should_skip(name):
            skipped.append(name)
            continue
        # Clone and mark provenance
        nr = deepcopy(r)
        nr["is_relevant"] = True
        crawl = dict(nr.get("_crawl") or {})
        crawl["merged_from_run3"] = True
        nr["_crawl"] = crawl
        # Ensure presentation enrich exists for column fixer
        if not isinstance(nr.get("_presentation_enrich"), dict):
            nr["_presentation_enrich"] = {}
        rows4.append(nr)
        keys4.add(norm(name))
        added.append(name)

    d4["relevant_companies"] = rows4
    d4["run3_merge"] = {
        "added": added,
        "skipped": skipped,
        "added_count": len(added),
        "total_after": len(rows4),
    }
    p4.write_text(json.dumps(d4, indent=2, ensure_ascii=False), encoding="utf-8")
    return {"market": mid, "added": added, "skipped": skipped, "total": len(rows4)}


def export_all(mid: str) -> dict:
    from export_founded_hq_v2 import export_one
    from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx

    md = ROOT4 / mid
    short = re.sub(r"_global$", "", mid)

    # founded+hq export
    e1 = export_one(md, suffix="founded_and_hq_v2")
    src = Path(e1["file"]) if e1.get("file") else None

    # FINAL = same as founded_and_hq_v2 content (production Founded+HQ set)
    if src and src.exists():
        final = md / f"{short}_FINAL.xlsx"
        shutil.copy2(src, final)
        cleaned = md / f"{short}_founded_and_hq_v2_cleaned.xlsx"
        shutil.copy2(src, cleaned)

    # Also refresh main market xlsx if present via full relevant list write
    data = json.loads((md / f"{mid}.json").read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or mid)
    country = str(ctx.get("country") or "global")
    result = {
        "query": market,
        "query_context": {"industry": market, "country": country},
        "relevant_companies": rows,
    }
    full_xlsx = ROOT4 / f"{mid}.xlsx"
    try:
        export_fmcg_xlsx(result, full_xlsx, market_type="fmcg")
    except Exception as exc:
        print(f"  full xlsx warn: {exc}")

    return {"founded": e1, "final": str(md / f"{short}_FINAL.xlsx")}


def main() -> None:
    print("=== MERGE run3 → run4 ===")
    for mid in MARKET_TYPES:
        stats = merge_market(mid)
        print(f"\n{mid}: +{len(stats['added'])} (total {stats['total']})")
        for n in stats["added"]:
            print(f"  + {n}")
        if stats["skipped"]:
            print(
                f"  skipped {len(stats['skipped'])}: {', '.join(stats['skipped'][:12])}"
                + ("…" if len(stats["skipped"]) > 12 else "")
            )

    print("\n=== FILL / VERIFY columns ===")
    for mid, market in MARKET_TYPES.items():
        md = ROOT4 / mid
        short = re.sub(r"_global$", "", mid)
        print(f"\n--- {mid} ({market}) ---")
        print(" JSON brands/specialty/price", apply_brand_json(md / f"{mid}.json", market))
        # contacts: light pass — ensure crawl fields used by export exist
        # (brand fixer already handles core/specialty/price/brands)

    print("\n=== EXPORT Excels ===")
    for mid, market in MARKET_TYPES.items():
        print(f"\n--- export {mid} ---")
        print(export_all(mid))
        # re-apply brand/specialty/price onto freshly exported Excels
        md = ROOT4 / mid
        short = re.sub(r"_global$", "", mid)
        for suffix in ["_FINAL.xlsx", "_founded_and_hq_v2.xlsx", "_founded_and_hq_v2_cleaned.xlsx"]:
            print(" XLSX", apply_excel(md / f"{short}{suffix}", market))
        mirror_dir(md, MIRROR / mid)

    print("\nDONE")


if __name__ == "__main__":
    main()
