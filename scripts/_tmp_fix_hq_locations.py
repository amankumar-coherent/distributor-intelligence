"""Sanitize junk Headquarters / Cities across run4 FINAL markets, re-export, push Neon."""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

BASE = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
MIRROR = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\output\run4")
SLUGS = [
    "global_smartwatch_market_global",
    "global_green_chemical_market_global",
    "global_firewall_market_global",
]

EMAIL = re.compile(r"[\w.+-]+@[\w.-]+")
PHONE = re.compile(r"\+?\d[\d\s().-]{6,}\d")
JUNK_PHRASES = (
    "focus on",
    "long-term",
    "value creation",
    "urban re",
    "we have",
    "network of",
    "strategically located",
    "depots providing",
    "our own logistics",
    "are located in",
    "innovation",
    "company profile",
    "about us",
    "leading provider",
    "our mission",
    "quality con",
)

# Curated verified HQ fixes (company key → headquarters, optional cities_regions)
CURATED: dict[str, dict[str, str]] = {
    "arpadis group": {
        "headquarters": "Antwerp, Belgium",
        "cities_regions": "Amsterdam, Netherlands; Paris, France; Singapore",
        "regions_served": "Belgium, Netherlands, France, United Kingdom, Singapore",
    },
    "barentz": {
        "headquarters": "Hoofddorp, Netherlands",
        "cities_regions": "North America, South America, Europe, Asia",
        "regions_served": "Global",
    },
    "comercial quimica masso": {
        "headquarters": "Barcelona, Spain",
        "cities_regions": "Barcelona, Spain",
        "regions_served": "Spain",
    },
    "tennants distribution": {
        "headquarters": "Manchester, United Kingdom",
        "cities_regions": "Manchester, Leeds, Glasgow, Belfast, Dublin",
        "regions_served": "United Kingdom, Ireland",
        "regional_extensions": "United Kingdom & Ireland depot network",
    },
}


def _norm_name(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def is_bad_location(value: str | None) -> bool:
    if value is None:
        return False
    s = str(value).strip()
    if not s:
        return False
    if s in {"—", "-", "N/A", "n/a", "Unknown", "unknown"}:
        return False
    if EMAIL.search(s) or PHONE.search(s):
        return True
    low = s.lower()
    if any(p in low for p in JUNK_PHRASES):
        return True
    if low.startswith("the ") and len(s.split()) > 8:
        return True
    if len(s) > 110 and s.count(",") <= 1:
        return True
    # Sentence fragment without a place-like shape
    if re.search(r"\b(we|our|have|providing|located in)\b", low) and "," not in s:
        return True
    return False


def is_plausible_hq(value: str | None) -> bool:
    if not value or not str(value).strip():
        return False
    s = str(value).strip()
    if is_bad_location(s):
        return False
    if len(s) > 120:
        return False
    # Prefer values that look like place names
    if not re.search(r"[A-Za-z]", s):
        return False
    return True


def crawl(row: dict) -> dict:
    c = row.get("_crawl")
    if not isinstance(c, dict):
        c = {}
        row["_crawl"] = c
    return c


def apply_fix(row: dict) -> list[str]:
    updates: list[str] = []
    c = crawl(row)
    name = _norm_name(str(row.get("company") or row.get("brand") or ""))
    curated = CURATED.get(name)

    if curated:
        for key, val in curated.items():
            old = (c.get(key) or "").strip()
            if old != val:
                if key == "headquarters" and old:
                    c.setdefault("headquarters_previous", old)
                c[key] = val
                updates.append(f"{key}->{val}")
        return updates

    hq = (c.get("headquarters") or "").strip()
    cities = (c.get("cities_regions") or c.get("cities") or "").strip()
    prev = (c.get("headquarters_previous") or "").strip()

    if is_bad_location(hq):
        if is_plausible_hq(prev):
            # Prefer short city, country form when previous is a full street address
            fixed = prev
            if "," in prev:
                parts = [p.strip() for p in prev.split(",") if p.strip()]
                if len(parts) >= 2:
                    # "Desguinlei 88A, 2018 Antwerpen, Belgium" → "Antwerp, Belgium"
                    country = parts[-1]
                    city_part = parts[-2]
                    city_part = re.sub(r"^\d+\s+", "", city_part).strip()
                    fixed = f"{city_part}, {country}"
            c["headquarters_previous"] = hq
            c["headquarters"] = fixed
            updates.append(f"hq_restored:{fixed}")
        else:
            c["headquarters_previous"] = hq
            c["headquarters"] = ""
            updates.append("hq_cleared")

    if is_bad_location(cities):
        # Prefer regions_served / regional_extensions if clean
        for alt_key in ("regions_served", "regional_extensions"):
            alt = (c.get(alt_key) or "").strip()
            if alt and not is_bad_location(alt) and alt.lower() not in {"global", "worldwide"}:
                c["cities_regions"] = alt
                updates.append(f"cities_from_{alt_key}:{alt[:60]}")
                break
        else:
            c["cities_regions"] = ""
            updates.append("cities_cleared")
        if "cities" in c and is_bad_location(str(c.get("cities") or "")):
            c["cities"] = c.get("cities_regions") or ""

    # regional_extensions sometimes copied from junk HQ
    for key in ("regional_extensions", "regions_served"):
        val = (c.get(key) or "").strip()
        if is_bad_location(val):
            c[key] = ""
            updates.append(f"{key}_cleared")

    return updates


def load_main(market_dir: Path) -> tuple[Path, dict]:
    from export_founded_hq_v2 import find_main_json

    path = find_main_json(market_dir)
    assert path is not None
    return path, json.loads(path.read_text(encoding="utf-8"))


def rebuild_neon(market_dir: Path, data: dict) -> Path:
    from export_founded_hq_v2 import has_founded, has_headquarters

    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    channel_mode = False
    try:
        from vendor_intel.funnel.channel_partner_types import channel_partner_mode_enabled

        channel_mode = channel_partner_mode_enabled(data.get("query_context") or {})
    except Exception:
        pass
    keep = "is_analyst_seed" if channel_mode else "is_seed"
    filtered = [r for r in rows if (has_founded(r) and has_headquarters(r)) or r.get(keep)]
    try:
        from vendor_intel.funnel.channel_partner_types import is_production_exportable

        if channel_mode:
            filtered = [r for r in filtered if is_production_exportable(r)]
    except Exception:
        pass

    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    country = str(ctx.get("country") or "global")
    out = {
        "query": market,
        "query_context": {"industry": market, "country": country},
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ FINAL set for Neon / Distributor Intelligence",
    }
    path = market_dir / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def export_excels(market_dir: Path) -> None:
    from export_founded_hq_v2 import export_one
    from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx

    stats = export_one(market_dir, suffix="founded_and_hq_v2")
    print(f"  founded_hq: {stats}")

    neon_path = market_dir / "_neon_filtered_export.json"
    neon = json.loads(neon_path.read_text(encoding="utf-8"))
    market = neon.get("query") or market_dir.name
    final_path = market_dir / f"{market_dir.name.replace('_global','')}_FINAL.xlsx"
    # Prefer stable FINAL naming used previously
    candidates = list(market_dir.glob("*_FINAL.xlsx"))
    out = candidates[0] if candidates else market_dir / f"{market.replace(' ', '_').lower()}_FINAL.xlsx"
    # Match existing FINAL filenames
    for p in market_dir.glob("*_FINAL.xlsx"):
        out = p
        break
    try:
        export_fmcg_xlsx(neon, out, market_type="fmcg")
        print(f"  FINAL: {out.name} ({len(neon.get('relevant_companies') or [])} rows)")
    except PermissionError:
        alt = out.with_name(out.stem + "_HQ_FIXED.xlsx")
        export_fmcg_xlsx(neon, alt, market_type="fmcg")
        print(f"  FINAL locked — wrote {alt.name}")


def mirror_copy(market_dir: Path) -> None:
    dest = MIRROR / market_dir.name
    dest.mkdir(parents=True, exist_ok=True)
    for name in (
        f"{market_dir.name}.json",
        "_neon_filtered_export.json",
    ):
        src = market_dir / name
        if src.is_file():
            try:
                (dest / name).write_bytes(src.read_bytes())
            except Exception as exc:
                print(f"  mirror skip {name}: {exc}")


def main() -> None:
    push = "--push" in sys.argv
    all_fixes: list[str] = []

    for slug in SLUGS:
        market_dir = BASE / slug
        print(f"\n=== {slug} ===")
        json_path, data = load_main(market_dir)
        rows = data.get("relevant_companies") or []
        fixed_n = 0
        for row in rows:
            if not isinstance(row, dict):
                continue
            ups = apply_fix(row)
            if ups:
                fixed_n += 1
                name = row.get("company") or row.get("brand")
                print(f"  FIXED {name}: {', '.join(ups)}")
                all_fixes.append(f"{slug}|{name}")

        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        neon_path = rebuild_neon(market_dir, data)
        print(f"  neon rows: {len(json.loads(neon_path.read_text(encoding='utf-8')).get('relevant_companies') or [])}")
        export_excels(market_dir)
        mirror_copy(market_dir)
        print(f"  companies touched: {fixed_n}")

    if push:
        from vendor_intel.neon_push import push_json_to_neon

        for slug in SLUGS:
            path = BASE / slug / "_neon_filtered_export.json"
            print(f"\nPushing {slug}...")
            stats = push_json_to_neon(path, merge_companies=False)
            print(json.dumps(stats, indent=2, default=str))

    print(f"\nDone. Fixed companies: {len(all_fixes)}")


if __name__ == "__main__":
    main()
