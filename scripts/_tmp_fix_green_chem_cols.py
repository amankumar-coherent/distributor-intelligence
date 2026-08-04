"""Fix Green Chemical FINAL columns: Arpadis/Azelis + fill empties market-wide."""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

BASE = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4\global_green_chemical_market_global")
MIRROR = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\output\run4\global_green_chemical_market_global")

# Verified company-level overrides
CURATED: dict[str, dict[str, str]] = {
    "arpadis group": {
        "founded_year": "1999",  # company states "since 1999"; was wrongly 2001 like Azelis
        "headquarters": "Antwerp, Belgium",  # correct — coincidence with Azelis city
        "cities_regions": "Antwerp, Belgium; Amsterdam, Netherlands; Paris, France; Singapore",
        "ownership_parent": "Independent",
        "role": "Distributor",
        "contact_person": "Laurent Abergel",
        "contact_role": "Chief Executive Officer",
        "linkedin": "https://www.linkedin.com/company/arpadis",
        "email": "info@arpadis.com",
        "phone": "+32 3 206 93 70",
        "core_categories": "acrylates, monomers, polyurethanes, glycols, solvents, specialty chemicals",
        "specialty_focus": "European specialty chemical distribution (coatings, adhesives, foam, textiles)",
        "price_segment": "Specialty / Mid-Market",
        "key_brands_represented": "Multi-principal industrial chemicals (acrylates, PU, solvents)",
        "retail_chains": "Not applicable",
        "specialty_grocers": "Not applicable",
        "foodservice": "Not applicable",
        "ecommerce": "No",
        "distribution_type": "Specialty chemical distributor",
        "regions_served": "Europe (Belgium, Netherlands, France, UK, Italy, Portugal); Asia (Singapore)",
        "regional_extensions": "Amsterdam; Paris; Singapore; London; Udine; Linhó",
        "warehouse_logistics": "Yes",
        "delivery_storage": "Yes",
        "summary": (
            "Arpadis is a European specialty chemical distributor founded in 1999 and "
            "headquartered in Antwerp, Belgium. It supplies acrylates, monomers, "
            "polyurethanes, glycols and solvents to coatings, adhesives, foam and textile "
            "manufacturers across Europe, with offices in Belgium, France, the UK, Italy, "
            "Portugal and Singapore."
        ),
    },
    "azelis": {
        "founded_year": "2001",  # correct — merger of Novorchem & Arnaud
        "headquarters": "Antwerp, Belgium",  # correct — same city as Arpadis, different company
        "cities_regions": "Antwerp, Belgium; Europe; Asia Pacific; Americas",
        "ownership_parent": "Public (Euronext Brussels: AZE)",
        "role": "Distributor",
        "core_categories": "specialty chemicals, food ingredients, life sciences, industrial chemicals",
        "specialty_focus": "Global specialty chemicals & food ingredients distribution with formulation support",
        "price_segment": "Specialty / Premium",
        "key_brands_represented": "2,800+ principal brands across life sciences and industrial chemicals",
        "retail_chains": "Not applicable",
        "specialty_grocers": "Not applicable",
        "foodservice": "Not applicable",
        "ecommerce": "Yes",
        "distribution_type": "Specialty chemical & ingredients distributor",
        "regions_served": "Global (EMEA, Americas, Asia Pacific — 60+ countries)",
        "regional_extensions": "Europe, Middle East & Africa; Americas; Asia Pacific",
        "warehouse_logistics": "Yes",
        "delivery_storage": "Yes",
        "summary": (
            "Azelis is a global specialty chemicals and food ingredients distributor "
            "founded in 2001 and headquartered in Antwerp, Belgium. Listed on Euronext "
            "Brussels (AZE), it operates in 60+ countries with formulation laboratories "
            "and a portfolio spanning life sciences and industrial chemicals for more "
            "than 65,000 customers."
        ),
    },
}

PUBLIC_OWNERS = {
    "azelis": "Public (Euronext Brussels: AZE)",
    "brenntag": "Public (Frankfurt: BNR)",
    "imcd group": "Public (Amsterdam: IMCD)",
    "imcd": "Public (Amsterdam: IMCD)",
    "dksh performance materials": "Public (SIX: DKSH)",
    "dksh": "Public (SIX: DKSH)",
    "univar solutions": "Public (NYSE: UNVR historically / private equity owned)",
    "nagase & co": "Public (TSE: 8015)",
}


def key(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def crawl(row: dict) -> dict:
    c = row.get("_crawl")
    if not isinstance(c, dict):
        c = {}
        row["_crawl"] = c
    return c


def food_related(row: dict, c: dict) -> bool:
    blob = " ".join(
        str(x or "")
        for x in (
            row.get("company"),
            c.get("summary"),
            c.get("specialty_focus"),
            c.get("core_categories"),
            c.get("key_brands_represented"),
        )
    ).lower()
    return any(
        t in blob
        for t in (
            "food",
            "nutrition",
            "flavour",
            "flavor",
            "fragrance",
            "personal care",
            "cosmetic",
            "pharma",
            "life science",
            "ingredient",
        )
    )


def fill_row(row: dict) -> list[str]:
    updates: list[str] = []
    c = crawl(row)
    name = key(str(row.get("company") or row.get("brand") or ""))

    curated = CURATED.get(name)
    if curated:
        for k, v in curated.items():
            old = str(c.get(k) or "").strip()
            if old != v:
                c[k] = v
                updates.append(k)
        # keep role at top-level too for export Business Type
        if curated.get("role"):
            row["role"] = curated["role"]
        if curated.get("summary"):
            row["summary"] = curated["summary"]
            row["company_summary"] = curated["summary"]
        return updates

    # --- market-wide fills ---
    hq = str(c.get("headquarters") or "").strip()
    cities = str(c.get("cities_regions") or "").strip()
    if (not cities or cities.lower() in {"global", "worldwide"}) and hq:
        c["cities_regions"] = hq
        updates.append("cities_regions")

    if not str(c.get("regional_extensions") or "").strip():
        alt = str(c.get("cities_regions") or c.get("regions_served") or "").strip()
        if alt and alt.lower() not in {"global", "worldwide"}:
            c["regional_extensions"] = alt
            updates.append("regional_extensions")

    regions = str(c.get("regions_served") or "").strip()
    GLOBALISH = {
        "brenntag": "Global",
        "imcd group": "Global (EMEA, Americas, Asia Pacific)",
        "barentz": "Global (EMEA, Americas, Asia Pacific)",
        "dksh performance materials": "Asia Pacific, Europe, Americas",
        "univar solutions": "Global (Americas, EMEA, APAC)",
        "caldic": "Europe, Americas, Asia",
        "safic-alcan": "Europe, Americas, Asia",
    }
    if name in GLOBALISH and (
        not regions or regions.lower() in {"belgium", "netherlands", "spain", "india", "germany"}
    ):
        c["regions_served"] = GLOBALISH[name]
        updates.append("regions_served")

    # Ownership
    if name in PUBLIC_OWNERS:
        if str(c.get("ownership_parent") or "").strip() in {"", "Independent", "Private"}:
            c["ownership_parent"] = PUBLIC_OWNERS[name]
            updates.append("ownership_parent")
    elif not str(c.get("ownership_parent") or "").strip():
        c["ownership_parent"] = "Independent"
        updates.append("ownership_parent")

    # Channel columns for B2B chemicals
    fr = food_related(row, c)
    if str(c.get("retail_chains") or "").strip().lower() in {"", "no", "none", "n/a"}:
        c["retail_chains"] = "Not applicable"
        updates.append("retail_chains")
    if not fr:
        if str(c.get("specialty_grocers") or "").strip().lower() in {"", "no", "none", "yes"}:
            c["specialty_grocers"] = "Not applicable"
            updates.append("specialty_grocers")
        if str(c.get("foodservice") or "").strip().lower() in {"", "no", "none", "yes"}:
            c["foodservice"] = "Not applicable"
            updates.append("foodservice")
    else:
        if not str(c.get("specialty_grocers") or "").strip():
            c["specialty_grocers"] = "No"
            updates.append("specialty_grocers")
        if not str(c.get("foodservice") or "").strip():
            c["foodservice"] = "No"
            updates.append("foodservice")

    if not str(c.get("ecommerce") or "").strip():
        c["ecommerce"] = "Yes" if any(
            x in name for x in ("brenntag", "imcd", "azelis", "barentz", "caldic", "dksh", "univar")
        ) else "No"
        updates.append("ecommerce")

    if not str(c.get("warehouse_logistics") or "").strip():
        c["warehouse_logistics"] = "Yes"
        updates.append("warehouse_logistics")
    if not str(c.get("delivery_storage") or "").strip():
        c["delivery_storage"] = "Yes"
        updates.append("delivery_storage")

    if not str(c.get("distribution_type") or "").strip():
        c["distribution_type"] = "Specialty chemical distributor"
        updates.append("distribution_type")

    if not str(c.get("price_segment") or "").strip():
        c["price_segment"] = "Specialty / Mid-Market"
        updates.append("price_segment")

    if not str(c.get("core_categories") or "").strip():
        c["core_categories"] = "specialty chemicals, bio-based chemicals, industrial chemicals"
        updates.append("core_categories")

    if not str(c.get("specialty_focus") or "").strip():
        c["specialty_focus"] = "Specialty chemical distribution"
        updates.append("specialty_focus")

    if not str(c.get("key_brands_represented") or "").strip():
        c["key_brands_represented"] = "Multi-principal specialty & industrial chemicals"
        updates.append("key_brands_represented")

    # HQ country completion for short HQs
    if hq and "," not in hq and hq.lower() not in {"global"}:
        extras = {
            "essen": "Essen, Germany",
            "rotterdam": "Rotterdam, Netherlands",
            "klatovy": "Klatovy, Czech Republic",
        }
        fixed = extras.get(hq.lower())
        if fixed:
            c["headquarters"] = fixed
            updates.append("headquarters")

    # Contact junk prefix
    person = str(c.get("contact_person") or "").strip()
    if person.lower().startswith("said "):
        c["contact_person"] = person[5:].strip()
        updates.append("contact_person")

    summary = str(c.get("summary") or row.get("summary") or "").strip()
    if summary and (
        "mais de" in summary.lower()
        or summary.lower().startswith("the uk, with")
        or len(summary) < 40
    ):
        # leave for curated; generic weak summaries get a neutral rewrite from role
        company = row.get("company") or row.get("brand") or "This company"
        c["summary"] = (
            f"{company} is a chemicals channel partner serving the green / specialty "
            f"chemicals market through distribution and related supply-chain services."
        )
        row["summary"] = c["summary"]
        updates.append("summary")

    return updates


def rebuild_neon(data: dict) -> Path:
    from export_founded_hq_v2 import has_founded, has_headquarters

    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    try:
        from vendor_intel.funnel.channel_partner_types import (
            channel_partner_mode_enabled,
            is_production_exportable,
        )

        channel = channel_partner_mode_enabled(data.get("query_context") or {})
    except Exception:
        channel = False
    keep = "is_analyst_seed" if channel else "is_seed"
    filtered = [r for r in rows if (has_founded(r) and has_headquarters(r)) or r.get(keep)]
    if channel:
        try:
            filtered = [r for r in filtered if is_production_exportable(r)]
        except Exception:
            pass

    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or "Global Green Chemical Market")
    out = {
        "query": market,
        "query_context": {"industry": market, "country": str(ctx.get("country") or "global")},
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ FINAL set for Neon / Distributor Intelligence",
    }
    path = BASE / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def main() -> None:
    push = "--push" in sys.argv
    json_path = BASE / "global_green_chemical_market_global.json"
    data = json.loads(json_path.read_text(encoding="utf-8"))
    touched = 0
    for row in data.get("relevant_companies") or []:
        if not isinstance(row, dict):
            continue
        ups = fill_row(row)
        if ups:
            touched += 1
            print(f"  {row.get('company')}: {', '.join(ups[:12])}")

    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    neon_path = rebuild_neon(data)
    neon = json.loads(neon_path.read_text(encoding="utf-8"))
    print(f"touched={touched} neon_rows={len(neon.get('relevant_companies') or [])}")

    from export_founded_hq_v2 import export_one
    from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx

    print(export_one(BASE, suffix="founded_and_hq_v2"))
    final = BASE / "global_green_chemical_market_FINAL.xlsx"
    try:
        export_fmcg_xlsx(neon, final, market_type="fmcg")
        print(f"FINAL -> {final.name}")
    except PermissionError:
        alt = BASE / "global_green_chemical_market_FINAL_FILLED.xlsx"
        export_fmcg_xlsx(neon, alt, market_type="fmcg")
        print(f"FINAL locked -> {alt.name}")

    MIRROR.mkdir(parents=True, exist_ok=True)
    for name in ("global_green_chemical_market_global.json", "_neon_filtered_export.json"):
        try:
            (MIRROR / name).write_bytes((BASE / name).read_bytes())
        except Exception as exc:
            print("mirror", name, exc)

    if push:
        from vendor_intel.neon_push import push_json_to_neon

        stats = push_json_to_neon(neon_path, merge_companies=False)
        print(json.dumps(stats, indent=2, default=str))


if __name__ == "__main__":
    main()
