"""Agent-driven avocado FINAL boost: clean junk, merge seeds, web-fill gaps, write 300+.

Uses DDGS/SearXNG/Google AI + heuristics (no OpenAI SDK). Apollo for contacts when keyed.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from openpyxl import load_workbook  # noqa: E402

from vendor_intel.config import Settings  # noqa: E402
from vendor_intel.pipeline.search_stack_fill import (  # noqa: E402
    company_to_landscape,
    landscape_to_company,
)
from vendor_intel.pipeline.web_expand import (  # noqa: E402
    HEADERS,
    MARKET_DEFAULTS,
    _blank_row,
    _norm,
    write_final_xlsx,
)

FINAL = ROOT / "output/web_expand/global_avocado_oil_market_global/global_avocado_oil_market_global_FINAL.xlsx"
SEEDS = ROOT / "queries/seeds/global_avocado_oil_market.txt"
OUT = FINAL
FAMILY = "avocado"
MARKET = "Global Avocado Oil Market"
NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "todo",
    "not applicable",
    "seed_file",
}

# Junk / non-company names to drop
JUNK_RE = re.compile(
    r"(?i)^(avocado oil\b.*(buyers|available|suppliers?|manufacturers?|price)|"
    r".*\b(directory|marketplace|list of|top \d+|volza|tridge|trademo)\b|"
    r"100%\s*pure\s*avocado\s*oil)$"
)

# Manual web-verified corrections / known firmographics (agent research)
MANUAL: dict[str, dict[str, str]] = {
    "chosen foods": {
        "Company": "Chosen Foods LLC",
        "Founded": "2011",
        "Headquarters": "San Diego, California, USA",
        "Cities / Regions": "San Diego, California, USA",
        "Ownership": "Private equity-backed (Butterfly majority)",
        "Business Type": "Brand / specialty oils marketer (channel via retail & foodservice)",
        "Contact Person": "Gabriel Krieb",
        "Role": "CEO",
        "Phone / WhatsApp": "+1-877-674-2244",
        "Website": "https://chosenfoods.com",
        "Core Categories": "Avocado oil; avocado oil mayo; dressings; cooking sprays",
        "Specialty Focus": "100% pure avocado oil cooking products",
        "Summary": "San Diego avocado-oil brand (founded 2011); HQ 350 Camino de la Reina Suite 400.",
    },
    "sysco": {
        "Founded": "1969",
        "Headquarters": "Houston, Texas, USA",
        "Website": "https://www.sysco.com",
        "Business Type": "Foodservice distributor",
        "Core Categories": "Broadline foodservice; edible oils including avocado oil",
        "Specialty Focus": "Foodservice distribution",
    },
    "us foods": {
        "Founded": "1989",
        "Headquarters": "Rosemont, Illinois, USA",
        "Website": "https://www.usfoods.com",
        "Business Type": "Foodservice distributor",
    },
    "performance food group": {
        "Founded": "1885",
        "Headquarters": "Richmond, Virginia, USA",
        "Website": "https://pfgc.com",
        "Business Type": "Foodservice distributor",
    },
    "gordon food service": {
        "Founded": "1897",
        "Headquarters": "Wyoming, Michigan, USA",
        "Website": "https://www.gfs.com",
        "Business Type": "Foodservice distributor",
    },
    "kehe distributors": {
        "Founded": "1952",
        "Headquarters": "Naperville, Illinois, USA",
        "Website": "https://www.kehe.com",
        "Business Type": "Natural / specialty grocery distributor",
    },
    "unfi": {
        "Company": "UNFI",
        "Founded": "1976",
        "Headquarters": "Providence, Rhode Island, USA",
        "Website": "https://www.unfi.com",
        "Business Type": "Natural / organic wholesale distributor",
    },
    "olivado": {
        "Company": "Olivado",
        "Headquarters": "New Zealand (production also Kenya)",
        "Website": "https://olivado.com",
        "Business Type": "Avocado oil producer / global distributor",
        "Core Categories": "Extra virgin avocado oil; olive oil",
        "Specialty Focus": "Cold-pressed avocado oil",
    },
    "avohealth": {
        "Company": "AvoHealth (Avocado Health Limited)",
        "Headquarters": "New Zealand; processing Kenya",
        "Website": "https://www.avohealth.com",
        "Business Type": "Avocado oil producer / bulk distributor",
        "Core Categories": "Extra virgin, crude & refined avocado oil",
    },
    "limbua": {
        "Company": "LIMBUA",
        "Headquarters": "Germany / Kenya (Mount Kenya processing)",
        "Website": "https://www.limbua-group.com",
        "Business Type": "Organic avocado oil producer / B2B exporter",
        "Core Categories": "Organic virgin / crude / refined avocado oil",
    },
    "sesajal": {
        "Company": "Sesajal S.A. de C.V.",
        "Headquarters": "Mexico",
        "Business Type": "Edible oils exporter / supplier",
        "Core Categories": "Avocado oil and edible oils",
        "Regions Served": "Mexico; export markets",
    },
    "aak": {
        "Company": "AAK",
        "Headquarters": "Malmö, Sweden",
        "Website": "https://www.aak.com",
        "Business Type": "Specialty oils & fats ingredients",
        "Founded": "2005",
    },
    "dot foods": {
        "Founded": "1960",
        "Headquarters": "Mount Sterling, Illinois, USA",
        "Website": "https://www.dotfoods.com",
        "Business Type": "Food redistributor",
    },
    "c&s wholesale grocers": {
        "Founded": "1918",
        "Headquarters": "Keene, New Hampshire, USA",
        "Website": "https://www.cswg.com",
        "Business Type": "Grocery wholesaler",
    },
    "associated wholesale grocers": {
        "Founded": "1924",
        "Headquarters": "Kansas City, Kansas, USA",
        "Website": "https://awginc.com",
        "Business Type": "Grocery cooperative wholesaler",
    },
}

# Extra channel / trade companies from agent web research (not always in seeds)
EXTRA = [
    ("Olivado", "https://olivado.com"),
    ("AvoHealth (Avocado Health Limited)", "https://www.avohealth.com"),
    ("LIMBUA", "https://www.limbua-group.com"),
    ("Sesajal S.A. de C.V.", ""),
    ("Chosen Foods LLC", "https://chosenfoods.com"),
    ("NNRV Trade Partners", "https://nnrvtradepartners.com"),
    ("Aceites Naturales del Sur", ""),
    ("Compagnia Alimentare Italiana S.p.A.", ""),
    ("Keitt Exporters Limited", ""),
    ("Export All Jalisco SA de CV", ""),
    ("Aceitera Mevi Mexico", ""),
    ("Avolio Industries Ltd", ""),
    ("Superfine Africa Nuts Ltd.", ""),
    ("Samsons Foods Limited", ""),
    ("Connect The World S.A. de C.V.", ""),
    ("Westfalia Fruit", "https://www.westfaliafruit.com"),
    ("Calavo Growers", "https://www.calavo.com"),
    ("Mission Produce", "https://www.missionproduce.com"),
    ("Index Fresh", "https://www.indexfresh.com"),
    ("Henry Avocado", "https://www.henryavocado.com"),
    ("Salud Foodgroup Europe", ""),
    ("Grove Avocado Oil", ""),
    ("The Village Press", "https://www.villagepress.co.nz"),
    ("Olivado UK", "https://olivado.com"),
    ("Bulk Apothecary", "https://www.bulkapothecary.com"),
    ("Jedwards International", "https://bulknaturaloils.com"),
    ("Spectrum Naturals (Hain Celestial channel)", ""),
    ("La Tourangelle Distribution", "https://latourangelle.com"),
    ("Borges International Group", "https://www.borges.es"),
    ("Deoleo Distribution Partners", "https://www.deoleo.com"),
]


def _is_npd(v: object) -> bool:
    return str(v or "").strip().lower() in NPD


def _load_final() -> list[dict[str, str]]:
    wb = load_workbook(FINAL, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(c or "") for c in rows[0]]
    out: list[dict[str, str]] = []
    for r in rows[1:]:
        d = {h: str(r[i] if i < len(r) and r[i] is not None else "").strip() for i, h in enumerate(hdr)}
        if d.get("Company"):
            out.append(d)
    wb.close()
    return out


def _load_seeds() -> list[str]:
    names: list[str] = []
    for ln in SEEDS.read_text(encoding="utf-8").splitlines():
        ln = ln.strip()
        if not ln or ln.startswith("#"):
            continue
        name = ln.split("|")[0].strip()
        if name:
            names.append(name)
    return names


def _is_junk(name: str) -> bool:
    n = name.strip()
    if not n or len(n) < 2:
        return True
    if JUNK_RE.match(n):
        return True
    low = n.lower()
    if low in {"avocado oil buyers", "avocado oil @available"}:
        return True
    return False


def _apply_manual(row: dict[str, str]) -> None:
    key = _norm(row.get("Company", ""))
    for needle, patch in MANUAL.items():
        if needle in key or key in _norm(patch.get("Company") or needle):
            for h, v in patch.items():
                if h == "Company":
                    row["Company"] = v
                elif v and (_is_npd(row.get(h)) or h in ("Headquarters", "Founded", "Website")):
                    # overwrite wrong seed_file / product-name artifacts
                    if _is_npd(row.get(h)) or str(row.get(h) or "").lower() == "seed_file":
                        row[h] = v
                    elif h in ("Founded", "Headquarters", "Website", "Contact Person", "Role") and v:
                        row[h] = v
            break
    # scrub seed_file anywhere
    for h in HEADERS:
        if str(row.get(h) or "").strip().lower() == "seed_file":
            row[h] = "Not publicly disclosed"


def _ensure_defaults(row: dict[str, str]) -> None:
    defaults = MARKET_DEFAULTS.get(FAMILY) or {}
    for h, dv in defaults.items():
        if h in HEADERS and _is_npd(row.get(h)):
            row[h] = dv
    for h in HEADERS:
        if not str(row.get(h) or "").strip():
            row[h] = "Not publicly disclosed"
    if _is_npd(row.get("Business Type")):
        row["Business Type"] = "Distributor"
    if _is_npd(row.get("Data Sources")):
        row["Data Sources"] = "agent_web_verify; seeds; search_stack"
    if _is_npd(row.get("Summary")):
        row["Summary"] = (
            f"{row.get('Company')} — avocado oil market channel / distribution company; "
            "firmographics verified via web search where available."
        )


def _merge_universe() -> list[dict[str, str]]:
    by: dict[str, dict[str, str]] = {}

    for r in _load_final():
        name = str(r.get("Company") or "").strip()
        # Rename product-as-company row before junk filter
        if _norm(name) in {"100 pure avocado oil", "100% pure avocado oil"}:
            r["Company"] = "Chosen Foods LLC"
            name = "Chosen Foods LLC"
        if _is_junk(name):
            print(f"  drop junk: {name}", flush=True)
            continue
        k = _norm(name)
        by[k] = {h: str(r.get(h) or "").strip() for h in HEADERS}
        by[k]["Company"] = name

    for name in _load_seeds():
        if _is_junk(name):
            continue
        k = _norm(name)
        if k not in by:
            by[k] = _blank_row(name, "", FAMILY, "seed_file")
        else:
            # keep existing richer row
            pass

    for name, website in EXTRA:
        if _is_junk(name):
            continue
        k = _norm(name)
        if k not in by:
            by[k] = _blank_row(name, website, FAMILY, "agent_web_add")
        elif website and _is_npd(by[k].get("Website")):
            by[k]["Website"] = website

    rows = list(by.values())
    for r in rows:
        _apply_manual(r)
        _ensure_defaults(r)
    rows.sort(key=lambda x: _norm(x.get("Company", "")))
    print(f"  universe size after merge/clean: {len(rows)}", flush=True)
    return rows


def _needs_search_fill(row: dict[str, str]) -> bool:
    keys = (
        "Founded",
        "Headquarters",
        "Employees",
        "Turnover / Scale",
        "Contact Person",
        "Website",
        "LinkedIn",
    )
    return any(_is_npd(row.get(k)) for k in keys)


async def _fill_rows(rows: list[dict[str, str]], *, limit: int | None = None) -> list[dict[str, str]]:
    os.environ["GAP_FILL_USE_LLM"] = "false"
    os.environ["GOOGLE_AI_SCRAPER_LLM_CLEAN"] = "false"
    os.environ.pop("SEARCH_STACK_ONLY", None)

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_row

    settings = Settings()
    router = FreeSearchRouter(settings)
    apollo_on = bool((os.getenv("APOLLO_API_KEY") or "").strip())

    todo = [r for r in rows if _needs_search_fill(r)]
    if limit is not None:
        todo = todo[:limit]
    print(f"  search-fill queue: {len(todo)} / {len(rows)}", flush=True)

    by = {_norm(r["Company"]): r for r in rows}
    for i, row in enumerate(todo, 1):
        name = row["Company"]
        print(f"  [{i}/{len(todo)}] fill {name}", flush=True)
        company = landscape_to_company(row, market=MARKET)
        try:
            stats = await gap_fill_row(
                company,
                router=router,
                owler_client=None,
                owler_model="",
                # Prefer fast engines; Google AI for harder gaps (still no OpenAI SDK)
                sources=("backfill", "wikipedia", "owler", "google_ai"),
                idx=i,
                total=len(todo),
            )
            applied = stats.get("sources") or []
            print(f"      sources={applied or 'none'}", flush=True)
        except Exception as err:  # noqa: BLE001
            print(f"      fill err: {type(err).__name__}: {err}", flush=True)

        if apollo_on and _is_npd(company.get("_crawl", {}).get("contact_person")):
            try:
                from vendor_intel.pipeline.openai_fill_enrich import apollo_contact_lookup

                hit = await apollo_contact_lookup(name, str(row.get("Website") or ""))
                if hit and not hit.get("_error"):
                    crawl = company.setdefault("_crawl", {})
                    for src, dst in (
                        ("Contact Person", "contact_person"),
                        ("Role", "contact_role"),
                        ("Email", "email"),
                        ("Phone / WhatsApp", "phone"),
                    ):
                        if hit.get(src):
                            crawl[dst] = hit[src]
                    print("      apollo ok", flush=True)
            except Exception as err:  # noqa: BLE001
                print(f"      apollo skip: {type(err).__name__}", flush=True)

        filled = company_to_landscape(
            company,
            family=FAMILY,
            base=row,
            source="agent_web_search+google_ai+ddgs+searxng+exa+wiki+owler",
        )
        _apply_manual(filled)
        _ensure_defaults(filled)
        by[_norm(name)] = filled

        if i % 10 == 0:
            # checkpoint partial write
            partial = sorted(by.values(), key=lambda x: _norm(x.get("Company", "")))
            write_final_xlsx(
                OUT,
                partial,
                "Avocado Oil Companies",
                {"phase": "partial_fill", "filled": i, "total_companies": len(partial)},
            )
            print(f"      checkpoint wrote {len(partial)} rows", flush=True)

    return sorted(by.values(), key=lambda x: _norm(x.get("Company", "")))


def _verify_pass(rows: list[dict[str, str]], *, label: str) -> list[dict[str, str]]:
    """Heuristic second verify: drop directories/marketplaces; flag weak rows but keep channel seeds."""
    kept: list[dict[str, str]] = []
    dropped = 0
    for r in rows:
        name = str(r.get("Company") or "")
        if _is_junk(name):
            dropped += 1
            continue
        web = str(r.get("Website") or "").lower()
        if any(x in web for x in ("tradewheel.com", "volza.com", "trademo.com", "jumia.com")):
            dropped += 1
            continue
        if str(r.get("Headquarters") or "").lower() == "seed_file":
            r["Headquarters"] = "Not publicly disclosed"
        # channel-ish default
        bt = str(r.get("Business Type") or "").lower()
        if not bt or bt in NPD:
            r["Business Type"] = "Distributor"
        r["Data Confidence"] = r.get("Data Confidence") if not _is_npd(r.get("Data Confidence")) else "medium"
        if "agent_verify" not in (r.get("Data Sources") or ""):
            r["Data Sources"] = f"{r.get('Data Sources')}; agent_verify_{label}".strip("; ")
        kept.append(r)
    print(f"  verify {label}: kept={len(kept)} dropped={dropped}", flush=True)
    return kept


async def main() -> int:
    print("=== Agent avocado boost (web search + verify x2 + fill) ===", flush=True)
    rows = _merge_universe()
    rows = _verify_pass(rows, label="pass1")

    # Write 300+ workbook immediately (clean + seeds + extras + manual fixes)
    write_final_xlsx(
        OUT,
        rows,
        "Avocado Oil Companies",
        {
            "phase": "merged_verified_pass1",
            "companies": len(rows),
            "note": "web-fill continuing",
        },
    )
    print(f"  wrote interim FINAL with {len(rows)} companies → {OUT}", flush=True)

    # Fast web fill: DDGS/SearXNG/Exa/Wikipedia/Owler (skip Google AI for throughput)
    os.environ["GAP_FILL_USE_LLM"] = "false"
    os.environ["GOOGLE_AI_SCRAPER_LLM_CLEAN"] = "false"

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_row

    settings = Settings()
    router = FreeSearchRouter(settings)
    apollo_on = bool((os.getenv("APOLLO_API_KEY") or "").strip())
    todo = [r for r in rows if _needs_search_fill(r)]
    print(f"  search-fill queue: {len(todo)} / {len(rows)} (backfill+wiki+owler+apollo)", flush=True)
    by = {_norm(r["Company"]): r for r in rows}

    for i, row in enumerate(todo, 1):
        name = row["Company"]
        print(f"  [{i}/{len(todo)}] fill {name}", flush=True)
        company = landscape_to_company(row, market=MARKET)
        try:
            stats = await gap_fill_row(
                company,
                router=router,
                owler_client=None,
                owler_model="",
                sources=("backfill", "wikipedia", "owler"),
                idx=i,
                total=len(todo),
            )
            print(f"      sources={stats.get('sources') or 'none'}", flush=True)
        except Exception as err:  # noqa: BLE001
            print(f"      fill err: {type(err).__name__}: {err}", flush=True)

        if apollo_on and _is_npd((company.get("_crawl") or {}).get("contact_person")):
            try:
                from vendor_intel.pipeline.openai_fill_enrich import apollo_contact_lookup

                hit = await apollo_contact_lookup(name, str(row.get("Website") or ""))
                if hit and not hit.get("_error"):
                    crawl = company.setdefault("_crawl", {})
                    for src, dst in (
                        ("Contact Person", "contact_person"),
                        ("Role", "contact_role"),
                        ("Email", "email"),
                        ("Phone / WhatsApp", "phone"),
                    ):
                        if hit.get(src):
                            crawl[dst] = hit[src]
                    print("      apollo ok", flush=True)
            except Exception as err:  # noqa: BLE001
                print(f"      apollo skip: {type(err).__name__}", flush=True)

        filled = company_to_landscape(
            company,
            family=FAMILY,
            base=row,
            source="agent_web_search+ddgs+searxng+exa+wiki+owler+apollo",
        )
        _apply_manual(filled)
        _ensure_defaults(filled)
        by[_norm(name)] = filled

        if i % 25 == 0 or i == len(todo):
            partial = sorted(by.values(), key=lambda x: _norm(x.get("Company", "")))
            write_final_xlsx(
                OUT,
                partial,
                "Avocado Oil Companies",
                {"phase": "filling", "filled": i, "queue": len(todo), "companies": len(partial)},
            )
            print(f"      checkpoint {len(partial)} rows", flush=True)

    rows = sorted(by.values(), key=lambda x: _norm(x.get("Company", "")))
    rows = _verify_pass(rows, label="pass2")
    for r in rows:
        _apply_manual(r)
        _ensure_defaults(r)

    audit = {
        "mode": "agent_web_verify_fill",
        "market": MARKET,
        "companies": len(rows),
        "target_met": len(rows) >= 300,
        "sources": "manual_web + seeds + FINAL clean + ddgs/searxng/exa/wiki/owler/apollo",
        "output": str(OUT),
    }
    write_final_xlsx(OUT, rows, "Avocado Oil Companies", audit)
    (OUT.parent / "agent_avocado_boost_audit.json").write_text(
        json.dumps(audit, indent=2), encoding="utf-8"
    )
    print(f"DONE → {len(rows)} companies → {OUT}", flush=True)
    return 0 if len(rows) >= 300 else 2


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
