"""Registry of every market batch this project drives, in one place.

Before this file, six+ scripts each hardcoded
`from drone_markets import MARKETS` and `from drone_markets2 import
MARKETS2` separately, which meant adding a third batch (fire_markets.py,
25 markets) required touching every one of those files individually and
getting the exact same merge logic right in each. This registry holds
that merge once; every consuming script now loops over BATCHES instead.

Each batch entry is (runner_script, profile_prefix, markets) where
`markets` is a tuple of (slug, market_name, product_scope, kind) -- the
SAME 4-tuple shape drone_markets2.MARKETS2 and fire_markets.MARKETS_FIRE
already use. drone_markets.MARKETS (the original ten, predating the kind
column) is normalised into that same 4-tuple shape here, with kind
inferred the same way _declared_kind's old fallback did: the four named
hybrid markets, "distributor" for the rest.

ALL_MARKETS is the flat (slug, market_name, product_scope, kind) list
every consumer actually wants. GROUPS is the (slug, market_name,
runner_script, profile_prefix) shape supervisor_all.py's discovery loop
needs. Both are derived from BATCHES so there is exactly one place that
knows what a "batch" is.
"""
from __future__ import annotations

from drone_markets import MARKETS as _MARKETS1_RAW
from drone_markets2 import MARKETS2 as _MARKETS2
from fire_markets import MARKETS_FIRE as _MARKETS_FIRE
from food_markets import MARKETS_FOOD as _MARKETS_FOOD
from wound_markets import MARKETS_WOUND as _MARKETS_WOUND
from water_markets import MARKETS_BRAZIL_RO as _MARKETS_BRAZIL_RO
from chemical_markets import MARKETS_CHEMICAL as _MARKETS_CHEMICAL
from ingredient_markets import MARKETS_INGREDIENT as _MARKETS_INGREDIENT
from geo_markets import MARKETS_GEO as _MARKETS_GEO
from h2_markets import MARKETS_H2 as _MARKETS_H2
from dc_markets import MARKETS_DC as _MARKETS_DC

# drone_markets.MARKETS predates the "kind" column and is (slug, market,
# scope) only. Four of its ten are hybrid (verify_distributors_ai_mode's
# _declared_kind fallback already encoded this); the rest are the
# original distributor rule.
_BATCH1_HYBRID_SLUGS = {"drone_ai", "drone_analytics", "anti_drone"}


def _normalise_batch1() -> tuple[tuple[str, str, str, str], ...]:
    out = []
    for slug, market, scope in _MARKETS1_RAW:
        kind = "hybrid" if slug in _BATCH1_HYBRID_SLUGS else "distributor"
        out.append((slug, market, scope, kind))
    return tuple(out)


_MARKETS1 = _normalise_batch1()

# (runner_script, profile_prefix, markets)
BATCHES: tuple[tuple[str, str, tuple[tuple[str, str, str, str], ...]], ...] = (
    ("run_drone_market.py", "drone_profile_", _MARKETS1),
    ("run_drone2_market.py", "drone2_profile_", _MARKETS2),
    ("run_fire_market.py", "fire_profile_", _MARKETS_FIRE),
    ("run_food_market.py", "food_profile_", _MARKETS_FOOD),
    ("run_wound_market.py", "wound_profile_", _MARKETS_WOUND),
    ("run_brazil_market.py", "brazil_profile_", _MARKETS_BRAZIL_RO),
    ("run_chemical_market.py", "chem_profile_", _MARKETS_CHEMICAL),
    ("run_ingredient_market.py", "ingr_profile_", _MARKETS_INGREDIENT),
    ("run_geo_market.py", "geo_profile_", _MARKETS_GEO),
    ("run_h2_market.py", "h2_profile_", _MARKETS_H2),
    ("run_dc_market.py", "dc_profile_", _MARKETS_DC),
)

# Flat (slug, market, scope, kind) across every batch.
ALL_MARKETS: tuple[tuple[str, str, str, str], ...] = tuple(
    entry for _script, _prefix, markets in BATCHES for entry in markets
)

# (slug, market, runner_script, profile_prefix) -- what supervisor_all.py's
# discovery loop iterates.
GROUPS: tuple[tuple[str, str, str, str], ...] = tuple(
    (slug, market, script, prefix)
    for script, prefix, markets in BATCHES
    for slug, market, _scope, _kind in markets
)


def by_slug(slug: str) -> tuple[str, str, str, str]:
    """(slug, market_name, product_scope, kind) for any market in any batch."""
    for entry in ALL_MARKETS:
        if entry[0] == slug:
            return entry
    raise KeyError(
        f"unknown market slug: {slug!r}. "
        f"Valid: {', '.join(m[0] for m in ALL_MARKETS)}"
    )


_DEFINITION_MODULES = ("dc_markets", "h2_markets","geo_markets","chemical_markets", "ingredient_markets",
                       "water_markets", "food_markets", "wound_markets", "fire_markets")


def definition_for(slug: str) -> str:
    """The client-facing MARKET DEFINITION for one market, for row 2 of the
    final Excel. A batch module may provide DEFINITIONS {slug: text} (the
    user's own wording); otherwise it is derived from the market's scope with
    the search-only instructions ("PRODUCT SCOPE (required):", "Name the ...",
    "The company must ...") removed. Before this, every final file showed the
    same generic rule sentence and no market-specific definition at all."""
    import importlib
    import re

    for name in _DEFINITION_MODULES:
        try:
            mod = importlib.import_module(name)
        except ImportError:
            continue
        text = getattr(mod, "DEFINITIONS", {}).get(slug)
        if text:
            return text
    scope = by_slug(slug)[2]
    scope = re.sub(r"^[A-Z/ ]*SCOPE \(required\):\s*", "", " ".join(scope.split()))
    sentences = re.split(r"(?<=[.])\s+", scope)
    keep = [s for s in sentences if not re.match(r"(Name the|The company must)", s)]
    return " ".join(keep).strip()


def kind_for_market_name(market_name: str) -> str:
    """The kind a market was collected under, matched by NAME.

    Needed because verify_distributors_ai_mode.py only has the market
    name at judging time (the runner strips region/suffix text before
    passing it), not the slug. Matches on the stem with any known
    naming-convention suffix removed, both directions, so "Fire Doors
    Market - India" and "Global Drone LiDAR Market" both resolve
    correctly regardless of which batch's naming convention was used.
    """
    low = market_name.strip().lower()
    for _slug, market, _scope, kind in ALL_MARKETS:
        stem = market
        for suffix in (" - Distributors Only", " - India"):
            stem = stem.replace(suffix, "")
        stem = stem.lower()
        if stem and (stem in low or low.startswith(stem)):
            return kind
    return "distributor"


def profile_prefix_for(slug: str) -> str:
    for script, prefix, markets in BATCHES:
        if any(entry[0] == slug for entry in markets):
            return prefix
    raise KeyError(f"unknown market slug: {slug!r}")


def runner_for(slug: str) -> str:
    for script, prefix, markets in BATCHES:
        if any(entry[0] == slug for entry in markets):
            return script
    raise KeyError(f"unknown market slug: {slug!r}")


def register_all() -> int:
    """Register every market's product scope and distributor kind with
    discovery.py, and export its current XLSX. Returns how many exported
    successfully.

    Registration matters even for markets that were already exported: the
    exporter reads product_scope_for()/distributor_definition_for()'s
    in-memory registry to write the banner text, and that registry is
    per-process -- a fresh subprocess (which export_all() in auto_cycle.py
    runs as, deliberately, so it can never share a browser profile lock
    with a discovery worker) starts with nothing registered.

    Previously this registration + export loop was inlined as a string of
    Python source built up line-by-line inside auto_cycle.py and executed
    via `python -c <code>` -- exactly the kind of fragile inline-code
    pattern that has repeatedly broken on escaping in this project. One
    real function, imported and called normally, replaces it.
    """
    import sys
    from pathlib import Path

    root = Path(__file__).resolve().parent
    sys.path.insert(0, str(root / "src"))
    from vendor_intel.pipeline.market_research import discovery as D
    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

    kind_map = {
        "distributor": D.DEF_PHYSICAL,
        "service": D.DEF_SERVICE,
        "hybrid": D.DEF_HYBRID,
        "software": D.DEF_SOFTWARE,
    }
    ok = 0
    for slug, market, scope, kind in ALL_MARKETS:
        D.set_product_scope(market, scope)
        D.set_distributor_definition(market, kind_map[kind])
        try:
            export_xlsx(market)
            ok += 1
        except Exception:  # noqa: BLE001 - one bad market must not stop the rest
            pass
    return ok


def export_one(slug: str) -> str:
    """Register and export ONE market. The rolling launcher exports each
    market as it finishes a round; exporting all 125 every time would have
    15 concurrent full exports racing over the same files."""
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
    from vendor_intel.pipeline.market_research import discovery as D
    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

    kind_map = {"distributor": D.DEF_PHYSICAL, "service": D.DEF_SERVICE,
                "hybrid": D.DEF_HYBRID, "software": D.DEF_SOFTWARE}
    _slug, market, scope, kind = by_slug(slug)
    D.set_product_scope(market, scope)
    D.set_distributor_definition(market, kind_map[kind])
    return str(export_xlsx(market))


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser()
    ap.add_argument("--only", help="export just this market slug")
    args = ap.parse_args()
    if args.only:
        print("exported", export_one(args.only))
    else:
        print("exported", register_all())
