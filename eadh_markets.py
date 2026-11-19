"""The same 120 Adhesives & Bonding / Technical Foam markets, EUROPE ONLY
(user, 2026-11-03): cap 250 each, 20 markets at once (one browser each).

Built from adh_markets (same scopes, company types, verifier rules and
product-evidence patterns; market names "Global ..." -> "Europe ..."). SEEDS:
every Europe-headquartered company of the delivered global final files
(output\\final\\adh\\adh distributors (status checked)) is added as-is by
scripts\\seed_eadh.py -- already verified and status-checked, so it skips every
step (verdict + status answer pre-filled) and counts toward the 250 cap.
Discovery then finds MORE Europe companies country by country: each round
searches one European country with the HQ gate set to that country (a full
Europe country list in the prompt would overflow Google's URL budget).
"""
from __future__ import annotations

import adh_markets as G

# Europe for this list: EU, UK, EFTA, Western Balkans, Ukraine, Moldova and
# Russia (Turkey counted with the Middle East, as in the CMI regions).
# Ordered by industrial size so a market that reaches 250 early still covers
# the big markets first.
EADH_REGIONS: tuple[str, ...] = (
    "Germany", "United Kingdom", "France", "Italy", "Spain", "Netherlands",
    "Belgium", "Switzerland", "Austria", "Poland", "Sweden", "Denmark",
    "Czech Republic", "Ireland", "Portugal", "Finland", "Norway", "Romania",
    "Hungary", "Greece", "Slovakia", "Slovenia", "Croatia", "Bulgaria",
    "Serbia", "Lithuania", "Latvia", "Estonia", "Luxembourg", "Ukraine",
    "Russia", "Cyprus", "Malta", "Iceland", "Bosnia and Herzegovina",
    "North Macedonia", "Albania", "Montenegro", "Moldova",
)

# HQ country names counted as Europe (seed selection + audit gate), lower case.
EUROPE_COUNTRIES: frozenset[str] = frozenset({
    "germany", "united kingdom", "uk", "england", "scotland", "wales",
    "northern ireland", "great britain", "france", "italy", "spain",
    "netherlands", "the netherlands", "holland", "belgium", "switzerland",
    "austria", "sweden", "poland", "denmark", "norway", "finland", "ireland",
    "portugal", "czech republic", "czechia", "hungary", "romania", "greece",
    "slovakia", "slovenia", "croatia", "bulgaria", "serbia", "lithuania",
    "latvia", "estonia", "luxembourg", "ukraine", "iceland", "malta",
    "cyprus", "bosnia and herzegovina", "north macedonia", "albania",
    "montenegro", "moldova", "liechtenstein", "monaco", "russia",
    "russian federation", "belarus", "kosovo",
})


def hq_country(headquarters: object) -> str:
    parts = [p.strip() for p in str(headquarters or "").split(",") if p.strip()]
    return parts[-1].lower() if parts else ""


def is_europe(headquarters: object) -> bool:
    return hq_country(headquarters) in EUROPE_COUNTRIES


def eu_slug(global_slug: str) -> str:
    return "e" + global_slug                     # adh001_... -> eadh001_...


def eu_name(global_name: str) -> str:
    return "Europe " + global_name[len("Global "):] if global_name.startswith("Global ") else "Europe " + global_name


MARKETS_EADH: tuple[tuple[str, str, str, str], ...] = tuple(
    (eu_slug(s), eu_name(n), scope, kind) for s, n, scope, kind in G.MARKETS_ADH)
GLOBAL_OF: dict[str, tuple[str, str]] = {eu_slug(s): (s, n) for s, n, _, _ in G.MARKETS_ADH}

SIBLINGS_EADH: dict[str, str] = {}

_GEO = (" GEOGRAPHY OF THIS LIST: EUROPE - companies headquartered in a European "
        "country (EU, United Kingdom, EFTA, Western Balkans, Ukraine, Moldova, "
        "Russia).")
DEFINITIONS: dict[str, str] = {eu_slug(s): G.DEFINITIONS[s] + _GEO for s, _, _, _ in G.MARKETS_ADH}
VERIFY_RULES: dict[str, str] = {eu_slug(s): G.VERIFY_RULES[s] for s, _, _, _ in G.MARKETS_ADH}
EVIDENCE: dict[str, str] = {eu_slug(s): G.EVIDENCE[s] for s, _, _, _ in G.MARKETS_ADH}


def by_slug_eadh(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_EADH:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown eadh market slug: {slug}")
