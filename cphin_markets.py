"""The 200 CPHI India 2026 pharma markets, INDIA ONLY (user, 2026-11-19:
"run these 200 market from CPHI Event Delhi_200 Titles Database.xlsx ...
read carefully definition, taxonomy/segmentation and exclusion ... run in 10
browser and only for India and do cap 200").

Same 200 rows as the global CPHI batch (queries/cphi_markets.json, checked
identical to the Excel on 2026-11-19), same company types and per-market
exclusions (cphi_markets helpers), but:
  * market names prefixed "India " (own state files; the paused global CPHI
    batch keeps its own data untouched),
  * India-HQ only (discovery.set_country_restriction) and the verifier rule
    says India-headquartered,
  * India regions: pharma-hub states first, then the other states, then
    pharma clusters and cities,
  * cap 200 per market, 10 browsers, one market at a time (Excel order),
  * LOW_YIELD_MIN in the launch script ends a market whose rounds add almost
    nothing (niche APIs have few India distributors).
"""
from __future__ import annotations

import re

import cphi_markets as C
from india_cities import INDIA_CITIES
from india_states import INDIA_STATES

_ROWS = C._ROWS


def _name(m: dict) -> str:
    return "India " + m["name"].strip()


def _slug(n: int, name: str) -> str:
    words = re.sub(r"[^a-z0-9]+", "_", C._short(name).lower()).strip("_")
    return f"cphin{n:03d}_{words[:33].rstrip('_')}"


def _scope(m: dict) -> str:
    return (C._scope(m) + " INDIA ONLY: companies headquartered in India.")


def _rule(m: dict) -> str:
    """cphi's KEEP/DROP rule with 'India-HQ' added, trimmed to the same
    criteria budget (the market name is 6 chars longer too)."""
    what = C._short(m["name"])
    if m["n"] in C._SERVICE:
        keep = f"KEEP only India-HQ CDMO/CMO/CRO or contract labs doing {what} for third parties."
        head = f"{keep}\nDROP in-house-only makers, brokers, non-India HQ. OUT OF SCOPE: "
    elif m["category"] in C._FINISHED:
        keep = f"KEEP only India-HQ pharma distributors, wholesalers or importers of {what} made by others."
        head = f"{keep}\nDROP drug makers, their subsidiaries, non-India HQ. OUT OF SCOPE: "
    else:
        keep = f"KEEP only India-HQ distributors, wholesalers or traders of {what} made by others."
        head = f"{keep}\nDROP makers, their subsidiaries, non-India HQ. OUT OF SCOPE: "
    room = C._CRITERIA_MAX - C._CRITERIA_FIXED - len(_name(m)) - len(head) - 1
    excl = C._excl(m["exclusion"])
    while len(excl) > room and "," in excl:
        excl = excl.rsplit(",", 1)[0]
    return f"{head}{excl[:max(room, 0)]}."


MARKETS_CPHIN: tuple[tuple[str, str, str, str], ...] = tuple(
    (_slug(m["n"], _name(m)), _name(m), _scope(m), C._kind(m["n"])) for m in _ROWS)
if len({e[0] for e in MARKETS_CPHIN}) != len(MARKETS_CPHIN):
    raise ValueError("duplicate cphin market slug")
assert len(MARKETS_CPHIN) == 200

SIBLINGS_CPHIN: dict[str, str] = {}

_SLUG = {m["n"]: _slug(m["n"], _name(m)) for m in _ROWS}
_OLD = {m["n"]: C._slug(m["n"], m["name"]) for m in _ROWS}

DEFINITIONS: dict[str, str] = {
    _SLUG[m["n"]]: C.DEFINITIONS[_OLD[m["n"]]].replace(
        "COMPANY TYPE IN THIS LIST",
        "GEOGRAPHY: India - companies headquartered in India. COMPANY TYPE IN THIS LIST", 1)
    for m in _ROWS
}

VERIFY_RULES: dict[str, str] = {_SLUG[m["n"]]: _rule(m) for m in _ROWS}

EVIDENCE: dict[str, str] = {_SLUG[m["n"]]: C.EVIDENCE[_OLD[m["n"]]] for m in _ROWS}

SIMPLE_REQUEST: dict[str, str] = {_SLUG[m["n"]]: C.SIMPLE_REQUEST[_OLD[m["n"]]] for m in _ROWS}

TARGETS: dict[str, int] = {s: 200 for s, *_ in MARKETS_CPHIN}

# Pharma-hub states first (Gujarat, Maharashtra, Telangana, Andhra Pradesh,
# Himachal Pradesh (Baddi), Delhi NCR, Karnataka, Tamil Nadu ...), then the
# rest; then the pharma clusters; then the 43 cities.
_HUBS = ("Gujarat", "Maharashtra", "Telangana", "Andhra Pradesh", "Himachal Pradesh",
         "Delhi NCR", "Karnataka", "Tamil Nadu", "Uttar Pradesh", "West Bengal",
         "Haryana", "Punjab", "Madhya Pradesh", "Sikkim", "Goa", "Kerala")


def _rank(region: str) -> int:
    for i, s in enumerate(_HUBS):
        if region.lower().startswith(s.lower()):
            return i
    return len(_HUBS)


_CLUSTERS = ("Ahmedabad and Vadodara pharma belt, Gujarat",
             "Ankleshwar, Vapi and Dahej chemical zones, Gujarat",
             "Baddi and Nalagarh pharma hub, Himachal Pradesh",
             "Hyderabad Genome Valley and Jeedimetla, Telangana",
             "Visakhapatnam pharma city, Andhra Pradesh",
             "Mumbai, Thane and Navi Mumbai, Maharashtra",
             "Tarapur, Boisar and Aurangabad MIDC, Maharashtra")

CPHIN_REGIONS: tuple[str, ...] = (tuple(sorted(INDIA_STATES, key=_rank)) + _CLUSTERS
                                  + tuple(INDIA_CITIES))


def by_slug_cphin(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_CPHIN:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown cphin market slug: {slug}")
