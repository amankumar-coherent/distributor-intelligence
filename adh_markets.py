"""120 Adhesives & Bonding / Technical Foam markets, GLOBAL, distributors
(user, 2026-10-30).

Built from the user's "Adhesives & Bonding, Technical Foam Expo_ CMI
Dashboard Topics.xlsx" (tabs Titles + Market Definitions: name, definition,
segmentation, exclusion). The definitions live as data in
queries/adhesives_foam_markets.json, so a wording fix there needs no code
change. Cap 400 per market, 20 markets at once, one browser per market.

COMPANY TYPE (the definitions measure sales of finished products; distributor
intelligence = the channel that resells them):

  DISTRIBUTOR -- adhesives, tapes, sealants, foams, surface chemicals,
  silanes, thermal interface materials and potting compounds: distributors,
  stockists and specialty resellers buying from the makers (Henkel, 3M, Sika,
  H.B. Fuller, Dow, Huntsman, Armacell, Rogers, Zotefoams...). For foams and
  foam parts also CONVERTERS / FABRICATORS that buy the foam from its makers
  and sell it cut, laminated or shaped -- that is how technical foam reaches
  industry. Makers of the chemistry / raw foam are excluded.

  HYBRID -- the six EQUIPMENT markets (109 mechanical surface preparation,
  110 plasma and corona treatment, 113 dispensing equipment, 114
  meter-mix-dispense, 115 hot melt dispensing, 116 bonding robots): sold
  through authorised dealers, agents and system integrators as much as
  stocking distributors. Equipment makers are excluded.

Each market's own exclusions are applied to COMPANIES: a distributor that
sells only excluded products is dropped.

Global: country by country in the ESU order (62 countries, continents mixed),
so a market that reaches 400 early still has a worldwide list. Scope rides in
ONE Google URL (budget 7,550 / hard 7,800): scopes have no semicolons (stage-4
shrink cuts semicolon lists) -- check with scripts/preflight_prompt.py --batch adh.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from esu_markets import ESU_REGIONS

_DATA = Path(__file__).resolve().parent / "queries" / "adhesives_foam_markets.json"

# Equipment markets (sheet S. No.) -> hybrid; everything else distributor.
_HYBRID = {109, 110, 113, 114, 115, 116}
# Foam markets (raw foams and converted foam parts): converters count.
_FOAM = set(range(51, 101))

# Product-evidence gate for the final build (market_batches.evidence_for):
# a verified row whose name / categories / products never match goes to
# flagged_for_review.xlsx. Kept broad (chemistry or product family words).
_EVIDENCE: dict[int, str] = {
    1: r"epox", 2: r"epox",
    3: r"acrylic|methacrylat|\bmma\b|plexus|structural adhes",
    4: r"anaerob|thread ?lock|retaining compound|thread seal|pipe seal|gasket|loctite",
    5: r"cyanoacryl|instant adhes|super ?glue|\bca glue",
    6: r"polyurethane|urethane|\bpu\b|\bpur\b",
    7: r"polyurethane|urethane|\bpu\b|\bpur\b",
    8: r"\bpur\b|polyurethane|urethane|hot ?melt|hotmelt",
    9: r"silicon|\brtv\b", 10: r"silicon",
    11: r"\bms\b|ms.?polymer|silane|\bsmp\b|hybrid polymer|polyether",
    12: r"phenol|resorcinol|\bprf\b",
    13: r"urea|melamine|\buf\b|\bmuf\b|formaldehyde|wood adhes|resin",
    14: r"hot ?melt|hotmelt|\beva\b",
    15: r"hot ?melt|hotmelt|polyolefin|\bapao\b",
    16: r"hot ?melt|hotmelt|polyamide",
    17: r"hot ?melt|hotmelt|polyester",
    18: r"hot ?melt|hotmelt|\bsis\b|\bsbs\b|sebs|block copolymer|pressure.?sensitive|\bpsa",
    19: r"glue ?stick|hot ?melt|hotmelt|glue gun",
    20: r"pvac|\bpva\b|polyvinyl acetate|white glue|wood glue|wood adhes",
    21: r"pvoh|\bpva\b|polyvinyl alcohol|water.?soluble|remoisten",
    22: r"acrylic|emulsion|water.?based|waterborne",
    23: r"\bvae\b|vinyl acetate|emulsion|water.?based",
    24: r"latex|\bsbr\b|styrene.?butadiene|carpet adhes",
    25: r"\bpud\b|polyurethane|urethane|waterborne|water.?based",
    26: r"chloroprene|neoprene|contact adhes|contact cement",
    27: r"spray|aerosol",
    28: r"starch|dextrin|corrugat|paper adhes",
    29: r"protein|casein|collagen|animal glue|hide glue|bone glue|\bsoy|gelatin",
    30: r"bio.?based|lignin|tannin|renewable|sustainab|natural adhes",
    31: r"\buv\b|light.?cur|uv.?cur|photo", 32: r"dual.?cur|\buv\b|light.?cur",
    33: r"debond|reversib|removable|temporary bond|release",
    34: r"polyimide|high.?temp",
    35: r"pressure.?sensitive|\bpsa|acrylic", 36: r"pressure.?sensitive|\bpsa|rubber",
    37: r"pressure.?sensitive|\bpsa|silicon",
    38: r"double.?sided|double.?coated|tape", 39: r"tape", 40: r"transfer|tape",
    41: r"\bvhb\b|foam tape|acrylic foam|bonding tape|mounting tape|tape",
    42: r"masking|tape", 43: r"electrical|insulat|tape", 44: r"thermal|tape",
    45: r"film|structural adhes|composite|aerospace",
    46: r"film|\bweb|heat.?activ|hot ?melt|laminat",
    47: r"conductive|silver|electronic adhes", 48: r"thermal|conductive",
    49: r"optical|\boca\b|clear adhes|display",
    50: r"underfill|die.?attach|semiconductor|electronic adhes|packaging material",
    51: r"polyurethane|\bpu\b|urethane|foam", 52: r"high.?resilien|\bhr\b|polyurethane|foam",
    53: r"memory|visco|polyurethane|foam", 54: r"reticulat|filter foam|polyurethane|foam",
    55: r"rebond|bonded foam|recycl|foam",
    56: r"rigid|polyurethane|\bpu\b|\bpur\b|insulat", 57: r"\bpir\b|polyiso|insulat",
    58: r"spray|polyurethane|\bspf\b", 59: r"foam|polyurethane|\bpu\b|aerosol|sealant",
    60: r"polyurethane|\bpu\b|urethane|foam", 61: r"integral skin|polyurethane|\bpu\b|foam",
    62: r"microcellular|cellular urethane|polyurethane|poron|foam",
    63: r"\beps\b|polystyrene|styro|thermocol|foam",
    64: r"graphite|neopor|\beps\b|polystyrene|insulat",
    65: r"\bxps\b|extruded polystyrene|styrofoam|polystyrene|insulat",
    66: r"\bepp\b|polypropylene", 67: r"polypropylene|\bpp\b|foam",
    68: r"\bxlpe\b|\bxpe\b|\bixpe\b|cross.?linked|polyethylene|foam",
    69: r"polyethylene|\bpe\b|\bepe\b|foam", 70: r"\beva\b|foam", 71: r"\bpoe\b|polyolefin|foam",
    72: r"\bpvc\b|forex|foam board|foam sheet", 73: r"\bpvc\b|divinycell|airex|core|composite",
    74: r"polyimide|solimide|foam", 75: r"phenolic|insulat", 76: r"\bpei\b|polyetherimide|ultem|foam",
    77: r"peek|foam", 78: r"\bpmi\b|rohacell|polymethacrylimide|core", 79: r"\bpet\b|core|composite",
    80: r"pvdf|fluoro|foam", 81: r"melamine|basotect|acoustic",
    82: r"\bnbr\b|nitrile|armaflex|elastomeric|rubber foam|insulat", 83: r"epdm",
    84: r"silicon|foam", 85: r"neoprene|chloroprene|\bcr\b|foam",
    86: r"\bfkm\b|fluoroelastomer|viton|foam", 87: r"tpu|infinergy|foam",
    88: r"syntactic|microsphere|buoyancy|subsea", 89: r"carbon|graphite",
    90: r"bio|\bsoy|polyurethane|foam", 91: r"\bpla\b|polylactic|biodegradable|compostable|foam",
    92: r"starch|loose.?fill|biodegradable|compostable|foam",
    93: r"bio|green pe|polyethylene|foam", 94: r"cellulose|sponge|foam",
    95: r"conductive|\bemi\b|\besd\b|shield|anti.?static",
    96: r"die.?cut|gasket|seal",
    97: r"pre.?compress|expanding|sealing tape|foam tape|compriband|illbruck|joint",
    98: r"laminat|foam", 99: r"mou?ld|foam", 100: r"thermoform|foam",
    101: r"silicon|sealant", 102: r"polyurethane|\bpu\b|urethane|sealant",
    103: r"\bms\b|ms.?polymer|silane|hybrid|sealant", 104: r"acrylic|sealant|caulk",
    105: r"polysulfide|polysulphide|thiokol|sealant",
    106: r"fire.?stop|fire.?protect|fire.?rated|passive fire|intumescent|fire seal",
    107: r"sealant|joint|caulk|construction chem",
    108: r"clean|degreas|surface (prep|treat)|pre.?treat|etch|solvent",
    109: r"abrasive|blast|sand|grind|polish|surface prep|finishing",
    110: r"plasma|corona|surface treat|flame treat",
    111: r"primer|adhesion promot", 112: r"silane|coupling agent|organosilicon|silicon",
    113: r"dispens|applicat|meter|glue|adhesive equipment",
    114: r"meter|mix|dispens|two.?component|\b2k\b",
    115: r"hot ?melt|hotmelt|melter|glue|dispens",
    116: r"robot|automat|dispens|bonding|glue|integrat",
    117: r"thermal|\btim\b|interface", 118: r"thermal|gap|pad",
    119: r"thermal|grease|paste|phase.?change|compound",
    120: r"potting|encapsul|resin|electronic",
}


def _short(name: str) -> str:
    """'Global EVA Hot Melt Adhesives Market' -> 'EVA Hot Melt Adhesives'."""
    return re.sub(r"^Global\s+|\s+Market$", "", name.strip())


def _slug(n: int, name: str) -> str:
    """'adh014_eva_hot_melt_adhesives'. Numbered (the sheet's S. No.) and cut
    to keep Chromium profile paths well under Windows' 260-char limit."""
    words = re.sub(r"[^a-z0-9]+", "_", _short(name).lower()).strip("_")
    return f"adh{n:03d}_{words[:34].rstrip('_')}"


def _kind(n: int) -> str:
    return "hybrid" if n in _HYBRID else "distributor"


def _first_sentence(text: str) -> str:
    """First sentence minus its 'The X market comprises' lead-in (the name
    is already in the scope; the URL budget is tight)."""
    first = re.split(r"(?<=[.])\s+", text.strip())[0].rstrip(".")
    return re.sub(r"^.*?\bcomprises\s+", "", first, count=1)


def _excl(text: str) -> str:
    # No semicolons in a scope (discovery's stage-4 shrink cuts them).
    return re.sub(r"\s*;\s*", ", ", text.strip()).rstrip(".")


def _who(n: int) -> str:
    if n in _HYBRID:
        return ("DISTRIBUTORS, dealers, agents or system integrators of this "
                "equipment made by others")
    if n in _FOAM:
        return ("DISTRIBUTORS, stockists or foam converters / fabricators "
                "selling these made by others")
    return "DISTRIBUTORS, stockists or specialty resellers of these made by others"


def _scope(m: dict) -> str:
    lead = "PRODUCT/SOLUTION SCOPE (required)" if m["n"] in _HYBRID else "PRODUCT SCOPE (required)"
    return (f"{lead}: {_short(m['name']).upper()} = {_first_sentence(m['definition'])}. "
            f"{_who(m['n'])}. EXCLUDE makers of these, and companies selling only: "
            f"{_excl(m['exclusion'])}.")


# verify_distributors_ai_mode wraps a rule in 597 fixed chars + the market
# name; past ~1,000 in total AI Mode returns empty answers.
_CRITERIA_FIXED = 597
_CRITERIA_MAX = 990


def _rule(m: dict) -> str:
    """Verifier KEEP/DROP lines, the exclusion list trimmed (at a comma) so
    the whole criteria block stays under _CRITERIA_MAX."""
    what = _short(m["name"])
    n = m["n"]
    if n in _HYBRID:
        keep = f"KEEP only distributors, dealers, agents or integrators of {what} made by others."
        # The verifier drops manufactures_own_products=true.
        note = "\nIntegrating others' equipment is not manufacturing."
    elif n in _FOAM:
        keep = f"KEEP only distributors, stockists or converters of {what} made by others."
        note = "\nCutting or laminating bought foam is not manufacturing."
    else:
        keep = f"KEEP only distributors, stockists or resellers of {what} made by others."
        note = ""
    head = f"{keep}\nDROP makers and their subsidiaries, sellers only of: "
    room = _CRITERIA_MAX - _CRITERIA_FIXED - len(m["name"]) - len(head) - len(note) - 1
    excl = _excl(m["exclusion"])
    while len(excl) > room and "," in excl:
        excl = excl.rsplit(",", 1)[0]
    return f"{head}{excl[:room]}.{note}"


def _load() -> list[dict]:
    return json.loads(_DATA.read_text(encoding="utf-8"))["markets"]


_ROWS = _load()

MARKETS_ADH: tuple[tuple[str, str, str, str], ...] = tuple(
    (_slug(m["n"], m["name"]), m["name"], _scope(m), _kind(m["n"])) for m in _ROWS)
if len({e[0] for e in MARKETS_ADH}) != len(MARKETS_ADH):
    raise ValueError("duplicate adh market slug")

SIBLINGS_ADH: dict[str, str] = {}

# WORLDWIDE rounds, not country by country (user, 2026-10-30 18:2x: "so many
# countries are skipped, search only global so companies of all countries
# can come"). Each round is one global search that excludes every company
# already collected; run_adh_market.configure turns "Worldwide round N" into
# a no-country-focus prompt. A market stops at 400 verified, or when 12
# rounds in a row find nothing new (the launcher's exhaustion rule).
# 2026-10-31 (user OK'd): when a worldwide round adds < 10 new companies the
# launcher (run_fire_10_browsers.switch_to_countries) marks the remaining
# worldwide rounds done and the market continues with the countries below it
# has not searched yet (the first rounds already did US, Germany, China).
# Countries: EUROPE FIRST (user, 2026-10-31 night), then the rest of ESU order.
_EUROPE = ("Germany", "United Kingdom", "France", "Italy", "Spain",
           "Netherlands", "Belgium", "Switzerland", "Austria", "Sweden",
           "Poland", "Denmark", "Ireland", "Portugal", "Czech Republic",
           "Romania", "Greece", "Hungary", "Turkey", "Russia")
ADH_REGIONS: tuple[str, ...] = (
    tuple(f"Worldwide round {i}" for i in range(1, 41))
    + _EUROPE + tuple(r for r in ESU_REGIONS if r not in _EUROPE))
assert sorted(ADH_REGIONS[40:]) == sorted(ESU_REGIONS)


def _company_type(n: int) -> str:
    if n in _HYBRID:
        return ("COMPANY TYPE IN THIS LIST - EQUIPMENT DISTRIBUTOR / DEALER: a "
                "distributor, authorised dealer, agent or system integrator that "
                "sells this equipment made by others to manufacturers. EXCLUDED "
                "companies: equipment makers and their own sales subsidiaries; "
                "companies selling only the excluded products; service-only firms.")
    if n in _FOAM:
        return ("COMPANY TYPE IN THIS LIST - DISTRIBUTOR / CONVERTER: a "
                "distributor, stockist or foam converter / fabricator that buys "
                "this foam from its makers and sells it (as supplied, or cut, "
                "laminated or shaped) to industry. EXCLUDED companies: foam and "
                "raw-material makers and their own sales subsidiaries; companies "
                "selling only the excluded products; service-only firms.")
    return ("COMPANY TYPE IN THIS LIST - DISTRIBUTOR: a distributor, stockist or "
            "specialty reseller that buys these products from their makers and "
            "sells them to industrial, construction or trade customers. EXCLUDED "
            "companies: product and raw-material makers and their own sales "
            "subsidiaries; companies selling only the excluded products; "
            "service-only firms.")


def _segmentation(text: str) -> str:
    """Multi-line bullet list -> one line: 'By Form: Paste, Liquid. By ...'."""
    parts = []
    for block in re.split(r"\n\s*\n", text.strip()):
        lines = [l.strip(" •-\t") for l in block.splitlines() if l.strip(" •-\t")]
        if lines:
            parts.append(f"{lines[0].upper()}: {', '.join(lines[1:])}." if len(lines) > 1 else lines[0])
    return " ".join(parts)


# Client-facing definition (row 2 of the final Excel): the user's own wording.
DEFINITIONS: dict[str, str] = {
    _slug(m["n"], m["name"]): " ".join((
        m["definition"].strip(),
        _segmentation(m["segmentation"]),
        "EXCLUDED from the market: " + m["exclusion"].strip(),
        _company_type(m["n"]),
    )) for m in _ROWS
}

VERIFY_RULES: dict[str, str] = {_slug(m["n"], m["name"]): _rule(m) for m in _ROWS}

EVIDENCE: dict[str, str] = {_slug(m["n"], m["name"]): _EVIDENCE[m["n"]] for m in _ROWS}
assert len(EVIDENCE) == 120 and set(_EVIDENCE) == {m["n"] for m in _ROWS}


def by_slug_adh(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_ADH:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown adh market slug: {slug}")
