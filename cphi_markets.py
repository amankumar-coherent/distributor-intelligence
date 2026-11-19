"""200 CPHI India 2026 pharma markets, GLOBAL (user, 2026-11-12).

Built from the user's "CPHI Event Delhi_200 Titles Database.xlsx" (tab
Category Breakdown: 10 categories; tab CPHI 200 Titles: title, definition,
segmentation, exclusions). The rows live as data in queries/cphi_markets.json.
Cap 400 per market, ONE market at a time with 20 browsers (20 countries
searched at once), every pipeline step per market before the next starts
(FINALIZE_EACH in the launcher), 20 browsers for verify and status check.

COMPANY TYPE (user: "see market definition and segmentation ... focus on
exclusion ... mostly distributor + wholesaler"):

  DISTRIBUTOR (titles 1-190) -- distributors, WHOLESALERS, traders, importers
  and stockists that buy the market's product from its makers and resell it:
  API / excipient / intermediate / extract / manufacturing-material traders
  and distributors, and pharmaceutical wholesalers / distributors / importers
  for the finished medicines, biologics, injectables and drug-delivery
  products. Makers are excluded (distributor rule).

  SERVICE (titles 191-200, Contract Manufacturing & Specialty Services) --
  the definitions measure "fees paid to third-party specialists", so the
  companies are the contract manufacturers / CDMOs / CMOs / CROs / testing
  labs performing that service; nobody distributes a service.

Each market's own exclusion text is applied to COMPANIES (a company selling
only excluded products is dropped).

Global, country by country (India and China first - the pharma supply hubs -
then the 62-country ESU order). Scope rides in ONE Google URL (hard 7,800):
no semicolons -- check with scripts/preflight_prompt.py --batch cphi.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from esu_markets import ESU_REGIONS

_DATA = Path(__file__).resolve().parent / "queries" / "cphi_markets.json"

_SERVICE = set(range(191, 201))

_CAT_API = "Active Pharmaceutical Ingredients (APIs)"
_CAT_DOSAGE = "Finished Dosage Formulations"
_CAT_EXCIP = "Pharmaceutical Excipients"
_CAT_INTER = "Fine Chemicals & Pharmaceutical Intermediates"
_CAT_EXTRACT = "Natural Extracts & Phytopharmaceuticals"
_CAT_BIO = "Biologics & Biosimilars"
_CAT_DDS = "Specialty Drug Delivery Systems"
_CAT_STERILE = "Sterile & Injectable Pharmaceutical Products"
_CAT_MATERIAL = "Pharmaceutical Manufacturing Materials"
_CAT_SERVICE = "Contract Manufacturing & Specialty Services"
_FINISHED = {_CAT_DOSAGE, _CAT_BIO, _CAT_DDS, _CAT_STERILE}

# Product-evidence gate for the final build (market_batches.evidence_for),
# per category and deliberately broad: a verified row whose name / categories
# / products never match goes to flagged_for_review.xlsx.
_EVIDENCE_BY_CAT: dict[str, str] = {
    _CAT_API: r"\bapis?\b|active pharma|bulk drug|drug substance|pharma|generic|molecule|medicin",
    _CAT_DOSAGE: r"pharma|medicin|drug|formulation|tablet|capsule|syrup|liquid|cream|gel|ointment|"
                 r"patch|spray|inhal|generic|healthcare|hospital|chemist",
    _CAT_EXCIP: r"excipient|pharma|ingredient|cellulose|starch|lactose|polymer|coating|chemical|"
                r"stearate|silica|talc|phosphate|mannitol|sorbitol|glycol|surfactant|preservativ|lipid",
    _CAT_INTER: r"intermediate|fine chemical|building block|reagent|catalyst|solvent|chemical|"
                r"pharma|synthes|amino acid|peptide|acetonitrile|\bdmf\b|dimethylformamide",
    _CAT_EXTRACT: r"extract|botanical|herbal|ayurved|nutraceutical|phyto|plant|natural|curcumin|"
                  r"standardi[sz]ed",
    _CAT_BIO: r"biolog|biosimilar|antibod|insulin|protein|recombinant|pharma|medicin|drug|"
              r"oncolog|hospital|specialty",
    _CAT_DDS: r"drug delivery|pharma|medicin|drug|formulation|liposom|nanoparticle|implant|"
              r"injectable|patch|microneedle|inhal|specialty",
    _CAT_STERILE: r"injectable|parenteral|sterile|infusion|\biv\b|intravenous|vial|syringe|"
                  r"pharma|medicin|drug|hospital|radiopharm|anesthe|nutrition",
    _CAT_MATERIAL: r"pharma|filter|membrane|resin|chromatograph|ethanol|alcohol|isopropyl|\bipa\b|"
                   r"solvent|carbon|nitrogen|gas|water|bioprocess|single.?use|chemical|laborator",
    _CAT_SERVICE: r"contract|cdmo|\bcmo\b|\bcro\b|outsourc|manufactur|testing|analytical|"
                  r"laborator|custom synthesis|development|services?",
}


def _short(name: str) -> str:
    """'Oncology Small Molecule APIs Market' -> 'Oncology Small Molecule APIs'."""
    return re.sub(r"^Global\s+|\s+Market$", "", name.strip())


def _slug(n: int, name: str) -> str:
    """'cphi001_oncology_small_molecule_apis'. Numbered (the sheet order) and
    cut to keep Chromium profile paths well under Windows' 260-char limit."""
    words = re.sub(r"[^a-z0-9]+", "_", _short(name).lower()).strip("_")
    return f"cphi{n:03d}_{words[:34].rstrip('_')}"


def _kind(n: int) -> str:
    return "service" if n in _SERVICE else "distributor"


def _first_sentence(text: str) -> str:
    """First sentence minus its 'The X market covers' lead-in (the name is
    already in the scope; the URL budget is tight)."""
    first = re.split(r"(?<=[.])\s+", text.strip())[0].rstrip(".")
    return re.sub(r"^The .*? market (covers|includes|measures|comprises|represents)\s+",
                  "", first, count=1, flags=re.I)


def _excl(text: str) -> str:
    # No semicolons in a scope (discovery's stage-4 shrink cuts them).
    return re.sub(r"\s*;\s*", ", ", text.strip()).rstrip(".")


def _who(m: dict) -> str:
    cat = m["category"]
    if m["n"] in _SERVICE:
        return ("CONTRACT SERVICE PROVIDERS - CDMOs, CMOs, CROs, contract manufacturers or "
                "testing laboratories that perform this service for pharmaceutical clients")
    if cat in _FINISHED:
        return ("pharmaceutical DISTRIBUTORS, WHOLESALERS, importers or stockists of these "
                "finished products made by others")
    return "DISTRIBUTORS, WHOLESALERS, traders, importers or stockists of these made by others"


def _scope(m: dict) -> str:
    lead = "SERVICE SCOPE (required)" if m["n"] in _SERVICE else "PRODUCT SCOPE (required)"
    if m["n"] in _SERVICE:
        tail = f"EXCLUDE companies doing only: {_excl(m['exclusion'])}."
    else:
        tail = (f"EXCLUDE the makers of these, and companies selling only: "
                f"{_excl(m['exclusion'])}.")
    return (f"{lead}: {_short(m['name']).upper()} = {_first_sentence(m['definition'])}. "
            f"{_who(m)}. {tail}")


# verify_distributors_ai_mode wraps a rule in 597 fixed chars + the market
# name; past ~1,000 in total AI Mode returns empty answers.
_CRITERIA_FIXED = 597
_CRITERIA_MAX = 990


def _rule(m: dict) -> str:
    """Verifier KEEP/DROP lines, the exclusion text trimmed (at a comma) so
    the whole criteria block stays under _CRITERIA_MAX."""
    what = _short(m["name"])
    if m["n"] in _SERVICE:
        keep = f"KEEP only CDMO/CMO/CRO or contract labs doing {what} for third parties."
        head = f"{keep}\nDROP in-house-only makers and brokers. OUT OF SCOPE: "
    elif m["category"] in _FINISHED:
        keep = f"KEEP only pharma distributors, wholesalers or importers of {what} made by others."
        head = f"{keep}\nDROP drug makers and their subsidiaries. OUT OF SCOPE: "
    else:
        keep = f"KEEP only distributors, wholesalers or traders of {what} made by others."
        head = f"{keep}\nDROP makers and their subsidiaries. OUT OF SCOPE: "
    room = _CRITERIA_MAX - _CRITERIA_FIXED - len(m["name"]) - len(head) - 1
    excl = _excl(m["exclusion"])
    while len(excl) > room and "," in excl:
        excl = excl.rsplit(",", 1)[0]
    return f"{head}{excl[:room]}."


def _load() -> list[dict]:
    return json.loads(_DATA.read_text(encoding="utf-8"))["markets"]


_ROWS = _load()

MARKETS_CPHI: tuple[tuple[str, str, str, str], ...] = tuple(
    (_slug(m["n"], m["name"]), m["name"], _scope(m), _kind(m["n"])) for m in _ROWS)
if len({e[0] for e in MARKETS_CPHI}) != len(MARKETS_CPHI):
    raise ValueError("duplicate cphi market slug")
assert len(MARKETS_CPHI) == 200

SIBLINGS_CPHI: dict[str, str] = {}

# India and China first (the pharma ingredient hubs), then the ESU order.
_FIRST = ("India", "China")
CPHI_REGIONS: tuple[str, ...] = _FIRST + tuple(r for r in ESU_REGIONS if r not in _FIRST)
assert sorted(CPHI_REGIONS) == sorted(ESU_REGIONS)


def _company_type(m: dict) -> str:
    if m["n"] in _SERVICE:
        return ("COMPANY TYPE IN THIS LIST - CONTRACT SERVICE PROVIDER: a CDMO, CMO, CRO, "
                "contract manufacturer or testing laboratory that performs this outsourced "
                "service for pharmaceutical and biotech clients. EXCLUDED companies: "
                "in-house-only manufacturers, brokers and marketplaces, and companies doing "
                "only the excluded activities.")
    if m["category"] in _FINISHED:
        return ("COMPANY TYPE IN THIS LIST - PHARMACEUTICAL DISTRIBUTOR / WHOLESALER: a "
                "distributor, wholesaler, importer or stockist that buys these finished "
                "products from their makers and sells them to pharmacies, hospitals, "
                "healthcare and trade customers. EXCLUDED companies: drug makers and "
                "their own sales subsidiaries, and companies selling only the excluded "
                "products.")
    return ("COMPANY TYPE IN THIS LIST - DISTRIBUTOR / WHOLESALER: a distributor, "
            "wholesaler, trader, importer or stockist that buys these products from "
            "their makers and sells them to pharmaceutical manufacturers. EXCLUDED "
            "companies: product makers and their own sales subsidiaries, and companies "
            "selling only the excluded products.")


def _segmentation(text: str) -> str:
    """Multi-line bullet list -> one line: 'BY DRUG CLASS: A, B. BY ...'."""
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
        _company_type(m),
    )) for m in _ROWS
}

VERIFY_RULES: dict[str, str] = {_slug(m["n"], m["name"]): _rule(m) for m in _ROWS}

EVIDENCE: dict[str, str] = {_slug(m["n"], m["name"]): _EVIDENCE_BY_CAT[m["category"]] for m in _ROWS}


def by_slug_cphi(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_CPHI:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown cphi market slug: {slug}")


# SIMPLE QUERY phrasing per market (discovery.set_simple_query, 2026-11-13):
# Google AI Mode now answers only plain requests; the company-type and
# exclusion rules are enforced by VERIFY_RULES in the verify step.
_CAT_HINT: dict[str, str] = {
    _CAT_API: "active pharmaceutical ingredients sold to drug manufacturers",
    _CAT_EXCIP: "pharmaceutical-grade excipients",
    _CAT_INTER: "pharmaceutical intermediates and fine chemicals",
    _CAT_EXTRACT: "standardized plant extracts for pharmaceutical and health products",
    _CAT_MATERIAL: "pharmaceutical-grade manufacturing materials",
}


def _simple_request(m: dict) -> str:
    what = _short(m["name"])
    if m["n"] in _SERVICE:
        return f"contract manufacturers, CDMOs, CROs or testing labs offering {what} services"
    if m["category"] in _FINISHED:
        return f"pharmaceutical distributors or wholesalers of {what}"
    hint = _CAT_HINT.get(m["category"], "")
    return f"distributors, wholesalers or traders of {what}" + (f" ({hint})" if hint else "")


SIMPLE_REQUEST: dict[str, str] = {_slug(m["n"], m["name"]): _simple_request(m) for m in _ROWS}
