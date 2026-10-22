"""U.S. Entresto Market -- ONE market, U.S. only, pharmaceutical
DISTRIBUTORS / WHOLESALERS (user, 2026-10-19).

Inputs: user's definition and exclusions, and "Report Sample_U.S.Entresto
Market_Coherent_Market_Insights.pdf" (Downloads): its Section 4 "Distributor
Landscape" and its "Pharmaceutical Distributors & Wholesalers" stakeholder
group (U.S. distribution networks, branded vs generic availability, wholesale
pricing, inventory, logistics) define the company type. Makers profiled in the
report (Novartis; generic makers Alembic, Alkem, Biocon, Dr. Reddy's, Lupin,
Intas, Cipla, Emcure, Macleods, Zydus, Torrent, Laurus, MSN, Crystal, Hetero,
Aurobindo, Novugen, Somerset, Umedica) are EXCLUDED, as are pharmacies,
repackagers/relabelers (user's choice), API suppliers and PBMs.

Searched state by state (report regions Northeast, Southeast, Midwest,
Southwest, West). Scope rides in ONE Google URL with the system prompt
(budget 7,550 / hard 7,800); measure with `python run_ent_market.py --slug
ent_distributor --region "District of Columbia" --show-prompt`.
"""
from __future__ import annotations

MARKETS_ENT: tuple[tuple[str, str, str, str], ...] = (
    (
        "ent_distributor",
        "U.S. Entresto Market",
        "PRODUCT SCOPE (required): ENTRESTO (sacubitril/valsartan) - branded "
        "ENTRESTO and generic sacubitril/valsartan finished tablets (24/26, "
        "49/51, 97/103 mg) and oral pellets for heart failure, from Novartis "
        "and generic makers such as Dr. Reddy's, Lupin, Zydus, Torrent, "
        "Alembic, Aurobindo, MSN. Pharmaceutical WHOLESALE DISTRIBUTORS that "
        "supply hospital, retail and online pharmacies across the U.S. "
        "EXCLUDE drug manufacturers and their U.S. arms, API and bulk "
        "suppliers, repackagers and relabelers, pharmacies, PBMs and GPOs. "
        "Name the generic and brand lines it distributes.",
        "distributor",
    ),
)

SIBLINGS_ENT: dict[str, str] = {}

_NORTHEAST = ("New York", "New Jersey", "Pennsylvania", "Massachusetts",
              "Connecticut", "Rhode Island", "New Hampshire", "Vermont", "Maine")
_SOUTHEAST = ("Florida", "Georgia", "North Carolina", "South Carolina",
              "Virginia", "Tennessee", "Alabama", "Mississippi", "Kentucky",
              "Louisiana", "Arkansas", "West Virginia", "Maryland", "Delaware",
              "District of Columbia")
_MIDWEST = ("Illinois", "Ohio", "Michigan", "Indiana", "Wisconsin",
            "Minnesota", "Missouri", "Iowa", "Kansas", "Nebraska",
            "North Dakota", "South Dakota")
_SOUTHWEST = ("Texas", "Arizona", "New Mexico", "Oklahoma")
_WEST = ("California", "Washington", "Oregon", "Colorado", "Utah", "Nevada",
         "Idaho", "Montana", "Wyoming", "Alaska", "Hawaii")
# Interleaved so each round of 10 browsers spans all five report regions,
# biggest states first.
_GROUPS = (_NORTHEAST, _SOUTHEAST, _MIDWEST, _SOUTHWEST, _WEST)
ENT_REGIONS: tuple[str, ...] = tuple(
    s for i in range(max(map(len, _GROUPS))) for g in _GROUPS if i < len(g) for s in (g[i],)
)

DEFINITIONS: dict[str, str] = {
    "ent_distributor": (
        "The U.S. Entresto Market covers commercially marketed finished dosage "
        "formulations (FDFs) containing the fixed-dose combination of "
        "sacubitril and valsartan, including branded ENTRESTO and generic "
        "equivalents. These prescription medicines are used for the treatment "
        "of heart failure in adult and pediatric patients, subject to "
        "applicable approved indications. The market includes finished oral "
        "pharmaceutical products, such as film-coated tablets and oral pellets "
        "(sprinkle formulations), distributed through hospital pharmacies, "
        "retail pharmacies, and online pharmacies across the United States. "
        "Market sizing focuses exclusively on revenue generated from finished "
        "pharmaceutical formulations, excluding active pharmaceutical "
        "ingredients (APIs), bulk drug substances, and intermediates. "
        "By Product Type: Branded ENTRESTO, Generic Sacubitril/Valsartan. "
        "By Dosage Form/Strength: Film-Coated Tablets (24/26 mg, 49/51 mg, "
        "97/103 mg), Oral Pellets/Sprinkle Formulations (6/6 mg, 15/16 mg). "
        "By Patient Type: Adult, Pediatric, Geriatric. By Heart Failure Type: "
        "HFrEF, HFmrEF, HFpEF. By Distribution Channel: Hospital, Retail and "
        "Online Pharmacies. By Region: Northeast, Southeast, Midwest, "
        "Southwest, West. "
        "EXCLUDED from the market: active pharmaceutical ingredients (bulk "
        "sacubitril, valsartan and sacubitril/valsartan API); pharmaceutical "
        "intermediates, raw materials and unfinished or bulk formulations; "
        "other cardiovascular medicines (standalone valsartan, other ARBs, ACE "
        "inhibitors, unrelated heart failure drugs); non-commercial, "
        "investigational and unapproved sacubitril/valsartan products. "
        "COMPANY TYPE IN THIS LIST - U.S. PHARMACEUTICAL DISTRIBUTOR / "
        "WHOLESALER: a licensed pharmaceutical wholesale distributor "
        "headquartered in the United States that buys finished branded "
        "ENTRESTO and/or generic sacubitril/valsartan from the manufacturers "
        "and supplies hospital, retail and online pharmacies - national, "
        "regional and secondary wholesalers. EXCLUDED companies: drug "
        "manufacturers (Novartis and generic makers such as Alembic, Alkem, "
        "Biocon, Dr. Reddy's, Lupin, Intas, Cipla, Emcure, Macleods, Zydus, "
        "Torrent, Laurus, MSN, Hetero, Aurobindo, Novugen, Somerset) and their "
        "U.S. subsidiaries, API / bulk / intermediate suppliers, repackagers "
        "and relabelers, hospital / retail / online pharmacies, PBMs and GPOs, "
        "and companies headquartered outside the United States."
    ),
}


def by_slug_ent(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_ENT:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown ent market slug: {slug}")

# Product-evidence gate for the final build (market_batches.evidence_for):
# a verified company whose profile never matches this goes to
# flagged_for_review.xlsx instead of the main Excel. Tested on the delivered
# list 2026-10-21 (all genuine rows match).
EVIDENCE: dict[str, str] = {
    "ent_distributor": r"pharma|drug|rx\b|prescription|generic|wholesale|medic",
}
