"""Global Biopharma Buffer Market -- ONE market, worldwide, DISTRIBUTORS
(user, 2026-10-28).

Inputs: user's definition, segmentation and 16 exclusions, and "Report
Sample_Global_Biopharma Buffer Market _Coherent Market Insights.pdf"
(Downloads). The report's company section is "Supplier Intelligence" and its
profiles are buffer MAKERS (Thermo Fisher, Merck KGaA, Avantor, Sartorius,
Danaher, Lonza, Repligen, Teknova, HiMedia, Nacalai Tesque...); the user chose
"Distributors only", so those makers and their own sales subsidiaries are
EXCLUDED and the list is distributors / resellers / importers of their
biopharma buffers. Physical-goods DISTRIBUTOR rule, global, cap 400, 20
browsers.

The user's exclusions are applied to COMPANIES too: a distributor whose
range is only cell culture media, WFI / water, plain acids, bases or salts,
cleaning chemicals, chromatography resins, filters, single-use bags, buffer
equipment, or PCR / ELISA / diagnostic reagents, or that sells buffers only
for food, cosmetics or agriculture, is out.

Scope rides in ONE Google URL (budget 7,550 / hard 7,800); measure with
scripts/preflight_prompt.py --batch buf.
"""
from __future__ import annotations

from caps_markets import CAPS_REGIONS

MARKETS_BUF: tuple[tuple[str, str, str, str], ...] = (
    (
        "buf_distributor",
        "Global Biopharma Buffer Market",
        "PRODUCT SCOPE (required): BIOPROCESS BUFFERS for making biologics - "
        "ready-to-use buffers, concentrates, custom buffers, bioprocess-grade "
        "Tris, HEPES, phosphate, acetate, citrate buffer salts by Thermo "
        "Fisher, Merck, Avantor, Lonza, Teknova. DISTRIBUTORS of these to "
        "biopharma, CDMOs and labs. EXCLUDE buffer makers, sellers of only "
        "media, water, plain acids/salts, cleaning chemicals, resins, filters, "
        "bags, equipment, PCR/ELISA/diagnostic reagents, food or cosmetic "
        "traders. Name the buffer brands it distributes.",
        "distributor",
    ),
)

SIBLINGS_BUF: dict[str, str] = {}

# The report's regions (North America, Europe, Asia Pacific, Latin America,
# Middle East, Africa) country by country, then the rest of the world: the
# same 62 countries as the global capsules market.
BUF_REGIONS: tuple[str, ...] = CAPS_REGIONS

DEFINITIONS: dict[str, str] = {
    "buf_distributor": (
        "The Global Biopharma Buffer Market encompasses the commercial "
        "production, supply, and sale of specialized buffering agents, buffer "
        "salts, pre-formulated buffer mixtures, buffer concentrates, and "
        "ready-to-use buffer solutions used in the development and "
        "manufacturing of biopharmaceutical products. These buffers are "
        "designed to maintain controlled pH conditions, support protein "
        "stability, and regulate critical physicochemical parameters during "
        "upstream processing, downstream purification, and biopharmaceutical "
        "formulation. They are used across manufacturing processes involving "
        "monoclonal antibodies, recombinant proteins, vaccines, antibody-drug "
        "conjugates, cell and gene therapies, and other biological "
        "therapeutics. "
        "By Product Type: Customized / Application-Specific Buffers, "
        "Ready-to-Use (RTU) Buffers, Buffer Concentrates. By Buffer Type: "
        "Phosphate Buffers, Tris Buffers, Acetate Buffers, HEPES Buffers, "
        "Others (Citrate Buffers, MES Buffers, etc.). "
        "EXCLUDED from the market: cell culture media and feeds (including "
        "those containing buffering ingredients); water for injection, "
        "purified and sterile water; standalone pH-adjustment chemicals "
        "(hydrochloric acid, sodium hydroxide, potassium hydroxide); "
        "non-buffer process salts and chemicals (sodium chloride, potassium "
        "chloride); cleaning and sanitization chemicals (CIP / COP, "
        "detergents, sanitizers); chromatography resins and columns; "
        "filtration and separation consumables; buffer preparation and "
        "management equipment; single-use equipment and empty containers; "
        "analytical and calibration buffers (pH standards, mobile phases); "
        "molecular biology and diagnostic reagents (PCR, ELISA, "
        "electrophoresis, diagnostic-kit buffers); non-buffer formulation "
        "excipients (sugars, surfactants, preservatives, stabilizers); "
        "finished biopharmaceutical products; non-biopharmaceutical buffer "
        "applications (food, cosmetics, agriculture, environmental testing, "
        "conventional chemical manufacturing); standalone buffer-related "
        "services; internal buffer preparation and intermediate "
        "transactions. "
        "COMPANY TYPE IN THIS LIST - BIOPHARMA BUFFER DISTRIBUTOR: a "
        "distributor, reseller or importer that buys biopharma / bioprocess "
        "buffers (ready-to-use solutions, concentrates, custom buffers, "
        "bioprocess-grade buffer salts) from the manufacturers (e.g. Thermo "
        "Fisher, Merck / MilliporeSigma, Avantor / J.T.Baker, Lonza, "
        "Sartorius, Repligen, Teknova, HiMedia) and resells them to "
        "biopharmaceutical companies, CDMOs, biotech start-ups and research "
        "institutes. EXCLUDED companies: buffer manufacturers and their own "
        "sales subsidiaries; distributors selling only excluded products "
        "(cell culture media, water, plain acids / bases / salts, cleaning "
        "chemicals, resins, filters, single-use bags, equipment, analytical "
        "or molecular-biology / diagnostic reagents); traders of buffers for "
        "food, cosmetics, agriculture or industrial chemistry; service-only "
        "companies."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; keep short).
VERIFY_RULES: dict[str, str] = {
    "buf_distributor": (
        "KEEP only distributors or importers of BIOPROCESS buffers (RTU "
        "buffers, concentrates, Tris/HEPES/phosphate salts for biologics "
        "manufacturing) made by others.\n"
        "DROP buffer makers and their subsidiaries, sellers of only media, "
        "water, plain chemicals, resins, filters, equipment or "
        "PCR/ELISA/diagnostic reagents, food or cosmetic traders."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for).
# Includes the main European-language words for "buffer".
EVIDENCE: dict[str, str] = {
    "buf_distributor": (
        r"buffer|\btris\b|hepes|phosphate|\bpbs\b|citrate|acetate|\bmes\b|"
        r"bioprocess|tampon|puffer|tamp[oó]n|tamp[aã]o"
    ),
}


def by_slug_buf(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_BUF:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown buf market slug: {slug}")
