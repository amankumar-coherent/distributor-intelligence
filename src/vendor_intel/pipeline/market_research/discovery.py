"""Combined discovery+draft-research for one market research batch.

Sends ONE prompt per batch asking the generative backend directly for N
distributor/channel-partner companies of MARKET with every schema column
drafted in the same response — "Give me N distributor companies of
<market>" with the full 21-column field list spelled out — rather than a
separate list-names-then-research-each-one-by-one flow.

BACKEND: Google AI Mode (ai_mode_client), not ChatGPT. The ChatGPT-account
route hit "usage limit reached" on a free plan and stalled long runs; AI
Mode has no account, no login and no per-account quota, and answers the
same batched JSON prompt in ~10s instead of ~166s. ai_mode_client keeps
chatgpt_oauth_client's exact chat()/chat_json() signatures, so switching
back is a one-line import change if ever needed.

Per explicit user instruction, batch_runner.py no longer independently
re-verifies these fields against a real fetched source (no live website/
contact-page/LinkedIn fetch) — the draft returned here IS the final data,
used as-is. Also per explicit user instruction, contact_person/role/email/
linkedin/office_no/office_email are asked to always have a value — "best
plausible" if a real one isn't known — rather than being left blank, since
schema.finalize_row still rejects any row with a blank field and the user
wants full rows over accuracy for these specific fields. The ONE field
still required to be strictly honest (never a guess) is verdict — a
mislabeled distributor/manufacturer would corrupt the whole point of this
market's dataset, which is why that field alone keeps the "unknown" escape
hatch the others no longer have.

Continuation ("give me N more, unique from above") is handled by always
passing the full excluded-company list from state — each call is worded as
a fresh "give me N more distributor companies... do not repeat any of
these" request.
"""
from __future__ import annotations

import json
import re
from typing import Any

from vendor_intel.pipeline.market_research import ai_mode_client

_DRAFT_SYSTEM = (
    "You are a precise market-research analyst. You only list real, "
    "verifiable companies and only state facts you actually know with "
    "confidence. You never invent company names, websites, emails, phone "
    "numbers, or people's names — for anything you are not sure of, use an "
    "empty string rather than guess."
)

_FIELDS = (
    "verdict", "verdict_reason", "manufactures_own_products",
    "website", "founded", "headquarters",
    "continent_geography", "operational_presence", "ownership", "employees",
    "core_categories", "specialty_focus", "key_products_technologies",
    "online_offline_both", "contact_person", "role", "email", "linkedin",
    "office_no", "office_email",
)


# What counts as a "distributor" is NOT the same in every market, so the
# definition is swappable rather than hard-coded.
#
# PHYSICAL_GOODS is the original rule: stock it, wholesale it onward. Right
# for semiconductors, chemicals, medical devices and so on.
#
# SOFTWARE_SERVICES exists because in a cloud/software market there is
# nothing to warehouse - per explicit user instruction for CPaaS,
# "channel partners are known as distributors over here". Applying the
# physical-goods rule there would reject essentially every legitimate
# company in the market, since none of them stock inventory.
_DEF_PHYSICAL_GOODS = """Give me {count} DISTRIBUTOR companies of this market ONLY. A qualifying \
company's PRIMARY business must be buying in volume from manufacturers and \
reselling/distributing onward (wholesale / authorized / value-added / \
stocking distributor). That is what it IS, not something it also does.

THE COMPANY MUST DISTRIBUTE OTHER COMPANIES' PRODUCTS. It buys from \
third-party manufacturers and resells their branded goods. A company that \
makes what it sells is a MANUFACTURER, not a distributor here, however \
much volume it moves.

STRICTLY EXCLUDE, even if related to this market:
- MANUFACTURERS THAT ALSO DISTRIBUTE THEIR OWN PRODUCTS. This is the most \
common wrong answer: a company with its own factory, own brand, own \
formulations or own product line that also sells/distributes them \
directly. Excluded even if it calls itself a distributor, has a \
distribution arm or division, or distributes a few other brands alongside \
its own. If it manufactures ANY of what it sells in this market, EXCLUDE \
it.
- Manufacturers/OEMs/ODMs and private-label producers (a manufacturer's \
sales subsidiary, trading arm or distribution division is still the \
manufacturer).
- CHANNEL PARTNERS, resellers, VARs, dealers, agents, brokers, sales reps, \
referral partners, "solution providers" — selling/representing a vendor's \
products WITHOUT stocking/wholesaling is NOT a distributor here.
- SYSTEM INTEGRATORS — primary business is designing/building/installing/ \
integrating solutions, even if they resell hardware as part of a project.
- Pure service providers, consultancies, managed-services firms, and any \
company only weakly/indirectly connected to this market.
If unsure whether a company is a true wholesale distributor, EXCLUDE it."""

_DEF_SOFTWARE_SERVICES = """Give me {count} DISTRIBUTOR / CHANNEL PARTNER companies of this market ONLY. This is a SOFTWARE/CLOUD market, so there is no physical stock: here a distributor IS the channel partner. A qualifying company RESELLS, distributes, white-labels or aggregates another vendor's platform to its own customers - authorized reseller, master agent / master distributor, value-added reseller, white-label or wholesale aggregator, or a telecom aggregator reselling messaging/voice capacity.

THE COMPANY MUST RESELL ANOTHER VENDOR'S PLATFORM. If it owns the platform it sells, it is the vendor, not a distributor here.

STRICTLY EXCLUDE, even if related to this market:
- PLATFORM OWNERS THAT ALSO RESELL. A company with its own CPaaS/messaging platform is excluded even if it also white-labels or resells other vendors' platforms alongside it, and even if it has a "channel" or "wholesale" division. If it owns ANY platform it sells in this market, EXCLUDE it.
- The PLATFORM VENDORS themselves (a company whose own product this is, and its own sales subsidiaries) - e.g. the CPaaS provider rather than a partner reselling it.
- SYSTEM INTEGRATORS and consultancies whose primary business is building or implementing bespoke solutions rather than reselling a platform.
- Pure marketing/referral affiliates that never carry the commercial relationship with the end customer.
If unsure whether a company genuinely resells or distributes another vendor's platform, EXCLUDE it."""

# Markets whose distributors are channel partners rather than stockists.
_SOFTWARE_MARKET_HINTS = (
    "cpaas", "communication platform as a service", "saas", "paas", "iaas",
    "software", "cloud", "api", "platform as a service",
)



# --- country restriction -------------------------------------------------
#
# The geo steer PREFERS under-covered countries; it never excludes anyone,
# so a US or German distributor is still accepted by it. That is right for
# a global market and wrong for a market defined as "distributors in these
# countries only", where an out-of-scope company is not a lesser answer
# but an incorrect one.
#
# Restriction is enforced in two places on purpose:
#   * in the PROMPT, so the backend spends its answer on countries that
#     can actually be accepted;
#   * at COLLECTION, in batch_runner, because the prompt is a request and
#     the model does not always honour it. The prompt saves queries; the
#     collection gate is what makes the guarantee.
_COUNTRY_RESTRICTIONS: dict[str, tuple[str, ...]] = {}

# Known-good companies per market, used to CALIBRATE a thin market.
#
# Asked cold for Dangjo pepper ingredient suppliers, the backend replied
# "NONE KNOWN". Shown two real ones (Euro-Atlantic in Malaysia, Freshmart
# in Singapore) it then named Xiliang Agricultural, H Mart, Dream of
# Farmer and Nara Company, each with a source. The examples tell it what
# KIND of company counts, which a description alone failed to convey.
_KNOWN_EXAMPLES: dict[str, tuple[str, ...]] = {}


def set_known_examples(market: str, companies: tuple[str, ...]) -> None:
    """Register confirmed companies to show the backend as examples."""
    _KNOWN_EXAMPLES[market.strip().lower()] = companies


def known_examples_for(market: str) -> tuple[str, ...]:
    return _KNOWN_EXAMPLES.get((market or "").strip().lower(), ())


def set_country_restriction(market: str, countries: tuple[str, ...]) -> None:
    """Restrict `market` to distributors headquartered in `countries`."""
    _COUNTRY_RESTRICTIONS[market.strip().lower()] = countries


def country_restriction_for(market: str) -> tuple[str, ...]:
    """Allowed HQ countries for `market`, or () when unrestricted."""
    return _COUNTRY_RESTRICTIONS.get((market or "").strip().lower(), ())


def country_is_allowed(market: str, country: str) -> bool:
    """True when `country` is acceptable for `market`.

    An unrestricted market allows everything. A restricted market with an
    UNKNOWN country returns False: the row cannot be shown to satisfy the
    restriction, and the whole point is that out-of-scope companies are
    wrong rather than merely unhelpful.
    """
    allowed = country_restriction_for(market)
    if not allowed:
        return True
    from vendor_intel.pipeline.market_research.state import normalise_country

    key = normalise_country(country)
    if not key:
        return False
    return key in {normalise_country(c) for c in allowed}


# --- product scope -------------------------------------------------------
#
# The distributor definition pins down what a company IS (a wholesaler,
# not a reseller). It says nothing about WHAT IT SELLS, and that gap let a
# whole class of wrong companies through: for "Dry Eyes Treatment" the
# pipeline returned ophthalmic EQUIPMENT and DEVICE distributors -
# diagnostic systems, clinical lasers, surgical instruments, microsurgical
# instrumentation - every one of which is a genuine ophthalmic wholesale
# distributor, and none of which distributes a dry-eye DRUG.
#
# The scope is expressed as the search operators the user asked for,
# because a boolean line states the required intersection far more sharply
# than prose: the company must carry a product that IS the market, not
# merely a product adjacent to it.
_PRODUCT_SCOPES: tuple[tuple[tuple[str, ...], str], ...] = (
    (
        ("dangjo", "당조고추", "capsicum annuum l. cv", "quercitrin"),
        'PRODUCT SCOPE (required): must distribute DANGJO PEPPER '
        '(당조고추 / Dangjo Cheongyang), the Korean blood-sugar pepper, in '
        'ANY form: FRESH peppers (incl. KFresh Korea Dangjo), INGREDIENT '
        '(powder, extract, juice, tea), or SEEDS/seedlings for '
        'cultivation. Roles: distributor, importer, exporter, wholesaler, '
        'ingredient or seed supplier. EXCLUDE generic Korean chili '
        '(gochugaru, gochujang, kimchi) with no Dangjo; the breeders who '
        'own the cultivar; farms; restaurants; retail. Name the Dangjo '
        'product in key_products_technologies. Give the company NAME in '
        'ENGLISH/Latin script only (romanised or the company\'s own English '
        'name), never Korean or Japanese characters.'
    ),
    (
        ("dry eye", "dry eyes"),
        'PRODUCT SCOPE (required): "dry eye" AND ("pharmaceuticals" OR '
        '"ophthalmic emulsion" OR "lubricant eye drops") AND ("distributor" '
        'OR "wholesaler"). It must distribute dry-eye DRUGS: artificial '
        'tears, lubricant drops, ophthalmic emulsions/gels, cyclosporine '
        '(Restasis/Ikervis/Cequa), lifitegrast (Xiidra), '
        'perfluorohexyloctane (Miebo), varenicline (Tyrvaya), punctal '
        'plugs. EXCLUDE ophthalmic EQUIPMENT/DEVICE/INSTRUMENT '
        'distributors (diagnostics, slit lamps, tear-film analysers, '
        'meibographers, lasers, IPL, microsurgical instruments, phaco, '
        'OCT, surgical consumables), contact lens/spectacle/optical '
        'retail, and IOL/cataract/glaucoma suppliers with no dry-eye drug '
        'line. Name the dry-eye drug in key_products_technologies; if you '
        'cannot, EXCLUDE the company.',
    ),
    (
        ("wearable glucometer", "glucometer", "glucose monitor"),
        'PRODUCT SCOPE (required): ("glucometer" OR "continuous glucose '
        'monitor" OR "CGM" OR "blood glucose meter") AND ("distributor" OR '
        '"wholesaler"). Qualifying: blood glucose meters, CGM systems and '
        'sensors (Dexcom, FreeStyle Libre, Medtronic Guardian), test '
        'strips, lancets. EXCLUDE insulin-only distributors, general '
        'diabetes-drug wholesalers, fitness trackers and smartwatches '
        'without glucose sensing.',
    ),
    (
        ("silicon carbide",),
        'PRODUCT SCOPE (required): ("silicon carbide" OR "SiC") AND '
        '("MOSFET" OR "Schottky diode" OR "power module" OR "wafer" OR '
        '"abrasive") AND ("distributor" OR "wholesaler"). Qualifying: SiC '
        'MOSFETs, Schottky/JBS diodes, SiC power modules, SiC '
        'wafers/substrates/epitaxy, SiC abrasives and refractories. '
        'EXCLUDE plain-silicon semiconductor distributors, generic '
        'component brokers with no named SiC line, and GaN-only suppliers.',
    ),
    (
        ("solar rooftop", "rooftop solar"),
        'PRODUCT SCOPE (required): ("rooftop solar" OR "solar PV") AND '
        '("module" OR "panel" OR "inverter" OR "mounting structure") AND '
        '("distributor" OR "wholesaler"). Qualifying: PV modules/panels, '
        'string and micro inverters, mounting/racking, balance-of-system '
        'components. EXCLUDE EPC contractors and installers (integrators, '
        'not distributors), utility-scale-only developers, and solar '
        'thermal or water-heater suppliers.',
    ),
    (
        ("cpaas", "communication platform as a service"),
        'PRODUCT SCOPE (required): the platform resold must be a '
        'COMMUNICATIONS platform - ("CPaaS" OR "SMS API" OR "voice API" OR '
        '"messaging API" OR "A2P messaging" OR "WhatsApp Business API" OR '
        '"RCS" OR "SIP trunking" OR "video API") AND ("reseller" OR '
        '"channel partner" OR "aggregator" OR "distributor"). Qualifying: '
        'resellers, master agents, white-label partners and aggregators of '
        'Twilio, Vonage, Sinch, Infobip, MessageBird, Bandwidth, Plivo, '
        'Kaleyra, Route Mobile, Gupshup or equivalent, and A2P/SMS '
        'aggregators reselling messaging or voice capacity. EXCLUDE generic '
        'cloud, SaaS, Microsoft 365 or AWS resellers with no named '
        'communications platform, and UCaaS/contact-centre-only resellers '
        'with no programmable API offering.',
    ),
    (
        ("polyurethane",),
        'PRODUCT SCOPE (required): "polyurethane" AND ("MDI" OR "TDI" OR '
        '"polyol" OR "isocyanate" OR "foam" OR "elastomer") AND '
        '("distributor" OR "wholesaler"). Qualifying: MDI/TDI '
        'isocyanates, polyols, PU systems and prepolymers, flexible/rigid '
        'PU foam, PU elastomers, adhesives, sealants, coatings, TPU '
        'resins. EXCLUDE general chemical distributors with no named PU '
        'line, furniture/mattress retailers, and PU machinery suppliers.',
    ),
)


# Scopes registered at runtime, keyed by exact market name.
#
# The keyword table below cannot separate markets whose names share a
# word: all ten drone markets contain "drone", so a substring match would
# give Drone Analytics the Drone AI scope. Explicit registration is the
# only thing that keeps them distinct.
_REGISTERED_SCOPES: dict[str, str] = {}


def set_product_scope(market: str, scope: str) -> None:
    """Register the product scope for one market, overriding the table."""
    _REGISTERED_SCOPES[market.strip().lower()] = scope


def product_scope_for(market: str) -> str:
    """The what-it-must-SELL constraint for a market, or "" if none.

    Deliberately separate from distributor_definition_for(), which
    constrains what the company must BE. Both are needed: the failure this
    fixes was a company passing the role test perfectly while selling an
    entirely wrong product category (an ophthalmic laser distributor in a
    dry-eye DRUG market).

    Returning "" for an unlisted market is the right default - a generic
    invented scope would be worse than none, since it would filter on
    guessed product terms.
    """
    low = (market or "").lower()
    registered = _REGISTERED_SCOPES.get((market or "").strip().lower())
    if registered:
        return registered

    for needles, scope in _PRODUCT_SCOPES:
        if any(n in low for n in needles):
            return scope
    return ""


# Markets that sell hardware AND software together, where BOTH a
# stocking distributor and a channel partner are legitimate answers.
#
# Routing these to the physical rule excluded resellers, VARs and system
# integrators -- which in an AI or analytics market is most of the
# channel. Routing them to the software rule would exclude the genuine
# hardware distributors (AI compute modules, edge inference boards).
# Neither pure rule fits, so this accepts either shape.
_DEF_HYBRID_TECH = """Give me {count} DISTRIBUTOR, RESELLER or CHANNEL PARTNER companies of this market ONLY. This market sells BOTH hardware and software, so two shapes qualify: (a) a stocking/authorized DISTRIBUTOR buying hardware in volume from manufacturers and reselling it, and (b) a CHANNEL PARTNER, value-added reseller, authorized reseller or SYSTEM INTEGRATOR that resells, bundles or implements another vendor's platform, software or modules for its own customers.

THE COMPANY MUST SELL SOMEBODY ELSE'S PRODUCTS. It resells, distributes, integrates or bundles hardware, software or platforms built by third-party vendors. A company that BUILDS what it sells is a manufacturer or vendor, NOT a distributor here -- own chips, own drones, own software platform or own branded product line means EXCLUDE, even if it also resells other brands.

STRICTLY EXCLUDE: manufacturers, OEMs and their sales subsidiaries; vendors selling their own platform; pure consultancies and training firms with nothing to resell; retailers selling to consumers only; and media, associations or holding companies. If unsure whether a company actually resells a third party's product, EXCLUDE it."""


# Explicit per-market overrides, keyed by exact market name. Registered
# at runtime by the caller that knows the market's economics.
#
# Explicit rather than keyword-matched because the keyword approach is
# what broke: "Drone AI" contains no hint word, and adding "ai" to the
# hints would also match "air", "aircraft" and "repair".
_DEFINITION_OVERRIDES: dict[str, str] = {}

DEF_PHYSICAL = "physical"
DEF_SOFTWARE = "software"
DEF_HYBRID = "hybrid"
DEF_SERVICE = "service"


def set_distributor_definition(market: str, kind: str) -> None:
    """Force a market to use the physical, software or hybrid rule."""
    if kind not in (DEF_PHYSICAL, DEF_SOFTWARE, DEF_HYBRID, DEF_SERVICE):
        raise ValueError(f"unknown definition kind: {kind}")
    _DEFINITION_OVERRIDES[market.strip().lower()] = kind


# Markets that sell a SERVICE rather than a product: inspection,
# surveillance, testing/certification, training. Nobody "distributes" a
# pilot-training course, so the reseller rules find almost nobody -- the
# same structural mismatch that capped anti_drone at 86.
#
# What carries over from the other rules: the company must genuinely do
# this as its business, and a manufacturer servicing only its own
# products does not count.
_DEF_SERVICE_PROVIDER = """Give me {count} SERVICE PROVIDER companies of this market ONLY. A qualifying company PERFORMS this service commercially for third-party customers - it is contracted to do the work, or sells the course, inspection, survey or certification as its own offering.

THE SERVICE MUST BE ITS ACTUAL BUSINESS, not a capability it mentions. A company that lists this among many unrelated activities, or that does it only internally, does NOT qualify. Operators, contractors, consultancies, accredited training organisations and certification bodies all qualify when this work is what customers pay them for.

STRICTLY EXCLUDE: manufacturers whose only "service" is support, warranty or familiarisation for their own products; pure equipment sellers and distributors with no service delivery; marketplaces and directories that merely list providers; and media, associations or holding companies. If unsure whether the company actually performs the work, EXCLUDE it."""


def distributor_definition_for(market: str) -> str:
    """Pick the definition that matches the market's own economics."""
    override = _DEFINITION_OVERRIDES.get((market or "").strip().lower())
    if override == DEF_SERVICE:
        return _DEF_SERVICE_PROVIDER
    if override == DEF_HYBRID:
        return _DEF_HYBRID_TECH
    if override == DEF_SOFTWARE:
        return _DEF_SOFTWARE_SERVICES
    if override == DEF_PHYSICAL:
        return _DEF_PHYSICAL_GOODS

    low = (market or "").lower()
    if any(h in low for h in _SOFTWARE_MARKET_HINTS):
        return _DEF_SOFTWARE_SERVICES
    return _DEF_PHYSICAL_GOODS


def verdict_kind_for(market: str) -> str:
    """The raw kind string (DEF_PHYSICAL/DEF_SOFTWARE/DEF_HYBRID/DEF_SERVICE)
    a market was registered under, WITHOUT resolving it to definition
    text. Needed by the verdict-instruction builder below and by
    batch_runner's acceptance gate, both of which need to know WHICH
    kind applies, not the definition prose itself.
    """
    override = _DEFINITION_OVERRIDES.get((market or "").strip().lower())
    if override in (DEF_SERVICE, DEF_HYBRID, DEF_SOFTWARE, DEF_PHYSICAL):
        return override
    low = (market or "").lower()
    if any(h in low for h in _SOFTWARE_MARKET_HINTS):
        return DEF_SOFTWARE
    return DEF_PHYSICAL


# The verdict field's own instructions and acceptance set, previously
# hardcoded into _DRAFT_PROMPT as "Include the company ONLY if verdict is
# distributor" -- unconditionally, regardless of the market's registered
# kind. That is a real, live bug: a SERVICE market (inspection, testing,
# training, disaster response, ...) has no legitimate way to answer
# "distributor" at all, so the model was being told in one paragraph
# (distributor_definition, correctly overridden to SERVICE PROVIDER) to
# include a testing lab, then told a few lines later to exclude it unless
# verdict=="distributor". Measured live on drone_testing before this fix:
# 18% pass rate, the worst of every drone market -- consistent with a
# verdict gate that structurally could not accept what the market was
# actually trying to collect.
#
# Each entry is (instruction_text, {accepted verdict values}). The
# ACCEPTED SET is what batch_runner._build_row actually gates on, so it
# must stay in sync with the values named in instruction_text.
_VERDICT_PHYSICAL_TEXT = (
    '- verdict: "distributor" ONLY if it resells OTHER companies\' products '
    'and makes none of what it sells here; "manufacturer_distributor" if it '
    'makes ANY of its own (own factory/brand/product line) and also sells '
    'them, even alongside other brands; else "channel_partner", '
    '"manufacturer", "system_integrator", "other" or "unknown". NEVER '
    'guess. Include the company ONLY if verdict is "distributor".'
)
_VERDICT_PHYSICAL_ACCEPT = frozenset({"distributor"})

_VERDICT_HYBRID_TEXT = (
    '- verdict: "distributor" if it resells OTHER companies\' hardware or '
    'stocks it, "channel_partner" if it resells/bundles/implements another '
    'vendor\'s platform or modules as a VAR/reseller/system integrator, '
    '"manufacturer_distributor" if it makes ANY of its own product and also '
    'sells it alongside other brands; else "manufacturer", "other" or '
    '"unknown". NEVER guess. Include the company if verdict is '
    '"distributor" OR "channel_partner".'
)
_VERDICT_HYBRID_ACCEPT = frozenset({"distributor", "channel_partner"})

_VERDICT_SOFTWARE_TEXT = (
    '- verdict: "channel_partner" if it resells, white-labels or '
    'aggregates another vendor\'s platform to its own customers (this is '
    'the DISTRIBUTOR for a software/cloud market); "distributor" also '
    'accepted if it more precisely describes a reseller relationship; else '
    '"manufacturer" (the platform vendor itself), "system_integrator" '
    '(builds bespoke, does not resell a packaged platform), "other" or '
    '"unknown". NEVER guess. Include the company if verdict is '
    '"channel_partner" OR "distributor".'
)
_VERDICT_SOFTWARE_ACCEPT = frozenset({"channel_partner", "distributor"})

_VERDICT_SERVICE_TEXT = (
    '- verdict: "service_provider" ONLY if it PERFORMS this service '
    'commercially for third-party customers as its actual business - '
    'contracted to do the work, or sells the course/inspection/survey/'
    'certification as its own offering; else "manufacturer" (only '
    'services its own products), "distributor" (sells equipment, performs '
    'no service), "other" or "unknown". NEVER guess. Include the company '
    'ONLY if verdict is "service_provider".'
)
_VERDICT_SERVICE_ACCEPT = frozenset({"service_provider"})

_VERDICT_BY_KIND: dict[str, tuple[str, frozenset[str]]] = {
    DEF_PHYSICAL: (_VERDICT_PHYSICAL_TEXT, _VERDICT_PHYSICAL_ACCEPT),
    DEF_HYBRID: (_VERDICT_HYBRID_TEXT, _VERDICT_HYBRID_ACCEPT),
    DEF_SOFTWARE: (_VERDICT_SOFTWARE_TEXT, _VERDICT_SOFTWARE_ACCEPT),
    DEF_SERVICE: (_VERDICT_SERVICE_TEXT, _VERDICT_SERVICE_ACCEPT),
}


def verdict_instructions_for(market: str) -> str:
    """The `- verdict: ...` line to render into the prompt, matching
    whatever kind this market was registered under."""
    kind = verdict_kind_for(market)
    return _VERDICT_BY_KIND[kind][0]


def accepted_verdicts_for(market: str) -> frozenset[str]:
    """The verdict value(s) batch_runner._build_row must accept for this
    market, matching verdict_instructions_for's own text exactly."""
    kind = verdict_kind_for(market)
    return _VERDICT_BY_KIND[kind][1]


# The global-coverage instruction, used only for UNRESTRICTED markets.
# A market limited to a fixed country list must never see it: it names
# Latin America, Africa and Eastern Europe as hard requirements, which
# directly contradicts the restriction and spends the answer on
# companies the collection gate then rejects.
_GLOBAL_COVERAGE = (
    "GLOBAL COVERAGE IS A HARD REQUIREMENT, not a preference. "
    "Distributors exist in every region \u2014 Latin America, Africa, "
    "the Middle East, South and Southeast Asia, Eastern Europe, the "
    "Nordics, Oceania \u2014 and a list drawn only from the US, "
    "Germany, Japan, Taiwan and China is INCOMPLETE.\n\nPrefer a real "
    "regional/national distributor in an under-covered country over "
    "yet another large US or Western-European name. Do NOT return a "
    "company just because it is famous."
)

_DRAFT_PROMPT = """Market: "{market}"

{distributor_definition}
{product_scope}

{global_coverage}
{geo_steer}
Do NOT repeat any of these (incl. their subsidiaries, aliases or former \
names):
{excluded}

Fill every FIRMOGRAPHIC field below for every company (no blank/"N/A"/ \
"unknown"). The six CONTACT fields at the end are the exception and MUST \
be left "" when the real value is not known — see their own rules:

{verdict_instructions}
- verdict_reason: one sentence naming WHOSE brands it distributes, or \
for a service market, naming the service it performs \
(e.g. "authorized distributor for BASF and Covestro polyols" or \
"NABL-accredited lab performing third-party fire door testing").
- manufactures_own_products: "yes" if it makes ANY product it sells \
here (own brand/factory/product line counts), "no" if it only resells, \
"unknown" if unsure.
- website: real official domain (not a directory/social page).
- founded: 4-digit year.
- headquarters: city, country.
- continent_geography: EVERY country where it has VERIFIED PHYSICAL \
PRESENCE (HQ, offices, plants, R&D, service centres, warehouses). Format \
as REGION(Country, Country); e.g. "APAC(India, Singapore, Japan); \
EU(Germany, France)". Name every country individually. NEVER write \
"Global", "Worldwide", a bare region, or a count like "50+ countries". \
Physical presence only — not places it merely sells to.
- operational_presence: EVERY country where it meaningfully does business \
even with no office (selling, distributing, supplying, servicing, support, \
supply chain). BROADER than continent_geography: include all of those \
countries PLUS any extra. Semicolon-separated country names, same rules — \
no "Global"/region-only/counts.
- ownership: e.g. "Public Company", "Private Company", "Subsidiary of X", \
"Family-Owned", "State-Owned", "Joint Venture".
- employees: ALWAYS a RANGE like "1,000-5,000" — NEVER a bare number such \
as "23,000". Convert an approximate figure into a surrounding band (~23,000 \
becomes "20,000-25,000").
- core_categories: the PRODUCT CATEGORIES this company actually \
distributes, e.g. "Ophthalmic pharmaceuticals; OTC eye care" or "PV \
modules; inverters; mounting systems". NEVER repeat the market name or \
any part of it - "{market}" is NOT a category and will be rejected.
- specialty_focus: its specific niche, in its own terms. Again never the \
market name.
- key_products_technologies: specific products/technologies it distributes \
in "{market}".
- online_offline_both: "Online", "Offline" or "Both".
The SIX CONTACT FIELDS below are different from every field above: for \
these, an EMPTY STRING IS REQUIRED unless you actually know the real \
published value. NEVER construct, pattern-generate, infer or guess them. \
Do NOT build an email from the domain. Do NOT build a LinkedIn URL from a \
person's name. Do NOT invent a phone number from a country calling code. \
A wrong contact detail is worse than an empty one.

- contact_person: real named sales/procurement leadership contact ONLY if \
you actually know who holds that role. Otherwise "".
- role: that person's real title, or "" if contact_person is "".
- email: ONLY a real published company/sales email address you actually \
know. Otherwise "". Never derive it from the website domain.
- linkedin: ONLY the contact person's real personal profile URL you \
actually know ("https://www.linkedin.com/in/<slug>"). Otherwise "". Never \
construct one from their name.
- office_no: ONLY a real published telephone number you actually know. \
Otherwise "". Never invent a plausible number.
- office_email: ONLY a real published office/contact email you actually \
know. Otherwise "".

Reply with ONLY a JSON array, no markdown fences, no commentary, one object \
per company:
[{{"company": "", "verdict": "", "verdict_reason": "", \
"manufactures_own_products": "", "website": "", \
"founded": "", "headquarters": "", "continent_geography": "", \
"operational_presence": "", "ownership": "", "employees": "", \
"core_categories": "", "specialty_focus": "", \
"key_products_technologies": "", "online_offline_both": "", \
"contact_person": "", "role": "", "email": "", "linkedin": "", \
"office_no": "", "office_email": ""}}]"""


def _strip_echoed_prompt(raw: str) -> str:
    """Drop AI Mode's echo of the question, keeping only the answer.

    AI Mode renders the whole conversation, so the page text starts with
    the prompt read back verbatim — and this prompt ENDS with a literal
    JSON template (`[{"company": "", ...}]`). A naive find("[") therefore
    locks onto that empty template instead of the real answer and yields
    zero companies (observed live: a perfectly good 3-company answer
    parsed as 0 drafts).

    Cutting at the LAST occurrence of the template's opening is what makes
    this robust: the echo appears before the answer, so whatever follows
    the final template marker is the generated payload.
    """
    text = raw or ""

    # Find every array-of-company-objects opening. Whitespace between "["
    # and "{" must be tolerated: the echoed template is compact
    # (`[{"company"`) but a real answer is often pretty-printed
    # (`[\n  {"company"`), so a literal '[{"company"' marker matches only
    # the echo and misses the actual payload.
    starts = [m.start() for m in re.finditer(r'\[\s*\{\s*"company"', text)]

    # Keep the text FROM the LAST such opening. The real answer comes after
    # the echoed prompt, so the last opening is the payload. Slicing FROM it
    # (not past it) preserves the first record — slicing past the marker
    # dropped it, silently turning a 3-company answer into 2 (observed live).
    if starts:
        return text[starts[-1] :]

    # No array opening at all (e.g. a reply that lost its brackets): fall
    # back to the first record object so nothing is dropped.
    obj = text.find('{"company"')
    return text[obj:] if obj > 0 else text


def _clean_value(value: Any) -> str:
    """Normalize one field value from the model's JSON.

    The google-ai-scraper relay extracts the answer from rendered page
    text, which can leave real newlines INSIDE string values — observed
    live: `"company": "\\nDatumstruct\\n"`. Those are valid JSON but would
    put line breaks into company names and Excel cells, so collapse all
    internal whitespace to single spaces.
    """
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _extract_json_array(text: str) -> list[dict[str, Any]]:
    raw = _strip_echoed_prompt((text or "").strip()).strip()
    if raw.startswith("```"):
        raw = raw.strip("`")
        if raw.lower().startswith("json"):
            raw = raw[4:]
    # Google appends its own UI footer after rendered code blocks.
    raw = raw.replace("Use code with caution.", "")

    start, end = raw.find("["), raw.rfind("]")
    if start < 0 or end <= start:
        # The echo-stripped answer can start mid-array (the leading "[" was
        # part of the echoed template), so rebuild it from the objects.
        first_obj, last_obj = raw.find("{"), raw.rfind("}")
        if first_obj < 0 or last_obj <= first_obj:
            return []
        candidate = "[" + raw[first_obj : last_obj + 1] + "]"
    else:
        candidate = raw[start : end + 1]

    try:
        data = json.loads(candidate)
    except (json.JSONDecodeError, ValueError):
        # The google-ai-scraper relay reads the answer out of rendered page
        # text, which can leave REAL newlines inside string values
        # (observed live: `"company": "\nDatumstruct\n"`). A literal newline
        # inside a JSON string is illegal, so json.loads rejects the whole
        # array and every company is lost. Collapse whitespace runs that sit
        # between quotes and retry once before giving up.
        repaired = re.sub(r"\s*\n\s*", " ", candidate)
        try:
            data = json.loads(repaired)
        except (json.JSONDecodeError, ValueError):
            return []
    return [item for item in data if isinstance(item, dict)] if isinstance(data, list) else []


# Google AI Mode is driven by a GET URL (?q=<prompt>&udm=50), so the WHOLE
# prompt travels in the query string — and Google rejects a URL over roughly
# 8 KB with "400. That's an error. Your client has issued a malformed or
# illegal request." Hit live at 148 excluded companies: the prompt reached
# 9,086 chars and the URL 10,701, so every request 400'd and the run stalled.
#
# The exclusion list is the only part that grows without bound (one line per
# company found so far, heading for 500), so it is what gets capped. Trimming
# it is SAFE for correctness: it is only a hint to the model. The real
# duplicate guard is state.is_seen() / the accepted-name set in
# batch_runner._run_one_batch, which check every returned company against the
# full history regardless of what the prompt mentioned. Worst case a trimmed
# name comes back again and is silently skipped as a duplicate.
#
# Newest-first, because a repeat is most likely among companies just returned.
# Budget for the exclusion block only. The base prompt is ~4.3 KB and
# percent-encoding inflates the URL by roughly 15%, so the exclusion list
# has to stay well under 2 KB for the final URL to clear MAX_URL_CHARS
# (7,800) without build_ai_mode_url having to truncate — truncation would
# cut the JSON output template off the end and break parsing entirely.
# Verified across 0/148/300/500 excluded companies: no truncation warning.
# Lowered from 40/1500 when the geographic-steering block was added: that
# block costs ~450 chars, and at 172 companies the combined prompt produced
# an 8,398-char URL — over Google's ~7,800 ceiling, which truncates the
# JSON output template off the END and breaks parsing entirely (every
# company lost, not just the trimmed names).
#
# Trimming exclusions is the right thing to sacrifice: the list is only a
# HINT to the model, while state.is_seen() is the authoritative duplicate
# guard. Steering toward uncovered regions is worth more than 15 extra
# "do not repeat" names, because a repeat is silently skipped anyway.
# Characters of calibration examples allowed in the prompt. The whole
# prompt shares one ~7.8 KB GET URL with the scope, the geo steer and
# the exclusion list, so this has to be bounded like everything else.
_MAX_EXAMPLE_CHARS = 60

_MAX_EXCLUDED_IN_PROMPT = 25
# Character allowance for the rendered exclusion list.
#
# Lowered from 1000 to 700 when the product-scope block was added: scope,
# geo steer and exclusions share one ~7.8 KB GET URL, and a stress test
# across every market x collected-count x batch-size combination peaked
# at 7,780 of 7,800 characters -- 20 characters of headroom, which one
# longer-than-average company name would erase. Truncation is not a soft
# failure: build_ai_mode_url cuts from the END, removing the JSON reply
# template, and the model then returns nothing at all (observed live as
# "round 1: got 0 draft(s) back" on every round).
_MAX_PROMPT_CHARS = 700

# Keep this much of the URL budget in reserve.
#
# Shrinking only when the prompt is exactly at the limit left 18
# characters of headroom, which is not a margin. The cost of the margin is
# a few named exclusions; the cost of exceeding the limit is the whole
# answer, because build_ai_mode_url truncates from the END and so removes
# the trailing JSON reply template.
_URL_SAFETY_MARGIN = 250


# Regions to steer TOWARD once the obvious markets are covered. Measured
# need: at 172 companies the list was 61% US/Germany/Japan/Taiwan and had
# ZERO from Latin America, the Middle East, Africa (bar South Africa),
# Southeast Asia, or the Nordics — even though real distributors exist in
# all of them. A single "aim for broad coverage" line was too weak to
# change that, so the prompt now names the actual gaps.
# Expanded from 58 to ~110 countries. The steer works by naming countries
# with NO coverage yet, so the list is the ceiling on how long it can keep
# finding new ground: Polyurethane reached 59 distinct countries against a
# 58-country list, at which point "absent" went empty and the prompt
# stopped steering anywhere new -- one real cause of its falling yield.
#
# Second-tier markets are where the remaining distributors actually are
# once the obvious ones are collected; a market with real distributors in
# Ghana, Uzbekistan or Paraguay will never surface them if the prompt only
# ever names Brazil and Germany.
_TARGET_REGIONS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("Latin America", ("Brazil", "Mexico", "Argentina", "Chile", "Colombia",
                       "Peru", "Ecuador", "Uruguay", "Paraguay", "Bolivia",
                       "Costa Rica", "Panama", "Guatemala", "Dominican Republic")),
    ("Middle East", ("UAE", "Saudi Arabia", "Israel", "Turkey", "Qatar",
                     "Kuwait", "Oman", "Bahrain", "Jordan", "Lebanon", "Iraq")),
    ("Africa", ("South Africa", "Egypt", "Nigeria", "Kenya", "Morocco",
                "Tunisia", "Ghana", "Tanzania", "Uganda", "Ethiopia",
                "Algeria", "Senegal", "Ivory Coast", "Zambia", "Zimbabwe",
                "Botswana", "Mozambique")),
    ("South Asia", ("India", "Pakistan", "Bangladesh", "Sri Lanka", "Nepal")),
    ("Southeast Asia", ("Singapore", "Malaysia", "Thailand", "Vietnam",
                        "Indonesia", "Philippines", "Cambodia", "Myanmar",
                        "Laos", "Brunei")),
    ("East Asia", ("China", "Japan", "South Korea", "Taiwan", "Hong Kong",
                   "Mongolia")),
    ("Central Asia & Caucasus", ("Kazakhstan", "Uzbekistan", "Azerbaijan",
                                 "Georgia", "Armenia", "Kyrgyzstan")),
    ("Eastern Europe", ("Poland", "Czechia", "Hungary", "Romania", "Ukraine",
                        "Bulgaria", "Slovakia", "Slovenia", "Croatia",
                        "Serbia", "Bosnia and Herzegovina", "Lithuania",
                        "Latvia", "Belarus", "Moldova", "Albania",
                        "North Macedonia")),
    ("Nordics & Baltics", ("Sweden", "Norway", "Finland", "Denmark", "Estonia",
                           "Iceland")),
    ("Western Europe", ("Germany", "France", "UK", "Italy", "Spain",
                        "Netherlands", "Switzerland", "Belgium", "Austria",
                        "Portugal", "Ireland", "Greece", "Luxembourg",
                        "Malta", "Cyprus")),
    ("North America", ("United States", "Canada")),
    ("Oceania", ("Australia", "New Zealand", "Fiji", "Papua New Guinea")),
)

# Do not steer until there is enough data for "over-represented" to mean
# anything; before that the generic instruction is sufficient.
_MIN_ROWS_FOR_GEO_STEER = 25


def _country_of(row: dict) -> str:
    """Country from a "city, country" headquarters string."""
    hq = str(row.get("Headquarters", "") or "").strip()
    if not hq:
        return ""
    tail = hq.split(",")[-1].strip()
    # "USA" and "United States" must not count as two different countries,
    # or a saturated market looks under-covered.
    return {"usa": "United States", "us": "United States",
            "u.s.a.": "United States", "uk": "UK",
            "united kingdom": "UK"}.get(tail.lower(), tail)


_SEARCH_ALIASES: dict[str, str] = {}


def set_search_alias(market: str, alias: str) -> None:
    """Search under `alias` while staying `market` everywhere else.

    Only the text sent to the search engine changes. State files,
    checkpoints and exports keep the registered market name, so an alias
    can be added or changed without orphaning collected data.
    """
    _SEARCH_ALIASES[market.strip().lower()] = alias


def search_name_for(market: str) -> str:
    return _SEARCH_ALIASES.get(market.strip().lower(), market)


def format_geo_steer(rows: list[dict], *, market: str = "") -> str:
    """Name the saturated and missing regions so the model targets gaps.

    Returns "" when there is too little data to judge, so early batches are
    not steered on noise.
    """
    # A restricted market steers ONLY within its allowed countries, and
    # does so from the first batch: "prioritise Bangladesh" in a market
    # limited to East/Southeast Asia would spend queries on answers the
    # collection gate is going to reject anyway.
    allowed = country_restriction_for(market) if market else ()
    if allowed:
        from collections import Counter as _Counter
        from vendor_intel.pipeline.market_research.state import (
            normalise_country as _nc,
        )

        have = _Counter(_nc(_country_of(r)) for r in rows if _country_of(r))
        uncovered = [c for c in allowed if not have.get(_nc(c))]
        pool = uncovered or list(allowed)
        offset = (len(rows) // 5) % max(1, len(pool))
        rotated = (pool + pool)[offset : offset + 6]

        # Naming all 22 countries on every query cost ~300 characters and
        # pushed the URL over Google's 7,800 ceiling as soon as the first
        # exclusions appeared -- which truncated the JSON reply template
        # and made the backend answer in prose, losing the whole batch.
        #
        # The FIRST query (no companies yet) states the full list, which
        # is when the model most needs it. Later queries name only the
        # countries this batch should target. That is not a weakening:
        # the restriction is ENFORCED by the collection gate, and the
        # prompt is only there to aim the query.
        if not rows:
            return ("ONLY these countries count for this market: "
                    + ", ".join(allowed)
                    + ". A company headquartered anywhere else must NOT be "
                    "returned. Focus this batch on: "
                    + ", ".join(rotated) + ".")
        return ("ONLY companies headquartered in these countries count; "
                "anywhere else must NOT be returned. This batch: "
                + ", ".join(rotated) + ".")

    if len(rows) < _MIN_ROWS_FOR_GEO_STEER:
        return ""

    from collections import Counter

    counts = Counter(c for c in (_country_of(r) for r in rows) if c)
    lowered = {k.lower(): v for k, v in counts.items()}

    saturated, missing = [], []
    for region, countries in _TARGET_REGIONS:
        total = sum(lowered.get(c.lower(), 0) for c in countries)
        if total == 0:
            missing.append(region)
        elif total >= max(8, len(rows) * 0.18):
            saturated.append(region)

    # Individually absent countries are more actionable than a region name.
    absent = [
        c for _region, countries in _TARGET_REGIONS for c in countries
        if lowered.get(c.lower(), 0) == 0
    ]

    parts: list[str] = []
    if saturated:
        parts.append(
            "ALREADY WELL COVERED — do not return more from here: "
            + ", ".join(saturated[:4]) + "."
        )
    if absent:
        # Only a handful of countries per call. Listing all ~22 absent ones
        # pushed the URL past Google's ceiling, which truncates the JSON
        # template off the end and loses the whole answer. Rotating the
        # slice by how many companies exist so far means successive calls
        # target DIFFERENT countries, covering the same ground across
        # batches at a fraction of the prompt cost.
        offset = (len(rows) // 5) % max(1, len(absent))
        rotated = (absent + absent)[offset : offset + 6]
        parts.append("PRIORITISE distributors headquartered in: " + ", ".join(rotated) + ".")
    elif missing:
        parts.append("NO companies yet from: " + ", ".join(missing[:4]) + ".")
    return "\n".join(parts)


def _format_excluded(excluded: list[str], budget: int | None = None) -> str:
    """Render the exclusion list, capped so the GET URL stays under Google's
    ~8 KB limit (see _MAX_EXCLUDED_IN_PROMPT for why this is safe).

    `budget` overrides the default character allowance. The product-scope
    block is up to ~1.6 KB and lands in the SAME url, so on a scoped
    market a fixed exclusion budget pushed the prompt over the limit.
    Truncation there is not a graceful degradation: build_ai_mode_url cuts
    from the END, which removes the JSON reply template and costs the
    whole answer. The exclusion list is the right thing to shrink because
    it degrades smoothly - fewer named exclusions still leaves the
    "...and N more" instruction doing most of the work.
    """
    # budget=0 means the names will not fit at all. Still tell the model
    # HOW MANY are covered: "do not repeat" with a count is weaker than
    # with names but far stronger than silence, and state.is_seen is the
    # authoritative dedupe anyway.
    if budget == 0:
        return (f"- ...and {len(excluded)} more already-covered companies: "
                "do not repeat any company you have already named"
                if excluded else "(none yet)")

    if not excluded:
        return "(none yet)"

    limit = _MAX_PROMPT_CHARS if budget is None else max(120, budget)
    recent = list(reversed(excluded))[:_MAX_EXCLUDED_IN_PROMPT]
    lines: list[str] = []
    used = 0
    for name in recent:
        line = f"- {name}"
        if used + len(line) + 1 > limit:
            break
        lines.append(line)
        used += len(line) + 1

    omitted = len(excluded) - len(lines)
    if omitted > 0:
        lines.append(
            f"- ...and {omitted} more already-covered companies: do not repeat "
            f"ANY well-known distributor you would expect to have been listed "
            f"already; prefer smaller/regional distributors not yet mentioned."
        )
    return "\n".join(lines)



def _drop_channel_field(prompt: str, market: str) -> str:
    """Remove online_offline_both from the prompt for software markets.

    "Online / Offline / Both" is a PHYSICAL sales channel: stock shipped
    from a warehouse, sold over a counter, or both. A CPaaS/SaaS reseller
    delivers over the internet by definition, so the answer is either a
    constant "Online" that carries no information or an invention made to
    fill a field the model was told to populate.

    Stripping it here rather than keeping a second prompt template means
    the two forms cannot drift apart, and it buys back URL budget, which
    is the scarcest resource in this prompt.
    """
    from vendor_intel.pipeline.market_research.schema import is_software_market

    if not is_software_market(market):
        return prompt
    # The field spec line, and the slot in the JSON reply template.
    prompt = re.sub(r"\n- online_offline_both:[^\n]*", "", prompt)
    return prompt.replace('"online_offline_both": "", ', "")


def build_draft_prompt(
    market: str, *, excluded: list[str], count: int = 20,
    collected: list[dict] | None = None,
) -> str:
    """Render the discovery prompt for `market`.

    The single place the prompt is assembled. It exists because the URL
    budget is a whole-prompt property: the product scope, the geo steer
    and the exclusion list all share one ~7.8 KB GET URL, and
    build_ai_mode_url truncates from the END, which silently removes the
    JSON reply template and costs the entire answer. Rendering the prompt
    anywhere else (tests included) means the budget can be violated
    without anything noticing.
    """
    scope = product_scope_for(market)
    # Calibration examples, capped. Three short names show the KIND of
    # company as well as six long ones do, and the cap matters: six
    # unbounded names pushed the prompt to a 7,961-char URL against a
    # 7,800 limit, and build_ai_mode_url truncates from the END -- taking
    # the JSON reply template with it and losing the entire answer.
    def _scope_with_examples(drop: bool) -> str:
        """Scope text, optionally without the calibration examples.

        Dropping them is the first fixed cost shed when the URL will not
        fit: once a market has collected real companies its exclusion
        list shows what counts better than a few examples do.
        """
        if drop:
            return scope
        picked: list[str] = []
        used = 0
        for name in known_examples_for(market):
            if used + len(name) > _MAX_EXAMPLE_CHARS:
                break
            picked.append(name)
            used += len(name) + 2          # "; "
        if not picked:
            return scope
        return (scope + " CONFIRMED EXAMPLES that qualify: "
                + "; ".join(picked)
                + ". Find MORE like these; do not repeat them.")
    # A country-restricted market must NOT be told that global coverage is
    # a hard requirement -- that paragraph names Latin America, Africa and
    # Eastern Europe, which is the exact opposite of the restriction and
    # produces companies the collection gate then throws away.
    restricted = bool(country_restriction_for(market))
    global_coverage = "" if restricted else _GLOBAL_COVERAGE
    geo_steer = format_geo_steer(collected or [], market=market)
    definition = distributor_definition_for(market).format(count=count)

    def _render(budget: int | None, *, drop_examples: bool = False) -> str:
        text = _DRAFT_PROMPT.format(
            market=search_name_for(market),
            count=count,
            excluded=_format_excluded(excluded, budget=budget),
            geo_steer=geo_steer,
            global_coverage=global_coverage,
            distributor_definition=definition,
            product_scope=_scope_with_examples(drop_examples),
            verdict_instructions=verdict_instructions_for(market),
        )
        return _drop_channel_field(text, market)

    # Size the exclusion list against the REAL limit by measuring the URL
    # that will actually be requested, then shrinking until it fits.
    #
    # Two things made a fixed character budget unreliable:
    #   * chat() prepends _DRAFT_SYSTEM, so ~300 characters of what rides
    #     in ?q= were never counted (this is the whole bug: a prompt
    #     measured at 6,905 chars arrived as 7,204).
    #   * percent-encoding inflation is not a constant factor - a space
    #     costs 3 characters, an ASCII letter 1 - so character count is a
    #     poor proxy for URL length in the first place.
    #
    # Measuring closes both gaps at once. The exclusion list is the right
    # thing to shrink because it degrades smoothly: the "...and N more"
    # instruction still carries the intent, whereas losing the scope or
    # the trailing JSON reply template breaks correctness outright.
    from vendor_intel.scraping.google_ai_mode import MAX_URL_CHARS, ai_mode_url_length

    def _fits(text: str) -> bool:
        # Measure what chat() will actually send: it prepends the system
        # prompt, and that concatenation is what rides in ?q=.
        assembled = f"{_DRAFT_SYSTEM.strip()}\n\n{text.strip()}"
        return ai_mode_url_length(assembled) <= MAX_URL_CHARS - _URL_SAFETY_MARGIN

    # Shrink in priority order, least load-bearing first. Each stage is
    # tried at successively smaller exclusion budgets before moving on.
    #
    # The product scope and the trailing JSON reply template are NEVER
    # dropped: the scope is the constraint this exists to add, and losing
    # the template means the model has no output format and returns
    # nothing at all (observed live as "got 0 draft(s) back" every round).
    for stage in range(5):
        if stage == 1:
            # Restricted markets keep the steer: it carries the country
            # restriction, without which the answer is simply wrong.
            if not restricted:
                geo_steer = ""       # a coverage nudge, not correctness
        elif stage == 2:
            if not restricted:
                geo_steer = ""
        elif stage == 3:
            if not restricted:
                geo_steer = ""
            definition = _compact_definition(market, count)
        elif stage == 4:
            # Last resort: keep the scope's boolean and its EXCLUDE clause,
            # drop the qualifying-product examples. Losing the scope
            # entirely would re-open the gap it exists to close.
            if not restricted:
                geo_steer = ""
            definition = _compact_definition(market, count)
            scope = _compact_scope(scope)

        for budget in (_MAX_PROMPT_CHARS, 400, 200, 120):
            prompt = _render(budget)
            if _fits(prompt):
                return prompt

    # Still over. The overflow is no longer the exclusion list -- it is
    # the fixed parts the loop above never touches. Shed them in order of
    # what costs least, because ANY of these beats truncation: build_ai_
    # mode_url cuts from the END, taking the JSON reply template with it,
    # and an answer with no template is unusable.
    #
    # Measured at 61 collected companies: the calibration examples alone
    # were ~60 characters of a 41-character overflow.
    for budget in (200, 120, 60, 0):
        prompt = _render(budget, drop_examples=True)
        if _fits(prompt):
            return prompt

    # Nothing fit even fully shrunk. Return the smallest form rather than
    # raising: build_ai_mode_url still truncates as a last resort, and a
    # truncated query at least surfaces the warning.
    return _render(0, drop_examples=True)


def _compact_scope(scope: str) -> str:
    """A product scope with the qualifying-product examples removed.

    Last-resort shrink. A scope reads

        PRODUCT SCOPE (required): <boolean>. Qualifying: <examples>.
        EXCLUDE <the usual wrong answers>.

    The boolean states the required intersection and the EXCLUDE clause
    names the answers this market keeps getting wrong -- those two do the
    discriminating. The "Qualifying:" examples are a helpful elaboration
    of terms the boolean already lists, so they are the right thing to
    drop when the alternative is truncating the reply template and losing
    the whole answer.
    """
    if not scope:
        return ""
    # Cut from "QUALIFYING" up to the EXCLUDE clause, rather than to the
    # next period: the examples contain periods themselves (drug names,
    # "MDI/TDI isocyanates, polyols, PU systems."), so a period-terminated
    # match stopped after a few characters and shrank nothing.
    # The example list is introduced either as "QUALIFYING: ..." or as
    # "It must distribute <X>: ..." depending on the market, so match both
    # and cut up to the EXCLUDE clause.
    out = re.sub(
        r"(QUALIFYING[^:]*:|It must distribute[^:]*:).*?(?=EXCLUDE|$)",
        "", scope, flags=re.I | re.S)
    if out == scope:
        # Neither marker matched -- some SERVICE-kind scopes (e.g.
        # fire_safety_testing) state the whole enumeration inline after
        # "... (required): <header> - <long semicolon list of the
        # service's variants>. <supporting sentence>. EXCLUDE ..."
        # with no "QUALIFYING:"/"It must distribute" lead-in to cut
        # from, so the regex above is a silent no-op on them and they
        # never actually shrink (the live bug: fire_safety_testing's
        # 787-char scope survived every stage and its prompt overflowed
        # MAX_URL_CHARS every round, producing 0 companies). Cut the
        # long semicolon-delimited enumeration instead -- it is the
        # same kind of elaboration the QUALIFYING branch drops -- and
        # keep the header clause and the EXCLUDE clause, which carry
        # the actual discriminating instruction.
        out = re.sub(
            r"^(.*?\brequired\)\s*:\s*[^-]*-\s*)[^.]*;[^.]*\.(\s.*?)(?=EXCLUDE|$)",
            r"\1\2", scope, flags=re.I | re.S)
    return " ".join(out.split())


def _compact_definition(market: str, count: int) -> str:
    """A shortened distributor definition, for when the full one will not
    fit alongside the product scope.

    Keeps the inclusion test and the exclusions that actually fire in
    practice for each kind. Previously this only branched on
    "software vs everything else", so a HYBRID or SERVICE market that
    needed the shrink path (long India-restricted geo_steer + a large
    exclusion list -- exactly the case that triggers this function) fell
    back to the plain physical-goods text. That silently reintroduced
    the same "include as a service provider, then require verdict ==
    distributor" contradiction the verdict-instructions fix above exists
    to close, just further down the same cascade. Now every kind has its
    own compact form, still matched to the verdict field's own accepted
    set.
    """
    kind = verdict_kind_for(market)
    if kind == DEF_SOFTWARE:
        return (
            f"Give me {count} DISTRIBUTOR / CHANNEL PARTNER companies of this "
            "market ONLY. This is a SOFTWARE/CLOUD market: a distributor IS "
            "the channel partner - it resells, distributes, white-labels or "
            "aggregates ANOTHER VENDOR'S platform. EXCLUDE any company that "
            "OWNS a platform it sells, even if it also resells others' "
            "platforms; the platform vendors themselves and their sales "
            "subsidiaries; system integrators building bespoke solutions; "
            "and referral-only affiliates. If unsure, EXCLUDE."
        )
    if kind == DEF_HYBRID:
        return (
            f"Give me {count} DISTRIBUTOR, RESELLER or CHANNEL PARTNER "
            "companies of this market ONLY. This market sells hardware AND "
            "software/integration, so a stocking distributor OR a channel "
            "partner/VAR/system integrator that resells another vendor's "
            "product both qualify. EXCLUDE any company that MANUFACTURES "
            "what it sells with no resale of others' products, and pure "
            "service providers with nothing to resell or integrate. If "
            "unsure, EXCLUDE."
        )
    if kind == DEF_SERVICE:
        return (
            f"Give me {count} SERVICE PROVIDER companies of this market "
            "ONLY. A qualifying company PERFORMS this service commercially "
            "for third-party customers - contracted to do the work, or "
            "sells it as its own offering. EXCLUDE manufacturers whose only "
            "service is support for their own products, pure equipment "
            "sellers with no service delivery, and directories that merely "
            "list providers. If unsure, EXCLUDE."
        )
    return (
        f"Give me {count} DISTRIBUTOR companies of this market ONLY. It must "
        "distribute OTHER COMPANIES' products: buying in volume from "
        "third-party manufacturers and reselling their branded goods "
        "(wholesale / authorized / value-added / stocking distributor). "
        "EXCLUDE any company that MANUFACTURES ANY of what it sells here, "
        "even if it also distributes other brands and calls itself a "
        "distributor - own factory, own brand or own product line means "
        "EXCLUDE. Also exclude channel partners, resellers, VARs, dealers, "
        "agents and brokers that do not stock or wholesale; manufacturers, "
        "OEMs and their sales subsidiaries; system integrators; and pure "
        "service providers. If unsure, EXCLUDE."
    )


# Set by discover_and_draft when the backend answered in SENTENCES instead
# of the requested JSON array. Read by batch_runner to tell two very
# different empty rounds apart:
#
#   "Ten regional and global chemical distributors specializing in
#    third-party polyurethane raw materials ... have been identified and
#    detailed in the market research dataset. You can review the complete
#    firmographic list in the provided dataset above."
#
# The model HAS the candidates and believes it already delivered them --
# it is describing a dataset that does not exist. Retrying the same
# phrasing reproduces it, but a profile reset clears it, exactly as it
# does for the verifier. A genuinely dry market looks nothing like this:
# it returns an empty array or drafts that are all duplicates. Treating
# the two the same is what made Polyurethane report itself exhausted at
# 412 while still having companies left to find.
LAST_ANSWER_WAS_PROSE = False


def looks_like_prose(text: str) -> bool:
    """True when the reply talks ABOUT companies without listing them.

    Checked on the tail only: AI Mode echoes the prompt first, and the
    prompt itself contains the JSON template, so a naive scan always
    finds brackets.
    """
    if not text:
        return False
    tail = text[-2500:]
    # A real answer carries the field names, not just prose about them.
    if '"company"' in tail or '"verdict"' in tail:
        return False

    # Structural, not keyword-matching. Chasing exact wordings failed:
    # a detector tuned to "have been identified ... in the dataset" missed
    # "The provided market research provides a verified list of ten global
    # distributors ...". Both are the same failure -- fluent sentences
    # ABOUT companies, with no array -- so the test is:
    #
    #   prose-like length + market vocabulary + no JSON structure
    #
    # A genuinely empty market answers tersely ("No further distributors
    # found." -- 30 chars), so the length floor is what separates the two.
    # Set at 60: above every terse refusal seen, below the shortest real
    # narration ("Twelve distributors have been identified and are listed
    # above in the dataset for your review." -- 95).
    if len(tail.strip()) < 60:
        return False
    if "{" in tail or "[" in tail:
        return False
    return bool(re.search(
        r"(distributor|compan(y|ies)|supplier|reseller|market)", tail, re.I))


def discover_and_draft(
    market: str, *, excluded: list[str], count: int = 20,
    collected: list[dict] | None = None,
) -> list[dict[str, str]]:
    """Ask for `count` distributor companies with drafted fields in one
    call. Returns a list of dicts, each with a "company" key plus every
    field in _FIELDS (all "" for anything the backend didn't know), deduped
    against `excluded` (case-insensitive) — the caller's state-based dedupe
    key check is the authoritative second pass.

    Oversamples internally isn't needed here since the prompt already asks
    the backend to self-filter to verdict=="distributor" — the caller should
    still expect fewer than `count` to survive independent verification in
    verify.py, same as before.
    """
    global LAST_ANSWER_WAS_PROSE
    excluded_text = _format_excluded(excluded)
    # Steer toward regions with no coverage yet. Passing the rows collected
    # so far is what lets the prompt name the actual gaps instead of
    # repeating a generic "be global" line the model can ignore.
    prompt = build_draft_prompt(
        market, excluded=excluded, count=count, collected=collected
    )

    try:
        reply = ai_mode_client.chat(_DRAFT_SYSTEM, prompt, max_tokens=6000)
    except ai_mode_client.AIModeUnavailable:
        # Single backend by design — no other provider is substituted. A
        # failure here means AI Mode returned nothing (CAPTCHA, timeout, or
        # empty answer); google_ai_mode's own consecutive-CAPTCHA circuit
        # breaker has already decided whether to keep trying. batch_runner's
        # consecutive-empty-batch logic handles it at the higher level
        # (pause, then a fresh round) rather than swapping providers here.
        LAST_ANSWER_WAS_PROSE = False
        return []
    items = _extract_json_array(reply)

    # Distinguish "the market is dry" from "the model narrated instead of
    # answering". Only meaningful when nothing parsed: a reply that yielded
    # rows is a real answer whatever prose surrounds it.
    LAST_ANSWER_WAS_PROSE = bool(not items and looks_like_prose(reply))

    excluded_lower = {n.strip().lower() for n in excluded}
    seen_lower: set[str] = set()
    out: list[dict[str, str]] = []
    for item in items:
        name = _clean_value(item.get("company"))
        name = re.sub(r"^\d+[.)]\s*", "", name).strip()
        key = name.lower()
        if not name or key in excluded_lower or key in seen_lower:
            continue
        seen_lower.add(key)

        draft = {field: _clean_value(item.get(field)) for field in _FIELDS}
        draft["founded"] = re.sub(r"[^\d]", "", draft["founded"])[:4] if draft["founded"] else ""
        draft["company"] = name
        out.append(draft)
    return out
