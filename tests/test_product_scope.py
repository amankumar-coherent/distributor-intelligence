"""A distributor of the wrong PRODUCT still passed every check.

Reported for the Dry Eyes Treatment market: the collected companies were
ophthalmic EQUIPMENT and DEVICE distributors -- diagnostic systems, slit
lamps, clinical lasers, microsurgical instruments, optical retail. Every
one is a genuine ophthalmic wholesale distributor, so the distributor
definition accepted them all; none distributes a dry-eye DRUG.

Measured on the live 175-company state file: 24 unambiguously drug, 98
device-only, 49 mentioning both. That is the whole market list being
mostly wrong while passing every rule the prompt stated.

The gap: distributor_definition_for() constrains what a company IS.
Nothing constrained what it SELLS. product_scope_for() adds that, phrased
as the boolean the user asked for --

    "dry eye" AND ("pharmaceuticals" OR "ophthalmic emulsion" OR
    "lubricant eye drops") AND ("distributor" OR "wholesaler")

-- because a boolean states the required intersection far more sharply
than prose does.

The second half of this file guards the URL budget. The scope block is up
to ~1.1 KB and rides in the same ~7.8 KB GET URL as everything else, and
build_ai_mode_url truncates from the END -- which removes the JSON reply
template and costs the entire answer. Adding the scope naively pushed
every scoped market over the limit.
"""
from __future__ import annotations

import pytest

from vendor_intel.pipeline.market_research import discovery
from vendor_intel.scraping.google_ai_mode import MAX_URL_CHARS, build_ai_mode_url

_MARKETS = (
    "Global Silicon Carbide Market - Distributors Only",
    "Global Wearable Glucometer Market - Distributors Only",
    "Global Dry Eyes Treatment Market - Distributors Only",
    "Global Solar Rooftop Market - Distributors Only",
    "Global Polyurethane Market - Distributors Only",
    "Global CPaaS (Communication Platform as a Service) Market - Distributors Only",
)


# --- the dry-eye scope, which is the reported failure --------------------


def test_dry_eye_scope_uses_the_requested_boolean():
    """The user supplied this form specifically; it pins the required
    intersection of condition, product class and channel role."""
    scope = discovery.product_scope_for("Global Dry Eyes Treatment Market")
    assert '"dry eye" AND (' in scope
    assert '"distributor" OR "wholesaler"' in scope
    for term in ('"pharmaceuticals"', '"ophthalmic emulsion"',
                 '"lubricant eye drops"'):
        assert term in scope, term


def test_dry_eye_scope_names_real_drugs():
    """Naming actual products is what makes the constraint checkable -- a
    model can satisfy "sells dry eye products" with an equipment
    catalogue, but not "distributes Restasis"."""
    scope = discovery.product_scope_for("Global Dry Eyes Treatment Market").lower()
    for drug in ("restasis", "xiidra", "miebo", "cyclosporine",
                 "lifitegrast", "artificial tears"):
        assert drug in scope, drug


@pytest.mark.parametrize(
    "wrong",
    [
        "slit lamp", "meibograph", "tear-film", "laser", "microsurgical",
        "phaco", "oct", "contact lens", "spectacle", "optical",
        "instrument", "equipment", "device",
    ],
)
def test_dry_eye_scope_excludes_the_wrong_answers(wrong):
    """Each of these appeared in the live collected list."""
    scope = discovery.product_scope_for("Global Dry Eyes Treatment Market").lower()
    assert wrong in scope, f"{wrong} must be named as excluded"


def test_dry_eye_scope_demands_a_named_drug():
    """A broad-line pharma wholesaler is legitimate ONLY if it actually
    carries dry-eye drugs, so the prompt requires naming them. Without
    this, "pharmaceutical distributor" alone would re-admit everything."""
    scope = discovery.product_scope_for("Global Dry Eyes Treatment Market")
    assert "key_products_technologies" in scope
    assert "EXCLUDE" in scope


# --- scope applies to the other markets too ------------------------------


@pytest.mark.parametrize(
    "market,required",
    [
        ("Global Silicon Carbide Market", ("silicon carbide", "mosfet")),
        ("Global Wearable Glucometer Market", ("glucometer", "cgm")),
        ("Global Solar Rooftop Market", ("solar", "inverter")),
        ("Global Polyurethane Market", ("polyurethane", "polyol")),
    ],
)
def test_each_physical_market_has_a_product_scope(market, required):
    scope = discovery.product_scope_for(market).lower()
    assert scope, f"no scope for {market}"
    for term in required:
        assert term in scope, term


def test_an_unknown_market_gets_no_scope():
    """Returning "" is correct: an invented scope would filter on guessed
    product terms, which is worse than not filtering at all."""
    assert discovery.product_scope_for("Global Widget Market") == ""
    assert discovery.product_scope_for("") == ""


def test_scope_is_independent_of_the_distributor_definition():
    """They answer different questions -- what a company SELLS versus what
    it IS -- and the two vary independently.

    CPaaS carries the SOFTWARE definition (channel partners ARE the
    distributors there) together with a product scope, while an unlisted
    market carries the physical-goods definition and no scope at all. Each
    layer is chosen on its own."""
    cpaas = "Global CPaaS (Communication Platform as a Service) Market"
    assert discovery.distributor_definition_for(cpaas) is discovery._DEF_SOFTWARE_SERVICES
    assert discovery.product_scope_for(cpaas), "CPaaS has a scope of its own"

    unlisted = "Global Widget Market"
    assert discovery.distributor_definition_for(unlisted) is discovery._DEF_PHYSICAL_GOODS
    assert discovery.product_scope_for(unlisted) == ""


def test_the_cpaas_scope_requires_a_communications_platform():
    """Without it, a generic cloud or Microsoft 365 reseller satisfies
    every stated rule -- the same gap that filled the dry-eye market with
    ophthalmic equipment distributors."""
    scope = discovery.product_scope_for(
        "Global CPaaS (Communication Platform as a Service) Market").lower()
    for term in ("sms api", "voice api", "a2p", "twilio"):
        assert term in scope, term
    # and the characteristic wrong answer is named
    assert "generic cloud" in scope or "microsoft 365" in scope


def test_solar_scope_excludes_epc_and_installers():
    """Rooftop solar's characteristic wrong answer is the installer, which
    is an integrator rather than a distributor."""
    scope = discovery.product_scope_for("Global Solar Rooftop Market").lower()
    assert "epc" in scope
    assert "installer" in scope


# --- the URL budget -------------------------------------------------------


@pytest.mark.parametrize("market", _MARKETS)
def test_prompt_fits_the_url_limit_with_a_full_exclusion_list(market):
    """The real failure mode of a too-long prompt is NOT a clean error:
    build_ai_mode_url truncates from the end, cutting off the JSON reply
    template, and the whole answer is lost. Measured live at 400
    exclusions, which a 400-company target reaches."""
    prompt = discovery.build_draft_prompt(
        market,
        excluded=[f"Some Distributor Company Number {i} Limited" for i in range(400)],
        count=20,
        collected=[{"Headquarters": f"City {i}, Country {i}"} for i in range(120)],
    )
    url = build_ai_mode_url(prompt)
    assert len(url) <= MAX_URL_CHARS, (
        f"{market}: {len(url)} chars (prompt {len(prompt)})"
    )


@pytest.mark.parametrize("market", _MARKETS)
def test_the_json_template_survives(market):
    """The concrete symptom of truncation: the reply template is last in
    the prompt, so it is the first thing lost, and the model then has no
    output format to follow."""
    prompt = discovery.build_draft_prompt(
        market,
        excluded=[f"Company {i} Ltd" for i in range(400)],
        count=20,
    )
    url = build_ai_mode_url(prompt)
    assert len(url) <= MAX_URL_CHARS
    assert '"office_email": ""}]' in prompt, "reply template must be intact"


def test_the_scope_is_paid_for_out_of_the_exclusion_budget():
    """The exclusion list is the right thing to shrink -- it degrades
    smoothly because the "...and N more" line still carries the intent,
    whereas dropping scope or the reply template breaks correctness."""
    excluded = [f"Company {i} Ltd" for i in range(400)]
    scoped = discovery.build_draft_prompt(
        "Global Dry Eyes Treatment Market - Distributors Only",
        excluded=excluded, count=20)
    unscoped = discovery.build_draft_prompt(
        "Global Widget Market - Distributors Only",
        excluded=excluded, count=20)

    assert "PRODUCT SCOPE" in scoped
    assert "PRODUCT SCOPE" not in unscoped
    # Both must fit, and the scoped one lists fewer exclusions by name.
    assert scoped.count("\n- Company ") < unscoped.count("\n- Company ")
    assert "and " in scoped and "more already-covered" in scoped


def test_exclusion_budget_never_collapses_to_nothing():
    """A very long scope must still leave some named exclusions, or the
    model repeats companies it was just given."""
    rendered = discovery._format_excluded(
        [f"Company {i} Ltd" for i in range(50)], budget=10)
    assert rendered.strip(), "must not be empty"
    assert "more already-covered" in rendered


def test_scope_reaches_the_rendered_prompt():
    prompt = discovery.build_draft_prompt(
        "Global Dry Eyes Treatment Market - Distributors Only",
        excluded=[], count=20)
    assert "PRODUCT SCOPE" in prompt
    assert '"dry eye" AND (' in prompt
    # and it must not displace the role definition
    assert "DISTRIBUTOR companies of this market ONLY" in prompt


def test_the_pipeline_uses_the_shared_builder():
    """Assembling the prompt anywhere else means the URL budget can be
    violated without any test noticing -- which is how this regression
    reached the live run."""
    import inspect

    src = inspect.getsource(discovery.discover_and_draft)
    assert "build_draft_prompt(" in src
    assert "_DRAFT_PROMPT.format(" not in src


# --- headroom, not just "fits" -------------------------------------------
#
# The first fix technically passed -- worst case 7,780 of 7,800 -- but 20
# characters is not a margin, it is luck. One longer-than-average company
# name in the exclusion list would push it over, and the failure is
# silent: the reply template is cut and the model returns nothing.


def test_the_url_budget_keeps_real_headroom():
    """Stress every market x collected-count x batch-size combination with
    deliberately long company names."""
    worst = 0
    for market in _MARKETS:
        for n_collected in (0, 57, 150, 300, 400, 500, 600):
            collected = [
                {
                    "Headquarters": f"Cityname {i}, Countryname {i}",
                    "Company": (
                        f"Some Very Long Distributor Company Name "
                        f"Number {i} Private Limited"
                    ),
                }
                for i in range(n_collected)
            ]
            for count in (10, 20):
                prompt = discovery.build_draft_prompt(
                    market,
                    excluded=[r["Company"] for r in collected],
                    count=count,
                    collected=collected,
                )
                worst = max(worst, len(build_ai_mode_url(prompt)))

    headroom = MAX_URL_CHARS - worst
    assert headroom >= 150, (
        f"only {headroom} chars of headroom (worst URL {worst}/"
        f"{MAX_URL_CHARS}); one long company name would truncate the "
        "reply template and lose the whole answer"
    )


# --- the budget must cover what is ACTUALLY sent -------------------------
#
# The bug that reached the live run twice. ai_mode_client.chat(system,
# user) concatenates the two and THAT is what rides in ?q=. The budget
# only measured the user prompt, under-counting by len(_DRAFT_SYSTEM) + 2.
# A prompt measured at 6,905 characters arrived as 7,204, produced an
# 8,333-char URL, was truncated, lost the reply template, and every round
# returned "got 0 draft(s) back".
#
# Measuring the user prompt alone is therefore not a valid check, and
# every test above that does so would have passed while the pipeline
# collected nothing.


def _assembled(prompt: str) -> str:
    """Exactly what ai_mode_client.chat() sends for a discovery query."""
    return f"{discovery._DRAFT_SYSTEM.strip()}\n\n{prompt.strip()}"


@pytest.mark.parametrize("n_excluded", [0, 57, 200, 400, 600, 1000])
@pytest.mark.parametrize("market", _MARKETS)
def test_the_assembled_request_fits(market, n_excluded):
    """Budget against system+prompt, not the prompt alone."""
    from vendor_intel.scraping.google_ai_mode import ai_mode_url_length

    collected = [
        {
            "Headquarters": f"Cityname {i}, Countryname {i}",
            "Company": f"Some Very Long Distributor Company Name Number {i} Private Limited",
        }
        for i in range(n_excluded)
    ]
    prompt = discovery.build_draft_prompt(
        market,
        excluded=[r["Company"] for r in collected],
        count=20,
        collected=collected,
    )
    length = ai_mode_url_length(_assembled(prompt))
    assert length <= MAX_URL_CHARS, (
        f"{market} with {n_excluded} exclusions: {length}/{MAX_URL_CHARS}"
    )


@pytest.mark.parametrize("market", _MARKETS)
def test_the_two_things_that_must_never_be_dropped(market):
    """The shrink loop may sacrifice named exclusions, the geo steer, even
    the long form of the distributor definition. It must never sacrifice
    the reply template (no output format => no answer at all) or the
    product scope (the constraint the whole change exists to add)."""
    collected = [
        {"Headquarters": f"City {i}, Country {i}",
         "Company": f"Company Number {i} With A Deliberately Long Name Limited"}
        for i in range(1000)
    ]
    prompt = discovery.build_draft_prompt(
        market, excluded=[r["Company"] for r in collected],
        count=20, collected=collected,
    )
    assert '"office_email": ""}]' in prompt, "reply template dropped"
    if discovery.product_scope_for(market):
        assert "PRODUCT SCOPE" in prompt, "product scope dropped"


def test_a_safety_margin_is_reserved():
    """Shrinking only at the hard limit left 18 characters of headroom.
    The margin costs a few named exclusions; exceeding the limit costs the
    whole answer."""
    assert discovery._URL_SAFETY_MARGIN >= 150


def test_the_measurement_helper_does_not_truncate():
    """build_ai_mode_url() truncates and logs a warning, so using it to
    size a candidate prompt spams the log for prompts never sent."""
    from vendor_intel.scraping.google_ai_mode import (
        ai_mode_url_length,
        build_ai_mode_url,
    )

    huge = "x " * 20_000
    # reports the true length...
    assert ai_mode_url_length(huge) > MAX_URL_CHARS
    # ...while the builder caps it
    assert len(build_ai_mode_url(huge)) <= MAX_URL_CHARS


def test_the_compact_definition_keeps_the_inclusion_test():
    """It is used when the full definition will not fit, so it must still
    state what qualifies and the exclusions that actually fire."""
    for market in ("Global Polyurethane Market - Distributors Only",
                   "Global CPaaS (Communication Platform as a Service) Market"):
        compact = discovery._compact_definition(market, 20)
        assert "20" in compact
        assert "EXCLUDE" in compact
        assert len(compact) < len(
            discovery.distributor_definition_for(market).format(count=20)
        ), "must actually be shorter"
