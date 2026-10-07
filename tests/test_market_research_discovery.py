"""discovery.py: combined discover+draft prompt parsing and exclusion."""
from __future__ import annotations

import json
import re
from unittest.mock import patch

from vendor_intel.pipeline.market_research import discovery


def _reply(companies: list[dict]) -> str:
    return json.dumps(companies)


def test_discover_and_draft_parses_json_array():
    payload = _reply(
        [
            {"company": "Acme Distribution", "verdict": "distributor", "website": "acme.com"},
            {"company": "Beta Channel Partners", "verdict": "distributor", "website": "beta.com"},
        ]
    )
    with patch.object(discovery.ai_mode_client, "chat", return_value=payload):
        results = discovery.discover_and_draft("Test Market", excluded=[], count=20)

    assert [r["company"] for r in results] == ["Acme Distribution", "Beta Channel Partners"]
    assert results[0]["verdict"] == "distributor"
    assert results[0]["website"] == "acme.com"
    # Every schema draft field is present even when the model omitted it.
    assert "linkedin" in results[0]


def test_discover_and_draft_excludes_case_insensitively():
    payload = _reply(
        [
            {"company": "Acme Distribution", "verdict": "distributor"},
            {"company": "New Co", "verdict": "distributor"},
        ]
    )
    with patch.object(discovery.ai_mode_client, "chat", return_value=payload):
        results = discovery.discover_and_draft(
            "Test Market", excluded=["acme distribution"], count=20
        )
    assert [r["company"] for r in results] == ["New Co"]


def test_discover_and_draft_dedupes_within_same_reply():
    payload = _reply(
        [
            {"company": "Acme Distribution", "verdict": "distributor"},
            {"company": "acme distribution", "verdict": "distributor"},
        ]
    )
    with patch.object(discovery.ai_mode_client, "chat", return_value=payload):
        results = discovery.discover_and_draft("Test Market", excluded=[], count=20)
    assert len(results) == 1


def test_discover_and_draft_handles_malformed_json():
    with patch.object(discovery.ai_mode_client, "chat", return_value="not json at all"):
        results = discovery.discover_and_draft("Test Market", excluded=[], count=20)
    assert results == []


def test_discover_and_draft_ignores_echoed_prompt_template():
    """Google AI Mode renders the whole conversation, so the answer text
    starts with the prompt echoed back — and this prompt ENDS with a
    literal empty JSON template. A naive find("[") locks onto that empty
    template and yields 0 companies; observed live, a good 3-company answer
    parsed as 0 drafts. The real payload after the echo must win."""
    echoed = (
        'AI Mode conversation: ...one object per company:\n'
        '[{"company": "", "verdict": "", "website": ""}]'
    )
    answer = (
        '\n[\n{"company": "Ingram Micro", "verdict": "distributor", "website": "ingrammicro.com"},\n'
        '{"company": "TD SYNNEX", "verdict": "distributor", "website": "tdsynnex.com"}\n]\n'
        'AI Mode response is ready'
    )
    with patch.object(discovery.ai_mode_client, "chat", return_value=echoed + answer):
        results = discovery.discover_and_draft("Test Market", excluded=[], count=20)

    assert [r["company"] for r in results] == ["Ingram Micro", "TD SYNNEX"]


def test_discover_and_draft_strips_google_code_footer():
    payload = (
        'json\n[{"company": "Arrow Electronics", "verdict": "distributor"}]\n'
        "\nUse code with caution.\n"
    )
    with patch.object(discovery.ai_mode_client, "chat", return_value=payload):
        results = discovery.discover_and_draft("Test Market", excluded=[], count=20)
    assert [r["company"] for r in results] == ["Arrow Electronics"]


def test_discover_and_draft_normalizes_founded_year():
    payload = _reply([{"company": "Acme", "verdict": "distributor", "founded": "Founded in 1998"}])
    with patch.object(discovery.ai_mode_client, "chat", return_value=payload):
        results = discovery.discover_and_draft("Test Market", excluded=[], count=20)
    assert results[0]["founded"] == "1998"


def test_discover_and_draft_embeds_exclusion_list_in_every_prompt():
    """discover_and_draft always builds the same prompt (with the full
    exclusion list embedded), so a retried call still receives the same
    "do not repeat these companies" instruction on every retry — no
    duplicate can slip through just because a call had to be retried.
    The exclusion list lives in the prompt text itself, not in any
    conversation/session state, which is what makes this hold for a
    stateless backend like AI Mode."""
    excluded = ["Acme Distribution", "Beta Channel Partners"]
    with patch.object(discovery.ai_mode_client, "chat", return_value="[]") as mock_chat:
        discovery.discover_and_draft("Test Market", excluded=excluded, count=20)

    prompt = mock_chat.call_args.args[1]
    assert "Acme Distribution" in prompt
    assert "Beta Channel Partners" in prompt


def test_discover_and_draft_returns_empty_list_when_backend_unavailable():
    """When ai_mode_client.chat() raises (CAPTCHA, timeout, or empty
    answer), discover_and_draft must return [] — letting batch_runner's own
    retry-the-batch logic handle it — rather than crashing the whole
    run_until_target loop."""
    with patch.object(
        discovery.ai_mode_client,
        "chat",
        side_effect=discovery.ai_mode_client.AIModeUnavailable("captcha"),
    ):
        results = discovery.discover_and_draft("Test Market", excluded=[], count=20)
    assert results == []


def test_prompt_url_stays_under_google_limit_as_exclusions_grow():
    """Google 400s a search URL past ~8 KB ("malformed or illegal request").
    The whole prompt rides in ?q=, and the exclusion list grows by one line
    per company found — at 148 companies this produced a 10,701-char URL and
    every request failed mid-run. _format_excluded must cap it so the URL
    stays flat regardless of how many companies have been collected."""
    from vendor_intel.scraping.google_ai_mode import MAX_URL_CHARS, build_ai_mode_url

    for n in (0, 50, 148, 300, 500, 1000):
        excluded = [f"Some Distributor Company Number {i} Ltd" for i in range(n)]
        prompt = discovery._DRAFT_SYSTEM + "\n\n" + discovery.build_draft_prompt(
            "Global Data Center System Market - Distributors Only",
            excluded=excluded,
            count=10,
            collected=[{"Headquarters": f"City {i}, Country {i}"} for i in range(n)],
        )
        url = build_ai_mode_url(prompt)
        assert len(url) <= MAX_URL_CHARS, f"{n} exclusions -> {len(url)}-char URL"


def test_format_excluded_keeps_newest_and_reports_omitted_count():
    """Trimming is safe (state.is_seen is the real dedupe guard) but the
    model should still be told how many were omitted, and should get the
    MOST RECENT names since a repeat is likeliest among those."""
    excluded = [f"Company {i}" for i in range(200)]
    out = discovery._format_excluded(excluded)

    assert "Company 199" in out  # newest kept
    assert "more already-covered companies" in out  # omission disclosed
    assert len(out) <= discovery._MAX_PROMPT_CHARS + 300


def test_format_excluded_empty_list():
    assert discovery._format_excluded([]) == "(none yet)"


# --- geographic steering -------------------------------------------------
#
# Added after measuring real output: at 172 companies the list was 61%
# US/Germany/Japan/Taiwan with ZERO from Latin America, the Middle East,
# Southeast Asia or the Nordics. A single "aim for broad global coverage"
# line was too weak to change that, so the prompt now names actual gaps.


def test_geo_steer_is_silent_on_early_batches():
    """Before there is enough data, "over-represented" is noise."""
    rows = [{"Headquarters": "Austin, United States"}] * 5
    assert discovery.format_geo_steer(rows) == ""
    assert discovery.format_geo_steer([]) == ""


def test_geo_steer_flags_saturated_regions():
    rows = [{"Headquarters": "Austin, United States"}] * 40
    out = discovery.format_geo_steer(rows)
    assert "ALREADY WELL COVERED" in out
    assert "North America" in out


def test_geo_steer_targets_countries_with_no_coverage():
    rows = [{"Headquarters": "Austin, United States"}] * 40
    out = discovery.format_geo_steer(rows)
    assert "PRIORITISE" in out
    # Somewhere genuinely absent from the sample. Asserted as a property
    # rather than a fixed list of five names: the steer rotates through
    # the target countries so successive batches cover different ground,
    # and which ones surface depends on the list's length (113 countries)
    # -- naming specific expected countries made this test depend on that
    # rotation offset instead of on the behaviour it is checking.
    priority = [l for l in out.splitlines() if "PRIORITISE" in l][0]
    named = [c.strip(" .") for c in priority.split(":", 1)[1].split(",")]
    assert named, "the steer must name at least one country"
    assert "United States" not in named, "must not steer at a covered country"
    known = {c for _region, countries in discovery._TARGET_REGIONS
             for c in countries}
    assert set(named) <= known, f"unexpected country in steer: {named}"


def test_geo_steer_does_not_target_a_country_already_covered():
    """Steering toward a country that already has companies wastes the
    instruction and re-surfaces duplicates."""
    rows = [{"Headquarters": "Sao Paulo, Brazil"}] * 40
    out = discovery.format_geo_steer(rows)
    priority = [l for l in out.split("\n") if "PRIORITISE" in l]
    assert priority and "Brazil" not in priority[0]


def test_geo_steer_treats_usa_and_united_states_as_one_country():
    """"USA" and "United States" must not read as two countries, or a
    saturated market looks under-covered."""
    rows = ([{"Headquarters": "Austin, USA"}] * 20 +
            [{"Headquarters": "Boston, United States"}] * 20)
    out = discovery.format_geo_steer(rows)
    assert "North America" in out, "US coverage must be recognised as saturated"


def test_geo_steer_rotates_targets_between_calls():
    """A fixed slice would ask for the same 6 countries every batch. The
    rotation is what covers the full gap list across many batches while
    keeping each prompt short enough for Google's URL limit."""
    base = [{"Headquarters": "Austin, United States"}] * 40
    first = discovery.format_geo_steer(base)
    later = discovery.format_geo_steer(base + [{"Headquarters": "Berlin, Germany"}] * 25)
    assert first != later, "successive batches must target different countries"


def test_geo_steer_keeps_the_url_under_googles_limit():
    """The steering block costs prompt budget, and an over-long URL is
    TRUNCATED — which cuts the JSON output template off the end and loses
    the entire answer, not just the trimmed part."""
    from vendor_intel.scraping.google_ai_mode import MAX_URL_CHARS, build_ai_mode_url

    for n in (30, 100, 200, 400, 600):
        rows = [{"Headquarters": f"City, Country{i % 7}"} for i in range(n)]
        prompt = discovery._DRAFT_SYSTEM + "\n\n" + discovery.build_draft_prompt(
            "Global Semiconductor Market - Distributors Only",
            excluded=[f"Company {i} Ltd" for i in range(n)],
            count=10,
            collected=rows,
        )
        url = build_ai_mode_url(prompt)
        assert len(url) <= MAX_URL_CHARS, f"{n} rows -> {len(url)}-char URL"


def test_prompt_demands_global_coverage_as_a_requirement():
    """The old wording ("Aim for broad global coverage") was advisory and
    measurably ignored.

    Asserted on the RENDERED prompt, not the template: the paragraph is
    injected at build time so a country-restricted market can be given a
    fixed country list instead. Rendering also proves the placeholder is
    actually filled, which a template-only check cannot.
    """
    prompt = discovery.build_draft_prompt(
        "Global Data Center System Market - Distributors Only",
        excluded=[], count=10, collected=[])
    assert "HARD REQUIREMENT" in prompt
    assert "INCOMPLETE" in prompt
    for region in ("Latin America", "Africa", "Middle East", "Southeast Asia"):
        assert region in prompt, region


def test_a_country_restricted_market_is_not_told_to_go_global():
    """The opposite case, which is why the paragraph became conditional.

    Telling a market limited to 16 Asian countries that "distributors
    exist in every region -- Latin America, Africa..." spends the answer
    on companies the collection gate rejects outright.
    """
    market = "Test Restricted Market - Distributors Only"
    discovery.set_country_restriction(market, ("Japan", "South Korea"))
    try:
        prompt = discovery.build_draft_prompt(
            market, excluded=[], count=10, collected=[])
        assert "HARD REQUIREMENT" not in prompt
        assert "Latin America" not in prompt
        assert "ONLY these countries count" in prompt
        assert "Japan" in prompt and "South Korea" in prompt
    finally:
        discovery._COUNTRY_RESTRICTIONS.pop(market.strip().lower(), None)


def test_country_of_parses_headquarters():
    assert discovery._country_of({"Headquarters": "Sao Paulo, Brazil"}) == "Brazil"
    assert discovery._country_of({"Headquarters": "Singapore"}) == "Singapore"
    assert discovery._country_of({"Headquarters": ""}) == ""


# --- market-specific distributor definition ------------------------------
#
# "Distributor" is not the same thing in every market. Per explicit user
# instruction for CPaaS: "channel partners are known as distributors over
# here". The default physical-goods rule EXCLUDES channel partners, so
# running CPaaS under it would reject essentially every valid company in
# the market - none of them stock inventory.


def test_physical_goods_markets_use_the_stocking_definition():
    for market in (
        "Global Silicon Carbide Market - Distributors Only",
        "Global Polyurethane Market - Distributors Only",
        "Global Solar Rooftop Market - Distributors Only",
        "Global Dry Eyes Treatment Market - Distributors Only",
        "Global Wearable Glucometer Market - Distributors Only",
    ):
        assert (
            discovery.distributor_definition_for(market)
            is discovery._DEF_PHYSICAL_GOODS
        ), market


def test_cpaas_uses_the_channel_partner_definition():
    for market in (
        "Global CPaaS (Communication Platform as a Service) Market - Distributors Only",
        "Global CPaaS Market",
        "Communication Platform as a Service Market",
    ):
        assert (
            discovery.distributor_definition_for(market)
            is discovery._DEF_SOFTWARE_SERVICES
        ), market


def test_physical_definition_excludes_channel_partners():
    assert "CHANNEL PARTNERS" in discovery._DEF_PHYSICAL_GOODS
    assert "STRICTLY EXCLUDE" in discovery._DEF_PHYSICAL_GOODS


def test_software_definition_includes_channel_partners():
    """The inverse rule: resellers/aggregators are what we WANT here."""
    d = discovery._DEF_SOFTWARE_SERVICES
    assert "CHANNEL PARTNER" in d
    assert "a distributor IS the channel partner" in d
    # ...but the platform vendors themselves are still excluded
    assert "PLATFORM VENDORS themselves" in d


def test_both_definitions_render_into_the_prompt():
    """A missing format key would raise at run time, mid-batch."""
    for market in (
        "Global Silicon Carbide Market - Distributors Only",
        "Global CPaaS Market - Distributors Only",
    ):
        prompt = discovery.build_draft_prompt(market, excluded=[], count=10)
        assert "Give me 10 DISTRIBUTOR" in prompt
        # No UNFILLED placeholders left. The JSON output template
        # legitimately contains braces, so look for {word} specifically.
        assert not re.search(r"\{[a-z_]+\}", prompt), prompt[:200]


def test_every_requested_market_prompt_fits_googles_url_limit():
    """The six markets the user asked for must not truncate - truncation
    cuts the JSON template off the end and loses the whole answer."""
    from vendor_intel.scraping.google_ai_mode import MAX_URL_CHARS, build_ai_mode_url

    markets = [
        "Global Silicon Carbide Market - Distributors Only",
        "Global Wearable Glucometer Market - Distributors Only",
        "Global Dry Eyes Treatment Market - Distributors Only",
        "Global Solar Rooftop Market - Distributors Only",
        "Global Polyurethane Market - Distributors Only",
        "Global CPaaS (Communication Platform as a Service) Market - Distributors Only",
    ]
    for market in markets:
        for n in (0, 200, 500):
            rows = [{"Headquarters": f"City, Country{i % 9}"} for i in range(n)]
            prompt = discovery._DRAFT_SYSTEM + "\n\n" + discovery.build_draft_prompt(
                market,
                excluded=[f"Company {i} Ltd" for i in range(n)],
                count=10,
                collected=rows,
            )
            url = build_ai_mode_url(prompt)
            assert len(url) <= MAX_URL_CHARS, f"{market} @ {n}: {len(url)} chars"
