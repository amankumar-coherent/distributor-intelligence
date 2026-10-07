"""Tests for company screening verdict parsing.

The Google AI path reads a free-text AI Overview and has to decide
distributor vs manufacturer from prose. Getting that wrong in the
distributor direction is the costly failure: it puts a manufacturer into
the export, which the whole screen exists to prevent.
"""
from __future__ import annotations

import pytest

from vendor_intel.discovery import chatgpt_company_screen as cs


class TestVerdictFromOverview:
    def test_oem_stated_as_not_a_reseller_is_a_manufacturer(self):
        """Regression: seen live for APROKOREA. The sentence contains the
        word "reseller", but negated — naive keyword matching read it as
        "distributor", the exact opposite of what it says."""
        verdict, _ = cs._verdict_from_overview(
            "APROKOREA is an original equipment manufacturer (OEM), not a third-party reseller."
        )
        assert verdict == "manufacturer"

    def test_plain_manufacturer(self):
        verdict, _ = cs._verdict_from_overview(
            "WellsCare Co., Ltd. is a manufacturer of medical devices. "
            "It designs and manufactures proprietary PBM devices."
        )
        assert verdict == "manufacturer"

    def test_manufactures_and_also_distributes_counts_as_manufacturer(self):
        """A maker that also distributes its own goods is still a maker —
        manufacturer evidence must win over distributor evidence."""
        verdict, _ = cs._verdict_from_overview(
            "M.I.Tech is primarily a manufacturer, though it also distributes "
            "the products it creates."
        )
        assert verdict == "manufacturer"

    def test_genuine_distributor(self):
        verdict, reason = cs._verdict_from_overview(
            "Acme Medical is a distributor of wearable monitoring devices, "
            "reselling brands such as Philips and Garmin."
        )
        assert verdict == "distributor"
        assert "distributor" in reason.lower()

    def test_authorised_distributor_phrasing(self):
        verdict, _ = cs._verdict_from_overview(
            "Beta Health is an authorised distributor for several wearable "
            "device brands across Europe."
        )
        assert verdict == "distributor"

    @pytest.mark.parametrize(
        "text",
        [
            "",
            "   ",
            "Your search did not match any documents.",
            "About 0 results (0.77s)",
            "Gamma Group is not a manufacturer of these products.",
        ],
    )
    def test_ambiguous_or_missing_is_unknown_never_a_guess(self, text):
        verdict, _ = cs._verdict_from_overview(text)
        assert verdict == "unknown"

    def test_reason_quotes_the_deciding_sentence(self):
        _, reason = cs._verdict_from_overview(
            "Some intro text. Delta Corp is a distributor of medical wearables. More text."
        )
        assert "Delta Corp is a distributor" in reason


class TestManufacturerPhrasingFromLiveAnswers:
    """Real AI Overview answers, previously misread as "unknown".

    They state the company IS a manufacturer while also saying it is NOT a
    distributor. The negation rule correctly blocked the distributor match,
    but no manufacturer phrase matched either, so clear evidence was thrown
    away."""

    @pytest.mark.parametrize(
        "label,text",
        [
            (
                "APROKOREA",
                "APROKOREA is an original manufacturer, not a distributor or "
                "reseller of other brands.",
            ),
            (
                "METAPHARM",
                "METAPHARM Co., Ltd. operates as both a manufacturer and a "
                "distributor within the medical device sector.",
            ),
            ("BODERM", "is a manufacturer and brand owner, not a mere distributor."),
            (
                "DEEPSONBIO",
                "the company designs and manufactures proprietary low-intensity "
                "ultrasound devices.",
            ),
        ],
    )
    def test_reads_as_manufacturer(self, label, text):
        verdict, _ = cs._verdict_from_overview(text)
        assert verdict == "manufacturer", label

    def test_a_real_distributor_is_still_detected(self):
        """Guard against over-correcting: the added manufacturer phrases must
        not swallow genuine distributors."""
        verdict, _ = cs._verdict_from_overview(
            "Acme is a distributor of wearable devices, reselling Philips and Garmin."
        )
        assert verdict == "distributor"


class TestNegationDetection:
    def test_negated_phrase_is_skipped_but_later_real_one_still_counts(self):
        verdict, _ = cs._verdict_from_overview(
            "Epsilon is not a reseller. Epsilon is a distributor of imported brands."
        )
        assert verdict == "distributor"

    @pytest.mark.parametrize(
        "negation", ["not a", "not an", "rather than a", "instead of a", "is not a"]
    )
    def test_each_negation_form_blocks_the_match(self, negation):
        verdict, _ = cs._verdict_from_overview(f"Zeta Ltd is {negation} distributor of devices.")
        assert verdict != "distributor"


class TestParseJsonArray:
    """Google AI Mode echoes the whole prompt above its answer, so the first
    "[...]" span in the reply is the TEMPLATE we asked it to follow, not the
    answer. Taking the first span discarded every batch (25/25 "unknown")
    even though a perfect array sat right below it."""

    _ECHOED = (
        "Reply with ONLY a JSON array, one object per company:\n"
        '[{"n": <number>, "verdict": "distributor|manufacturer|other|unknown", '
        '"reason": "<short reason>"}]\n'
        "[\n"
        '{"n": 1, "verdict": "unknown", "reason": "No evidence found."},\n'
        '{"n": 2, "verdict": "manufacturer", "reason": "Makes its own devices."},\n'
        '{"n": 3, "verdict": "distributor", "reason": "Resells third-party brands."}\n'
        "]\n"
        "AI Mode response is ready"
    )

    def test_ignores_the_echoed_prompt_template_and_finds_the_real_answer(self):
        items = cs._parse_json_array(self._ECHOED)
        assert len(items) == 3
        assert items[2]["verdict"] == "distributor"

    def test_tolerates_trailing_page_chrome_after_the_array(self):
        items = cs._parse_json_array(
            '[{"n": 1, "verdict": "other"}]\nAI Mode response is ready\nMore results'
        )
        assert len(items) == 1
        assert items[0]["verdict"] == "other"

    def test_plain_array(self):
        assert len(cs._parse_json_array('[{"n":1,"verdict":"distributor"}]')) == 1

    def test_markdown_fenced_array(self):
        assert len(cs._parse_json_array('```json\n[{"n":1,"verdict":"other"}]\n```')) == 1

    @pytest.mark.parametrize("junk", ["", "   ", "no json here", "[not, valid, json", "[]"])
    def test_unparseable_returns_empty_instead_of_raising(self, junk):
        assert cs._parse_json_array(junk) == []

    def test_later_array_wins_when_several_are_present(self):
        reply = '[{"n":1,"verdict":"unknown"}]\nrevised:\n[{"n":1,"verdict":"distributor"}]'
        items = cs._parse_json_array(reply)
        assert len(items) == 1
        assert items[0]["verdict"] == "distributor"


class TestCountryExtraction:
    """The geography columns must list NAMED COUNTRIES.

    Regression: the prompt asked for "countries or regions", so live answers
    came back as "Europe" / "Global" — useless in a spreadsheet cell. The
    question now demands named countries, and the prose extractor has to
    capture the resulting list (previously it parsed only website/founded/
    employees, so the countries were discarded).
    """

    GAPS = ["Continent / Geography", "Operational Presence"]

    def test_extracts_a_full_country_list(self):
        names = cs._country_names_in(
            "Italy, India, Colombia, Ecuador, Germany, Spain, South Korea, "
            "United States, Canada, Japan."
        )
        for country in ("Italy", "India", "Colombia", "South Korea", "Japan"):
            assert country in names
        assert len(names) >= 8, "a full multi-country list must survive intact"

    def test_normalises_common_aliases(self):
        names = cs._country_names_in(
            "Operates in Germany and the UK, with offices in the USA."
        )
        assert "United Kingdom" in names
        assert "United States" in names
        assert "Germany" in names

    def test_the_two_geo_columns_are_asked_separately(self):
        """Bundling both questions into one query made both columns receive
        the SAME answer, destroying the offices-vs-sales-territory
        distinction. They must be distinct questions."""
        geo = cs._GEO_QUESTIONS["Continent / Geography"]
        ops = cs._GEO_QUESTIONS["Operational Presence"]
        assert set(geo).isdisjoint(ops), "the two columns must ask different things"
        # Assert the INTENT, not exact wording — the phrasings get retuned
        # against measured results (the default for Operational Presence has
        # already changed once, from "distribution network" to "supply
        # customers in", which scored 24 vs 3 across three companies).
        assert any("office" in q or "physical location" in q for q in geo), (
            "Geography must ask about offices/physical presence"
        )
        assert any(
            w in q
            for q in ops
            for w in ("supply", "sell", "export", "distribut", "market")
        ), "Operational Presence must ask where it sells/distributes"
        assert set(cs._GEO_COLUMNS) == {"Continent / Geography", "Operational Presence"}

    def test_one_phrasing_per_column_by_default(self, monkeypatch):
        """Default is ONE question per column. Each extra phrasing is another
        AI Overview call per company — measured, using all of them took a
        9,460-company market from ~12h to ~23h."""
        monkeypatch.delenv("GOOGLE_AI_GEO_DEEP", raising=False)
        for column in cs._GEO_COLUMNS:
            assert len(cs._geo_questions_for(column)) == 1

    def test_deep_mode_adds_more_phrasings(self, monkeypatch):
        """Google AI Overview is non-deterministic — the same query returned
        0-7 countries across six runs, union 7. Unioning extra phrasings
        widens coverage when that matters more than turnaround."""
        monkeypatch.setenv("GOOGLE_AI_GEO_DEEP", "true")
        for column in cs._GEO_COLUMNS:
            assert len(cs._geo_questions_for(column)) >= 2

    def test_deep_mode_keeps_the_default_question_first(self, monkeypatch):
        monkeypatch.setenv("GOOGLE_AI_GEO_DEEP", "true")
        for column in cs._GEO_COLUMNS:
            assert cs._geo_questions_for(column)[0] == cs._GEO_QUESTIONS[column][0]

    def test_geo_questions_avoid_negative_instructions(self):
        """Measured: "do NOT answer 'Global'" phrasing returned 0 countries
        where a plain question returned 7. Negative instructions read as
        noise to a search engine."""
        for questions in list(cs._GEO_QUESTIONS.values()) + list(cs._GEO_QUESTIONS_DEEP.values()):
            for q in questions:
                assert "do NOT" not in q

    def test_longest_name_wins_over_substring(self):
        """"United States" must not also register a bare "US"."""
        names = cs._country_names_in("Sells across the United States.")
        assert names == ["United States"]

    def test_no_duplicates_and_first_seen_order(self):
        names = cs._country_names_in("Italy, then France, then Italy again.")
        assert names == ["Italy", "France"]

    def test_vague_answer_yields_only_real_countries(self):
        """"worldwide"/"Europe" are not countries and must not be recorded."""
        names = cs._country_names_in(
            "Distributes worldwide with a concentration in Italy and Europe."
        )
        assert names == ["Italy"]

    def test_nothing_named_extracts_nothing(self):
        out = cs._extract_fields_from_prose(
            "No specific countries are named in the available information.", self.GAPS
        )
        assert "Operational Presence" not in out

    def test_word_boundaries_prevent_false_matches(self):
        """"Chile" must not fire inside another word."""
        assert cs._country_names_in("The Chilean-style architecture") == []

    def test_company_name_is_not_mined_for_countries(self):
        """Regression: "JORDAN GAMMA MEDIZINTECHNIK GMBH" recorded the
        country Jordan, purely because the name was echoed in the answer."""
        text = "JORDAN GAMMA MEDIZINTECHNIK GMBH operates in Germany and Austria."
        names = cs._country_names_in(text, ignore="JORDAN GAMMA MEDIZINTECHNIK GMBH")
        assert "Jordan" not in names
        assert names == ["Germany", "Austria"]

    def test_ignore_is_case_insensitive(self):
        names = cs._country_names_in(
            "Jordan Gamma sells in Spain.", ignore="JORDAN GAMMA"
        )
        assert names == ["Spain"]

    def test_ignore_does_not_remove_a_genuine_country_elsewhere(self):
        """Stripping the name must not blind us to that country when it is
        legitimately mentioned as a market."""
        names = cs._country_names_in(
            "Jordan Gamma GmbH exports to Jordan and Egypt.", ignore="Jordan Gamma GmbH"
        )
        assert "Jordan" in names
        assert "Egypt" in names


class TestContactExtraction:
    """Contact Person and Role sat at exactly 0/405 in a live run, because
    the prose extractor had no logic for them at all — no answer could ever
    have filled them. Only sales/procurement leadership qualifies."""

    GAPS = ["Contact Person", "Role"]

    @pytest.mark.parametrize(
        "text,person,role",
        [
            (
                "The Head of Sales is Marco Rossi, who leads the team.",
                "Marco Rossi",
                "Head of Sales",
            ),
            ("Marco Rossi, Sales Director, oversees distribution.", "Marco Rossi", "Sales Director"),
            ("Sales Manager Anna Bianchi handles key accounts.", "Anna Bianchi", "Sales Manager"),
            (
                "Contact John Smith, Chief Procurement Officer, for orders.",
                "John Smith",
                "Chief Procurement Officer",
            ),
        ],
    )
    def test_extracts_sales_and_procurement_leaders(self, text, person, role):
        out = cs._extract_fields_from_prose(text, self.GAPS)
        assert out.get("Contact Person") == person
        assert out.get("Role") == role

    @pytest.mark.parametrize(
        "text",
        [
            "The CEO is Marco Rossi and the founder is Anna Bianchi.",
            "Marketing Manager Lisa Green runs campaigns.",
            "Chief Technology Officer Paolo Verdi leads engineering.",
            "General Manager Sofia Neri runs the site.",
        ],
    )
    def test_rejects_non_sales_roles(self, text):
        """Standing rule: only sales/procurement leadership is recorded, so
        a CEO/founder/marketing contact must be left empty, not substituted."""
        out = cs._extract_fields_from_prose(text, self.GAPS)
        assert "Contact Person" not in out
        assert "Role" not in out

    def test_lead_in_verbs_are_not_part_of_the_name(self):
        """"Contact John Smith" must yield "John Smith", not "Contact John"."""
        out = cs._extract_fields_from_prose(
            "Email Anna Bianchi, Sales Director, to order.", self.GAPS
        )
        assert out.get("Contact Person") == "Anna Bianchi"

    def test_lowercase_role_mid_sentence(self):
        """Regression from a live answer: "the overall head of sales is
        Corrado Damascati". A capitals-only role pattern missed it."""
        out = cs._extract_fields_from_prose(
            "At Eurospital S.p.A., the overall head of sales is Corrado Damascati.",
            self.GAPS,
        )
        assert out.get("Contact Person") == "Corrado Damascati"

    def test_real_live_answer_with_several_people(self):
        out = cs._extract_fields_from_prose(
            "Marco Poles is the Diagnostic Division Director at Eurospital S.p.A. "
            "Corrado Damascati is a Sales Manager (Responsabile vendite). "
            "Fernando Rebollo is the Sales Director Iberia and LATAM.",
            self.GAPS,
        )
        # Any of the sales-titled people is acceptable; a non-sales one is not.
        assert out.get("Contact Person") in {"Corrado Damascati", "Fernando Rebollo"}

    def test_lowercase_filler_is_not_mistaken_for_a_name(self):
        """Making roles case-insensitive briefly let "not currently disclosed"
        through as a person. Names must still require real capitals."""
        out = cs._extract_fields_from_prose(
            "the head of sales is not currently disclosed.", self.GAPS
        )
        assert "Contact Person" not in out

    def test_name_does_not_absorb_a_trailing_verb(self):
        out = cs._extract_fields_from_prose(
            "Sales Manager Anna Bianchi handles key accounts.", self.GAPS
        )
        assert out.get("Contact Person") == "Anna Bianchi"  # not "Anna Bianchi handles"

    def test_title_in_parentheses_after_local_language_title(self):
        """Regression from a live answer that yielded nothing across 431
        companies. Google gives the local-language title first and the
        English one in parentheses, so the qualifying role is NOT adjacent
        to the name."""
        out = cs._extract_fields_from_prose(
            "Giulia Fratta is the Responsabile Acquisti (Pharma Division "
            "Purchasing Manager) at Eurospital S.p.A.",
            self.GAPS,
        )
        assert out.get("Contact Person") == "Giulia Fratta"
        assert "Purchasing Manager" in out.get("Role", "")

    def test_distant_title_still_respects_the_role_gate(self):
        """The wider gap must not let a non-sales title through."""
        out = cs._extract_fields_from_prose(
            "Paolo Verdi is the Direttore Tecnico (Chief Technology Officer) at Acme.",
            self.GAPS,
        )
        assert "Contact Person" not in out

    def test_picks_the_sales_person_not_an_earlier_ceo(self):
        out = cs._extract_fields_from_prose(
            "Marco Rossi is the CEO. Anna Bianchi is the Sales Director.", self.GAPS
        )
        assert out.get("Contact Person") == "Anna Bianchi"

    def test_role_is_the_title_only_not_a_run_on_phrase(self):
        out = cs._extract_fields_from_prose(
            "The Head of Sales is Marco Rossi, who leads the team.", self.GAPS
        )
        assert out["Role"] == "Head of Sales"  # not "Head of Sales is"


class TestConfig:
    def test_google_ai_screening_off_by_default(self, monkeypatch):
        monkeypatch.delenv("KOMPASS_GOOGLE_AI_SCREEN", raising=False)
        assert cs.google_ai_screening_enabled() is False

    @pytest.mark.parametrize("value", ["true", "1", "yes", "on", "TRUE"])
    def test_google_ai_screening_enabled_values(self, monkeypatch, value):
        monkeypatch.setenv("KOMPASS_GOOGLE_AI_SCREEN", value)
        assert cs.google_ai_screening_enabled() is True

    def test_concurrency_default_and_clamping(self, monkeypatch):
        monkeypatch.delenv("GOOGLE_AI_SCREEN_CONCURRENCY", raising=False)
        assert cs.google_ai_concurrency() == 5

        monkeypatch.setenv("GOOGLE_AI_SCREEN_CONCURRENCY", "999")
        assert cs.google_ai_concurrency() == 10  # capped

        monkeypatch.setenv("GOOGLE_AI_SCREEN_CONCURRENCY", "0")
        assert cs.google_ai_concurrency() == 1  # floored

        monkeypatch.setenv("GOOGLE_AI_SCREEN_CONCURRENCY", "not-a-number")
        assert cs.google_ai_concurrency() == 5  # falls back
