"""A company must distribute THIS market, not merely be a distributor.

Fritz Minke GmbH & Co. KG was exported as a Wearable Medical Devices
distributor. Google's own answer to the screening question was:

    "No, Fritz Minke GmbH & Co. KG is not a distributor of wearable
     medical devices. They operate as a specialized B2B wholesaler and
     manufacturer for traditional orthopedic shoemaking..."

The rejection sits in sentence 1; the distributor vocabulary ("wholesaler")
sits in sentence 2 and describes a DIFFERENT industry. Because the negation
check is sentence-scoped it could not reach across, so the row was accepted.
"""
from __future__ import annotations

from vendor_intel.discovery.chatgpt_company_screen import _verdict_from_overview

MARKET = "Wearable Medical Devices"

MINKE_LIVE_ANSWER = (
    "No, Fritz Minke GmbH & Co. KG is not a distributor of wearable medical devices. "
    "They operate as a specialized B2B wholesaler and manufacturer for traditional "
    "orthopedic shoemaking and footwear technology."
)


def verdict(text: str) -> str:
    return _verdict_from_overview(text, market=MARKET)[0]


class TestRejectsWrongMarket:
    def test_the_real_answer_that_leaked_into_the_export(self):
        assert verdict(MINKE_LIVE_ANSWER) == "manufacturer"

    def test_reason_records_why_it_was_rejected(self):
        _, reason = _verdict_from_overview(MINKE_LIVE_ANSWER, market=MARKET)
        assert "not in" in reason.lower()

    def test_distributor_vocabulary_for_another_industry_does_not_rescue_it(self):
        assert verdict(
            "No, Acme is not a distributor of wearable medical devices. "
            "It is a wholesaler of shoes."
        ) == "manufacturer"

    def test_does_not_sell_phrasing(self):
        assert verdict(
            "Acme does not sell wearable medical devices; it is a wholesaler of textiles."
        ) == "manufacturer"


class TestStillAcceptsGenuineDistributors:
    """The gate must not become a blanket 'reject anything with No in it'."""

    def test_plain_yes(self):
        assert verdict("Yes, Acme is a distributor of wearable medical devices in Italy.") == "distributor"

    def test_reseller_wording(self):
        assert verdict("Acme is a reseller of wearable medical devices and monitors.") == "distributor"

    def test_answer_that_never_names_the_market(self):
        assert verdict("Acme is a distributor of medical equipment.") == "distributor"

    def test_negation_about_something_other_than_the_market(self):
        """'not publicly listed' is not a market rejection — the market
        words appear in the NEXT sentence and must not be borrowed."""
        assert verdict(
            "No, Acme is not publicly listed. Acme is a distributor of wearable medical devices."
        ) == "distributor"

    def test_not_a_manufacturer_is_a_positive_signal(self):
        assert verdict(
            "Acme is not a manufacturer. It is a distributor of wearable medical devices."
        ) == "distributor"


class TestNoRegression:
    def test_manufacturer_still_detected(self):
        assert verdict("Acme manufactures wearable medical devices.") == "manufacturer"

    def test_market_is_optional(self):
        """Callers that pass no market keep the old behaviour."""
        assert _verdict_from_overview("Acme is a distributor of devices.")[0] == "distributor"

    def test_empty_answer_is_unknown(self):
        assert _verdict_from_overview("", market=MARKET)[0] == "unknown"


class TestBotBlockPage:
    """Google serves a CAPTCHA interstitial instead of results when it
    rate-limits us. That page echoes the query string back — which contains
    the words "distributor" and "manufacturer" — so parsing it as prose
    produced a confident verdict for every company screened during a block.
    Observed live: three companies in a row scored "distributor" purely
    from the echoed URL."""

    BLOCK_PAGE = (
        "About this page Our systems have detected unusual traffic from your computer "
        "network. This page checks to see if it's really you sending the requests, and "
        "not a robot. Why did this happen? ... com/search?q=Is%20JIN%20SRL%20a%20"
        "distributor%20(reseller%20of%20other%20brands)%20or%20a%20manufacturer"
    )

    def test_block_page_is_never_a_verdict(self):
        assert _verdict_from_overview(self.BLOCK_PAGE, market=MARKET)[0] == "unknown"

    def test_block_page_reason_names_the_cause(self):
        _, reason = _verdict_from_overview(self.BLOCK_PAGE, market=MARKET)
        assert "block" in reason.lower() or "captcha" in reason.lower()

    def test_echoed_query_alone_does_not_create_a_distributor(self):
        """The echoed URL is the only place 'distributor' appears."""
        assert "distributor" != _verdict_from_overview(self.BLOCK_PAGE, market=MARKET)[0]

    def test_recaptcha_marker(self):
        assert _verdict_from_overview(
            "Please complete the reCAPTCHA to continue.", market=MARKET
        )[0] == "unknown"

    def test_a_real_answer_is_unaffected(self):
        assert verdict("Acme is a distributor of wearable medical devices.") == "distributor"
