"""Website / email / phone extraction from Google AI Overview answers.

Google answers these three plainly ("The official website of X is minke.de",
"Email: info@minke.de", "Phone: +49 203 28101-0") but the heuristic
extractor had no patterns for them, so every such answer was thrown away —
which is why Website, Email and Office No. sat near zero across 431 rows
despite Google returning the data on the very first query.

The strings below are the REAL text the live backend returned for
Fritz Minke GmbH & Co. KG.
"""
from __future__ import annotations

from vendor_intel.enrichment.gap_fill.heuristic_extract import heuristic_extract_fields as ex

COMPANY = "Fritz Minke GmbH & Co. KG"

LIVE_WEBSITE_ANSWER = (
    "The official website of Fritz Minke GmbH & Co. KG is minke.de Company Details "
    "Address: Cecilienstrasse 31, 47051 Duisburg, Germany Email: info@minke.de "
    "Phone: +49 203 28101-0 Registration: Amtsgericht Duisburg, HRA 1482"
)


class TestLiveAnswer:
    def test_all_four_fields_come_out_of_one_answer(self):
        got = ex(LIVE_WEBSITE_ANSWER, company=COMPANY)
        assert got["website"] == "minke.de"
        assert got["email"] == "info@minke.de"
        assert got["office_email"] == "info@minke.de"
        assert got["phone"] == "+49 203 28101-0"


class TestWebsite:
    def test_company_name_with_periods_between_phrase_and_domain(self):
        """'Co. KG' and 'S.p.A.' contain periods; a filler span that stopped
        at '.' silently failed on exactly the companies in this dataset."""
        assert ex("The official website for Acme S.p.A. is acme.it")["website"] == "acme.it"

    def test_colon_form(self):
        assert ex("Website: acme.de")["website"] == "acme.de"

    def test_url_is_reduced_to_a_host(self):
        assert ex("Official website: https://www.acme.de/en")["website"] == "acme.de"

    def test_directory_citation_is_never_taken_as_the_company_site(self):
        """Google cites D&B/LinkedIn beside its answer. Those hosts must
        never be recorded as the company's own website."""
        assert "website" not in ex("See dnb.com profile. Website: dnb.com")

    def test_no_website_claim_means_no_value(self):
        assert "website" not in ex("A distributor of medical devices in Germany.")


class TestEmail:
    def test_generic_inbox_fills_office_email(self):
        got = ex("Email: info@acme.de")
        assert got["office_email"] == "info@acme.de"

    def test_named_mailbox_preferred_for_the_person_email(self):
        got = ex("Email: info@acme.de Email: j.doe@acme.de. Website: acme.de")
        assert got["office_email"] == "info@acme.de"
        assert got["email"] == "j.doe@acme.de"

    def test_noreply_and_placeholder_addresses_are_ignored(self):
        assert "email" not in ex("Email: no-reply@acme.de")
        assert "email" not in ex("Email: someone@example.com")

    def test_prefers_an_address_on_the_company_domain(self):
        """A citation page can leak a third-party address into the blob."""
        got = ex("Website: acme.de. Contact info@acme.de or press@yumpu.com")
        assert got["email"].endswith("@acme.de")

    def test_image_filename_is_not_an_email(self):
        assert "email" not in ex("logo@2x.png")


class TestPhone:
    def test_international_format(self):
        assert ex("Phone: +49 203 28101-0")["phone"] == "+49 203 28101-0"

    def test_tel_label(self):
        assert ex("Tel: +39 040 7783111")["phone"] == "+39 040 7783111"

    def test_too_short_to_dial_is_rejected(self):
        assert "phone" not in ex("Phone: 123")

    def test_bare_number_without_a_label_is_not_taken(self):
        """A postcode or registry number must not become a phone number."""
        assert "phone" not in ex("Registration HRA 1482, 47051 Duisburg")


class TestNoRegression:
    def test_founded_and_hq_still_extract(self):
        got = ex("Fritz Minke GmbH & Co. KG was founded in 1912, based in Duisburg, Germany.")
        assert got["founded_year"] == "1912"
        assert got["headquarters"] == "Duisburg, Germany"

    def test_empty_text_is_safe(self):
        assert ex("") == {}
