"""Employees and Ownership extraction from Google AI Overview prose.

Both columns were ASKED for on every company and both came back near-zero
across 431 rows. The cause was not Google's silence — it answers these
plainly — but the parser:

  * Employees required the literal word "employees" straight after the
    number, so "employs approximately 16 people" matched nothing.
  * Ownership had NO extraction pattern at all, guaranteeing an empty
    column no matter what Google returned.

The strings below are the real live answers for Fritz Minke GmbH & Co. KG.
"""
from __future__ import annotations

from vendor_intel.discovery.chatgpt_company_screen import _extract_fields_from_prose as ex

LIVE_EMPLOYEES = (
    "Fritz Minke GmbH & Co. KG employs approximately 16 people. Some business "
    "directories estimate the employee count between 10 and 49. The company is "
    "headquartered in Duisburg, Germany."
)
LIVE_OWNERSHIP = (
    "FRITZ MINKE GmbH & Co. KG is a private, family-owned company. It is not listed "
    "on any public stock exchange. The business has been managed across generations, "
    "maintaining a traditional family-run corporate structure for over 112 years."
)


class TestEmployees:
    def test_live_answer_number_after_the_verb(self):
        assert ex(LIVE_EMPLOYEES, ["Employees"])["Employees"] == "16"

    def test_plain_n_employees_still_works(self):
        assert ex("The company has 250 employees.", ["Employees"])["Employees"] == "250"

    def test_range_phrasing(self):
        assert ex("Employee count between 10 and 49.", ["Employees"])["Employees"] == "10-49"

    def test_thousands_separator(self):
        assert ex("Acme employs around 1,200 staff.", ["Employees"])["Employees"] == "1,200"

    def test_not_asked_means_not_extracted(self):
        assert "Employees" not in ex(LIVE_EMPLOYEES, ["Website"])

    def test_no_headcount_gives_nothing(self):
        assert ex("Acme distributes medical devices.", ["Employees"]) == {}


class TestOwnership:
    def test_live_answer_yields_both_labels(self):
        got = ex(LIVE_OWNERSHIP, ["Ownership"])["Ownership"]
        assert "Family-owned" in got
        assert "Private" in got

    def test_not_listed_on_an_exchange_is_not_public(self):
        """'It is not listed on any public stock exchange' contains the word
        'public' — reading that as a public company inverts the answer."""
        assert "Public" not in ex(LIVE_OWNERSHIP, ["Ownership"])["Ownership"]

    def test_publicly_traded(self):
        assert ex("Acme is publicly traded on the NYSE.", ["Ownership"])["Ownership"] == "Public"

    def test_subsidiary(self):
        assert ex("Acme is a subsidiary of Siemens AG.", ["Ownership"])["Ownership"] == "Subsidiary"

    def test_state_owned(self):
        assert ex("Acme is a state-owned enterprise.", ["Ownership"])["Ownership"] == "State-owned"

    def test_value_is_a_label_not_a_sentence(self):
        """The column must stay sortable in a spreadsheet."""
        got = ex(LIVE_OWNERSHIP, ["Ownership"])["Ownership"]
        assert len(got) < 40 and "." not in got

    def test_no_ownership_signal_gives_nothing(self):
        assert ex("Acme distributes wearable devices.", ["Ownership"]) == {}
