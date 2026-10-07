"""refill_contacts_ai_mode.py: --only-empty and --fields scoping.

Requested explicitly: refill ONLY the empty contact cells, for a chosen
subset of columns. Two guarantees matter enough to pin:

  1. An existing value is authoritative. It was either verified earlier or
     supplied by the user, so a fresh lookup must never overwrite or erase
     it — the whole point of the flag is that already-good data is safe.
  2. A column not named in --fields is not merely skipped but never
     written, so a run targeting five columns cannot blank the sixth.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

_SPEC = importlib.util.spec_from_file_location(
    "refill_contacts_ai_mode",
    Path(__file__).resolve().parents[1] / "refill_contacts_ai_mode.py",
)
refill = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(refill)

FIVE = ["Contact Person", "Role", "LinkedIn", "Office No.", "Office Email"]


class _Args:
    def __init__(self, **kw):
        self.dry_run = True
        self.only_empty = False
        self.market = "M"
        self._columns = list(refill._FIELD_TO_COLUMN.values())
        self.__dict__.update(kw)


def _row(**kw):
    row = {c: "" for c in refill._FIELD_TO_COLUMN.values()}
    row["Company"] = kw.pop("Company", "Acme Inc")
    row.update(kw)
    return row


def test_existing_value_is_never_overwritten(capsys):
    """A verified contact must survive a run that finds something else."""
    rows = [_row(**{"Contact Person": "Real Person", "Role": "VP Sales"})]
    results = {
        "Acme Inc": {
            "contact_person": "Someone Else", "role": "CTO", "email": "",
            "linkedin": "", "office_no": "", "office_email": "",
        }
    }
    refill._apply(_Args(only_empty=True, _columns=FIVE), rows, results, None)

    assert rows[0]["Contact Person"] == "Real Person"
    assert rows[0]["Role"] == "VP Sales"


def test_existing_value_is_never_blanked(capsys):
    """An empty answer must not erase a value that is already there."""
    rows = [_row(**{"Office No.": "+1 555 123 4567"})]
    results = {"Acme Inc": {f: "" for f in refill._FIELD_TO_COLUMN}}
    refill._apply(_Args(only_empty=True, _columns=FIVE), rows, results, None)

    assert rows[0]["Office No."] == "+1 555 123 4567"


def test_empty_field_is_filled(capsys):
    rows = [_row()]
    results = {
        "Acme Inc": {
            "contact_person": "Jane Doe", "role": "Head of Sales",
            "linkedin": "https://www.linkedin.com/in/jane-doe",
            "office_no": "+44 20 7946 0958", "office_email": "j.doe@acme.com",
            "email": "",
        }
    }
    refill._apply(_Args(only_empty=True, _columns=FIVE), rows, results, None)

    assert rows[0]["Contact Person"] == "Jane Doe"
    assert rows[0]["LinkedIn"] == "https://www.linkedin.com/in/jane-doe"
    assert rows[0]["Office Email"] == "j.doe@acme.com"


def test_unrequested_column_is_never_written(capsys):
    """Email is NOT in the five requested columns, so a value returned for
    it must be ignored — and an existing Email must not be blanked."""
    rows = [_row(**{"Email": "keep@acme.com"})]
    results = {
        "Acme Inc": {
            "email": "different@acme.com", "contact_person": "Jane Doe",
            "role": "VP", "linkedin": "", "office_no": "", "office_email": "",
        }
    }
    refill._apply(_Args(only_empty=True, _columns=FIVE), rows, results, None)

    assert rows[0]["Email"] == "keep@acme.com", "unrequested column touched"
    assert rows[0]["Contact Person"] == "Jane Doe"


def test_unrequested_column_not_blanked_even_without_only_empty(capsys):
    rows = [_row(**{"Email": "keep@acme.com"})]
    results = {"Acme Inc": {f: "" for f in refill._FIELD_TO_COLUMN}}
    refill._apply(_Args(only_empty=False, _columns=FIVE), rows, results, None)

    assert rows[0]["Email"] == "keep@acme.com"


def test_unanswered_company_is_left_alone(capsys):
    """A CAPTCHA/parse failure is not evidence a value is wrong."""
    rows = [_row(**{"Contact Person": "Existing"})]
    refill._apply(_Args(only_empty=True, _columns=FIVE), rows, {}, None)

    assert rows[0]["Contact Person"] == "Existing"


def test_full_mode_still_blanks_unverified(capsys):
    """Without --only-empty the original behaviour is intact: an
    unconfirmed value is removed rather than kept."""
    rows = [_row(**{"Contact Person": "Invented Person", "Role": "VP"})]
    results = {"Acme Inc": {f: "" for f in refill._FIELD_TO_COLUMN}}
    refill._apply(
        _Args(only_empty=False, _columns=list(refill._FIELD_TO_COLUMN.values())),
        rows, results, None,
    )

    assert rows[0]["Contact Person"] == ""
    assert rows[0]["Role"] == ""


def test_field_to_column_covers_the_five_requested():
    for col in FIVE:
        assert col in refill._FIELD_TO_COLUMN.values(), col
