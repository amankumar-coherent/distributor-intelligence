"""Quarantine bookkeeping for the market-relevance re-screen.

The danger in this script is silent data loss: it removes companies from a
deliverable the user has already been shown. So it must never delete, must
be reversible, and a transient error must never be read as evidence against
a company.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("rescreen_accepted", ROOT / "rescreen_accepted.py")
rescreen = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rescreen)


class FakeCkpt:
    """Minimal stand-in for ExpandCheckpoint."""

    def __init__(self, **data):
        self._d = dict(data)
        self.bumps = 0

    def data(self, key, default=None):
        return self._d.get(key, default)

    def bump(self, step, substep, *, progress=None, note="", **data):
        self._d.update(data)
        self.bumps += 1


def rec(name, url, **kw):
    base = {"company_name": name, "kompass_profile_url": url}
    base.update(kw)
    return base


class TestSave:
    def test_rejected_company_moves_out_of_accepted(self):
        accepted = [rec("Good Co", "u1"), rec("Minke", "u2")]
        ck = FakeCkpt(kompass_accepted=accepted)
        rescreen._save(ck, accepted, {"u2"}, {"u1", "u2"}, "Wearable Medical Devices")
        assert [d["company_name"] for d in ck.data("kompass_accepted")] == ["Good Co"]
        assert [d["company_name"] for d in ck.data("kompass_rescreen_rejected")] == ["Minke"]

    def test_nothing_is_ever_deleted(self):
        """Every input company must still exist somewhere afterwards."""
        accepted = [rec("A", "u1"), rec("B", "u2"), rec("C", "u3")]
        ck = FakeCkpt(kompass_accepted=accepted)
        rescreen._save(ck, accepted, {"u2", "u3"}, {"u1", "u2", "u3"}, "M")
        surviving = {d["company_name"] for d in ck.data("kompass_accepted")} | {
            d["company_name"] for d in ck.data("kompass_rescreen_rejected")
        }
        assert surviving == {"A", "B", "C"}

    def test_quarantine_does_not_duplicate_across_saves(self):
        """Periodic saves re-scan the same list; a company must not be
        quarantined twice."""
        accepted = [rec("A", "u1"), rec("B", "u2")]
        ck = FakeCkpt(kompass_accepted=accepted)
        rescreen._save(ck, accepted, {"u2"}, {"u1", "u2"}, "M")
        rescreen._save(ck, accepted, {"u2"}, {"u1", "u2"}, "M")
        assert len(ck.data("kompass_rescreen_rejected")) == 1

    def test_progress_is_recorded_for_resume(self):
        accepted = [rec("A", "u1")]
        ck = FakeCkpt(kompass_accepted=accepted)
        rescreen._save(ck, accepted, set(), {"u1"}, "M")
        assert ck.data("kompass_rescreened_urls") == ["u1"]

    def test_no_rejections_leaves_accepted_untouched(self):
        accepted = [rec("A", "u1"), rec("B", "u2")]
        ck = FakeCkpt(kompass_accepted=accepted)
        rescreen._save(ck, accepted, set(), {"u1", "u2"}, "M")
        assert len(ck.data("kompass_accepted")) == 2
        assert ck.data("kompass_rescreen_rejected") == []


class TestReport:
    def test_empty_quarantine_reports_cleanly(self, capsys):
        assert rescreen._report([]) == 0
        assert "empty" in capsys.readouterr().out.lower()

    def test_lists_each_quarantined_company_with_its_reason(self, capsys):
        rescreen._report([rec("Minke", "u2", rejection_reason="not in Wearable Medical Devices")])
        out = capsys.readouterr().out
        assert "Minke" in out
        assert "not in Wearable" in out
