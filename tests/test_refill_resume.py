"""The column refill must resume, not restart.

Two bugs made a stopped refill lose everything:
  * the checkpoint was written only after ALL companies finished, so
    stopping part-way discarded the completed work;
  * nothing recorded WHICH companies were done, so the next run began at
    company 1 and re-paid Google for hundreds of answers it already had.

Observed live: a run stopped at ~70 of 2,306 restarted from the first
company with nothing saved.
"""
from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("refill_accepted", ROOT / "refill_accepted.py")
refill = importlib.util.module_from_spec(spec)
spec.loader.exec_module(refill)


class Args:
    all_columns = True
    restart = False


class FakeCompany:
    """Stands in for ScreenedCompany."""

    def __init__(self, name="", profile_url="", country_hint=""):
        self.name = name
        self.profile_url = profile_url
        self.country_hint = country_hint
        self.fields = {"Website": "acme.com"}


class Dist:
    def __init__(self, i):
        self.company_name = f"Company {i}"
        self.kompass_profile_url = f"https://kompass.com/c/co-{i}/md{i:06d}/"
        self.country = "IT"
        self.market = "Test Market"
        self.website = ""
        self.founded_year = ""
        self.headquarters = ""
        self.employees = ""
        self.contact_person = ""
        self.email = ""
        self.phone = ""
        self.brands = ""
        self.extra_columns = {}


COLUMNS = ["Website", "Founded"]


def run(subset, done_urls, saves, asked):
    def fake_fill(companies, *, columns, market):
        asked.extend(c.name for c in companies)
        return companies

    refill._refill(
        subset, Args(), fake_fill, FakeCompany, COLUMNS,
        save=lambda urls: saves.append(set(urls)),
        done_urls=done_urls,
    )


class TestResume:
    def test_already_refilled_companies_are_skipped(self):
        subset = [Dist(i) for i in range(5)]
        done = {subset[0].kompass_profile_url, subset[1].kompass_profile_url}
        asked = []
        run(subset, done, [], asked)
        assert asked == ["Company 2", "Company 3", "Company 4"]

    def test_nothing_done_means_everything_is_asked(self):
        subset = [Dist(i) for i in range(3)]
        asked = []
        run(subset, set(), [], asked)
        assert asked == ["Company 0", "Company 1", "Company 2"]

    def test_all_done_asks_google_for_nothing(self):
        """The whole point: a completed refill must cost zero queries."""
        subset = [Dist(i) for i in range(4)]
        done = {d.kompass_profile_url for d in subset}
        asked = []
        run(subset, done, [], asked)
        assert asked == []

    def test_completed_companies_are_recorded(self):
        subset = [Dist(i) for i in range(3)]
        done = set()
        run(subset, done, [], [])
        assert done == {d.kompass_profile_url for d in subset}


class TestPeriodicSave:
    def test_progress_is_saved_before_the_run_ends(self):
        """A crash at company 30 of 2306 must not lose the first 25."""
        subset = [Dist(i) for i in range(60)]
        saves = []
        run(subset, set(), saves, [])
        assert len(saves) >= 2, "expected periodic saves, not just a final one"

    def test_final_partial_batch_is_saved(self):
        subset = [Dist(i) for i in range(30)]
        saves = []
        run(subset, set(), saves, [])
        assert len(saves[-1]) == 30

    def test_each_save_grows_monotonically(self):
        subset = [Dist(i) for i in range(60)]
        saves = []
        run(subset, set(), saves, [])
        for a, b in zip(saves, saves[1:]):
            assert a.issubset(b)

    def test_no_save_callback_is_tolerated(self):
        """--limit style callers may pass nothing; must not crash."""
        subset = [Dist(i) for i in range(3)]
        refill._refill(subset, Args(), lambda c, **k: c, FakeCompany, COLUMNS)


class TestDeepFillResume:
    """Stage 3 crawls each company's website — the slowest stage in the
    pipeline. It previously saved nothing until every company finished, so
    stopping it discarded hours of work and the next run re-crawled from
    the first company."""

    def _subset(self, n):
        return [Dist(i) for i in range(n)]

    def _run(self, subset, done, saves, monkeypatch, visited):
        """Drive _deep_fill with the network layer stubbed out."""
        async def fake_gap_fill_row(row, **kw):
            visited.append(row.get("company"))
            return {}

        import vendor_intel.enrichment.gap_fill.orchestrator as orch
        monkeypatch.setattr(orch, "gap_fill_row", fake_gap_fill_row)

        refill._deep_fill(subset, save=lambda urls: saves.append(set(urls)), done_urls=done)

    def test_already_crawled_companies_are_skipped(self, monkeypatch):
        subset = self._subset(5)
        done = {subset[0].kompass_profile_url, subset[1].kompass_profile_url}
        visited, saves = [], []
        self._run(subset, done, saves, monkeypatch, visited)
        assert visited == ["Company 2", "Company 3", "Company 4"]

    def test_progress_saves_before_the_end(self, monkeypatch):
        subset = self._subset(60)
        visited, saves = [], []
        self._run(subset, set(), saves, monkeypatch, visited)
        assert len(saves) >= 2

    def test_crawl_marker_is_separate_from_the_refill_marker(self):
        """A company can be done for the Google-AI pass but not yet
        crawled; sharing one marker would skip real work."""
        src = (ROOT / "refill_accepted.py").read_text(encoding="utf-8")
        assert "kompass_refilled_urls" in src
        assert "kompass_crawled_urls" in src
