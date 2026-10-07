"""Cap on VERIFIED DISTRIBUTORS, not on candidates discovered.

MAX_COMPANIES_PER_MARKET caps discovery, and only ~21% of candidates
survive screening — so setting it to 500 yields ~105 distributors, not 500.
TARGET_DISTRIBUTORS_PER_MARKET caps the accepted rows instead.

The rule that must never bend: reaching the target is not a reason to
accept a manufacturer. A market with few real distributors finishes UNDER
target rather than padding the file.
"""
from __future__ import annotations

import pytest

from vendor_intel.discovery import kompass_distributors as kd


class TestTargetSetting:
    def test_unset_means_no_cap(self, monkeypatch):
        monkeypatch.delenv("TARGET_DISTRIBUTORS_PER_MARKET", raising=False)
        assert kd.target_distributors_per_market() == 0

    def test_zero_means_no_cap(self, monkeypatch):
        monkeypatch.setenv("TARGET_DISTRIBUTORS_PER_MARKET", "0")
        assert kd.target_distributors_per_market() == 0

    def test_value_is_respected(self, monkeypatch):
        monkeypatch.setenv("TARGET_DISTRIBUTORS_PER_MARKET", "500")
        assert kd.target_distributors_per_market() == 500

    def test_garbage_falls_back_to_no_cap(self, monkeypatch):
        monkeypatch.setenv("TARGET_DISTRIBUTORS_PER_MARKET", "abc")
        assert kd.target_distributors_per_market() == 0

    def test_is_independent_of_the_candidate_cap(self, monkeypatch):
        """The two knobs must not be confused: one caps discovery, the
        other caps accepted rows."""
        monkeypatch.setenv("MAX_COMPANIES_PER_MARKET", "500")
        monkeypatch.setenv("TARGET_DISTRIBUTORS_PER_MARKET", "500")
        assert kd.max_companies_per_market() == 500
        assert kd.target_distributors_per_market() == 500


class TestStopsAtTarget:
    @pytest.fixture(autouse=True)
    def _screening_path(self, monkeypatch):
        monkeypatch.setenv("KOMPASS_GOOGLE_AI_SCREEN", "true")

    def _candidates(self, n):
        return [
            kd.KompassCompanyCandidate(
                name=f"Company {i}", profile_url=f"https://kompass.com/c/co-{i}/md{i:06d}/"
            )
            for i in range(n)
        ]

    def _run(self, monkeypatch, *, n_candidates, accept_per_chunk, target):
        monkeypatch.setenv("TARGET_DISTRIBUTORS_PER_MARKET", str(target))
        candidates = self._candidates(n_candidates)
        monkeypatch.setattr(kd, "find_market_product_page", lambda m: ("https://k/p/1", "1"))
        monkeypatch.setattr(kd, "discover_distributor_candidates", lambda *a, **k: candidates)

        seen = {"chunks": 0}

        def fake_screen(chunk, *, market, kompass_product_id):
            seen["chunks"] += 1
            # IDs must be unique across chunks, or dedup collapses them and
            # the test measures dedup rather than the target cap.
            found = [
                kd.VerifiedDistributor(
                    company_name=c.name,
                    market=market,
                    kompass_company_id=c.profile_url.rstrip("/").rsplit("/", 1)[-1].upper(),
                    kompass_profile_url=c.profile_url,
                    kompass_product_id=kompass_product_id,
                    kompass_company_type="Distributor",
                    is_actual_market_distributor=True,
                    verification_confidence=60,
                )
                for c in chunk[:accept_per_chunk]
            ]
            return found, len(chunk) - len(found), 0, set()

        monkeypatch.setattr(kd, "_screen_candidates_via_chatgpt", fake_screen)
        result = kd.run_kompass_distributor_search("Smartphones")
        return result, seen

    def test_stops_once_the_target_is_reached(self, monkeypatch):
        """1000 candidates, 50 accepted per 100-chunk, target 100 =>
        must stop after 2 chunks, not screen all 10."""
        result, seen = self._run(
            monkeypatch, n_candidates=1000, accept_per_chunk=50, target=100
        )
        assert len(result.distributors) == 100
        assert seen["chunks"] == 2

    def test_never_exceeds_the_target(self, monkeypatch):
        """Chunks are 100 wide, so the count can overshoot inside a chunk
        and must be trimmed back."""
        result, _ = self._run(
            monkeypatch, n_candidates=1000, accept_per_chunk=80, target=100
        )
        assert len(result.distributors) == 100

    def test_thin_market_finishes_under_target(self, monkeypatch):
        """THE IMPORTANT ONE: only 30 real distributors exist, target 500.
        The result must be 30 — never padded to reach the number."""
        result, _ = self._run(
            monkeypatch, n_candidates=300, accept_per_chunk=10, target=500
        )
        assert len(result.distributors) == 30

    def test_no_target_screens_everything(self, monkeypatch):
        result, seen = self._run(
            monkeypatch, n_candidates=500, accept_per_chunk=10, target=0
        )
        assert seen["chunks"] == 5
        assert len(result.distributors) == 50
