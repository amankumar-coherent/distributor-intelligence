"""Resetting verification must never destroy discovery.

Discovery swept the global listing plus 60 country listings to build 9,460
candidates over many hours; the verify stage is cheap to redo by comparison.
A reset that dropped candidates would silently throw that away, and the
resumed run would re-scrape Kompass from scratch.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("reset_verification", ROOT / "reset_verification.py")
reset = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reset)


def test_discovery_keys_are_kept_and_verify_keys_are_cleared():
    """The two lists must be disjoint, or a key would be both kept and
    cleared and the outcome would depend on dict ordering."""
    assert not set(reset._KEEP) & set(reset._CLEAR)


def test_candidates_are_on_the_keep_list():
    assert "kompass_candidates" in reset._KEEP


def test_every_verify_artifact_is_cleared():
    """Anything the verify stage writes must be reset, otherwise a stale
    accepted list would survive and the re-run would look already-done."""
    for key in (
        "kompass_processed_urls",
        "kompass_accepted",
        "kompass_rejected_count",
        "kompass_unreachable_count",
    ):
        assert key in reset._CLEAR


def test_counts_reset_to_zero_not_to_empty_list():
    """These are ints in the checkpoint; clearing them to [] would break
    the int() coercion on resume."""
    assert reset._CLEAR["kompass_rejected_count"] == 0
    assert reset._CLEAR["kompass_unreachable_count"] == 0


def test_lists_reset_to_empty_lists():
    assert reset._CLEAR["kompass_processed_urls"] == []
    assert reset._CLEAR["kompass_accepted"] == []
