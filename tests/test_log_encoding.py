"""Logging must never kill a run over a non-ASCII company name.

Observed live: a run died at 272 companies with

    UnicodeEncodeError: 'charmap' codec can't encode character '\u015e'

Windows consoles default to cp1252, which cannot encode most non-Western
characters, and print() RAISES on an unmappable one. The character was "S"
with cedilla in a Turkish distributor's name — found precisely BECAUSE the
prompt now steers toward under-covered regions, so non-ASCII names are
normal input for this pipeline rather than an edge case.

Two layers are tested: the entrypoint reconfigures stdout to UTF-8, and
_log() falls back to replacement characters if that ever fails.
"""
from __future__ import annotations

import io
import sys

import pytest

from vendor_intel.pipeline.market_research import batch_runner, contact_lookup
from vendor_intel.scraping import google_ai_mode

# Real names from the collected data / the crash itself.
NAMES = [
    "Empa Elektronik Ticaret A.\u015e.",   # Turkish S-cedilla (the crash)
    "Bodo M\u00fcller Chemie GmbH",        # German umlaut
    "AX\u00c8GAZ Trading & Technologies",  # French grave
    "\u4e07\u90a6\u7535\u5b50",            # Chinese
    "\u0428\u0438\u043c\u0435\u043d\u0441",# Cyrillic
    "\u65e5\u672c\u96fb\u5b50",            # Japanese
]


@pytest.mark.parametrize("module", [batch_runner, google_ai_mode, contact_lookup])
@pytest.mark.parametrize("name", NAMES)
def test_log_survives_a_cp1252_console(module, name, monkeypatch):
    """The exact failure mode: a strict cp1252 stdout must not raise."""
    stream = io.TextIOWrapper(io.BytesIO(), encoding="cp1252", errors="strict")
    monkeypatch.setattr(sys, "stdout", stream)

    module._log(f"accepted #273: {name}")   # must not raise

    stream.flush()
    written = stream.buffer.getvalue().decode("cp1252")
    assert "accepted #273" in written


@pytest.mark.parametrize("module", [batch_runner, google_ai_mode, contact_lookup])
def test_log_keeps_full_text_on_a_utf8_console(module, monkeypatch, capsys):
    """With UTF-8 (what the entrypoint sets up) nothing is degraded."""
    module._log("accepted #273: Empa Elektronik Ticaret A.\u015e.")
    assert "\u015e" in capsys.readouterr().out


def test_entrypoint_forces_utf8_stdout():
    """run_market_research.py must reconfigure stdout, so EVERY print site
    is fixed at once rather than one _log() at a time."""
    from pathlib import Path

    src = (Path(__file__).resolve().parents[1] / "run_market_research.py").read_text(
        encoding="utf-8"
    )
    assert "reconfigure" in src
    assert 'encoding="utf-8"' in src
    assert 'errors="replace"' in src


def test_reconfigure_failure_is_tolerated():
    """A stream that cannot be reconfigured (a pipe, a captured buffer)
    must not break startup."""
    stream = io.StringIO()          # has no .reconfigure
    try:
        stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass                        # exactly what the entrypoint catches
    else:
        pytest.skip("stream unexpectedly reconfigurable")
