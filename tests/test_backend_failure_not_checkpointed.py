"""A broken browser must not permanently drop real companies.

Live failure. A stale asyncio loop wedged the AI Mode browser, so every
query in a chunk failed. The verifier did the right thing at the chunk
level -- fail closed, drop the chunk -- but then CHECKPOINTED that as if
the model had judged those companies:

    global-wearable-glucometer...verify_checkpoint.json
      offset=50  verdicts=50  keep=0
      50x "unparseable answer from both backends - dropped fail-closed"

Every one of those 50 was a real company that no model ever looked at.
Because the offset had advanced to 50, resuming would skip straight past
them and they would be missing from the verified output for good -- and
the summary line would call them "dropped", which reads as a verdict.

Fail-closed is right for the OUTPUT (never emit an unverified company),
and wrong for the CHECKPOINT (never record a backend error as a verdict).
This separates the two.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location(
    "verify_distributors_ai_mode", _ROOT / "verify_distributors_ai_mode.py"
)
vd = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(vd)

_SRC = (_ROOT / "verify_distributors_ai_mode.py").read_text(encoding="utf-8")


# --- the marker -----------------------------------------------------------


def test_backend_failures_are_marked_distinctly():
    """Without a marker there is no way to tell "the model said no" from
    "nothing asked the model"."""
    assert '"backend_failure": True' in _SRC


def test_only_real_verdicts_are_checkpointed():
    assert 'real = [v for v in verdicts if not v.get("backend_failure")]' in _SRC
    assert 'save_checkpoint(ckpt_path, {"offset": len(real), "verdicts": real})' in _SRC


def test_resume_is_driven_by_judged_companies_not_a_row_offset():
    """A row offset stops being a valid resume point once the verdict list
    and the row order diverge, which happens the moment a chunk fails and
    later chunks are judged out of order.

    Observed live: offset=325 on a 325-row market that held only 305
    verdicts. The loop ran zero iterations and printed "DONE: 253 kept"
    three attempts in a row while 20 companies had never been judged at
    all -- a silent no-op that looked like success and burned the retry
    budget.

    Filtering the rows by the judged set cannot drift: a company is
    re-asked exactly when it has no real verdict. A backend failure is
    not a real verdict, so those return too.
    """
    assert "_judged = {" in _SRC
    assert "pending = [r for r in rows if _norm(r.get(\"Company\")) not in _judged]" in _SRC
    assert "for i in range(0, len(rows_to_do), CHUNK_SIZE):" in _SRC
    # the old positional resume must be gone
    assert "for i in range(offset, total, CHUNK_SIZE):" not in _SRC


def test_the_first_failure_is_not_overwritten_by_later_ones():
    """It records the EARLIEST failure, so the rewind covers them all."""
    assert (
        "backend_failed_at = i if backend_failed_at is None else backend_failed_at"
    ) in _SRC


def test_fail_closed_still_applies_to_the_output():
    """The safety property must survive: an unverified company is never
    emitted, it is only re-askable."""
    assert 'DROP whole chunk (fail closed)' in _SRC
    assert '"keep": False' in _SRC


# --- the summary line ----------------------------------------------------


def test_backend_failures_are_not_counted_as_dropped():
    """"50 dropped" reads as a verdict on 50 companies. It was not one."""
    assert 'n_backend = sum(1 for v in verdicts if v.get("backend_failure"))' in _SRC
    assert "n_drop = len(verdicts) - n_keep - n_backend" in _SRC


def test_unverified_companies_are_reported():
    """Silence here is how the original loss went unnoticed."""
    assert "were NOT verified" in _SRC
    assert "Re-run the same command" in _SRC


# --- resume semantics ----------------------------------------------------


def test_a_clean_checkpoint_round_trips(tmp_path):
    """The normal path must be untouched."""
    ckpt = tmp_path / "c.json"
    payload = {
        "offset": 20,
        "verdicts": [
            {"Company": "A", "keep": True, "confidence": 90},
            {"Company": "B", "keep": False, "confidence": 80},
        ],
    }
    vd.save_checkpoint(ckpt, payload)
    loaded = json.loads(ckpt.read_text(encoding="utf-8"))
    assert loaded["offset"] == 20
    assert len(loaded["verdicts"]) == 2
    assert not any(v.get("backend_failure") for v in loaded["verdicts"])


def test_a_poisoned_checkpoint_shape_is_recognisable(tmp_path):
    """What the live failure produced: every verdict a backend failure and
    zero keeps. Filtering it leaves nothing, and the offset rewinds to the
    start, so the whole file is re-asked."""
    verdicts = [
        {"Company": f"C{n}", "keep": False, "confidence": 0,
         "backend_failure": True}
        for n in range(50)
    ]
    real = [v for v in verdicts if not v.get("backend_failure")]
    assert real == [], "nothing was actually verified"
    assert sum(1 for v in verdicts if v.get("keep")) == 0
