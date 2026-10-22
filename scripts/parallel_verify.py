#!/usr/bin/env python3
"""Verify one market's discovery export with N browsers in parallel.

The launcher's verify step runs ONE verifier (one browser) per market, so a
300-company round takes 35-45 min. This splits the still-unjudged companies
into N part files, verifies them in parallel (own browser profile each), then
merges every part's verdicts into the market's own checkpoint and runs the
normal verifier once on the full file -- with nothing left to judge it only
writes <market>_verified_distributors.xlsx. The result is identical to a
single-verifier run, so the launcher continues from it as usual.

Usage:
    python scripts/parallel_verify.py --market "U.S. Entresto Market" --parts 10
Resumable: re-running reuses part files and their checkpoints.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from vendor_intel.pipeline.market_research.state import market_slug  # noqa: E402

VERIFIER = ROOT / "verify_distributors_ai_mode.py"


def log(msg: str) -> None:
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def _norm(s: object) -> str:
    return "".join(ch for ch in str(s or "").lower() if ch.isalnum())


def _load_ckpt(p: Path) -> dict:
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"offset": 0, "verdicts": []}


def _good(verdicts: list[dict]) -> list[dict]:
    return [v for v in verdicts if not v.get("backend_failure")]


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--market", required=True)
    ap.add_argument("--parts", type=int, default=10)
    args = ap.parse_args()

    stem = market_slug(args.market)
    src = ROOT / "output" / "market_research" / f"{stem}.xlsx"
    main_ck = src.with_name(f"{stem}_verify_checkpoint.json")
    work = ROOT / "output" / "market_research" / f"_parallel_{stem}"
    work.mkdir(parents=True, exist_ok=True)

    # Header row + rows of the discovery export.
    wb = openpyxl.load_workbook(src, read_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    hdr = next(i for i, r in enumerate(rows[:25]) if r and "Company" in [str(c).strip() for c in r if c is not None])
    head = list(rows[hdr])
    ci = [str(c).strip() if c is not None else "" for c in head].index("Company")
    data = [r for r in rows[hdr + 1:] if r and r[ci]]

    judged = {_norm(v.get("Company")) for v in _good(_load_ckpt(main_ck).get("verdicts", []))}
    for ck in work.glob("*_verify_checkpoint.json"):
        judged |= {_norm(v.get("Company")) for v in _good(_load_ckpt(ck).get("verdicts", []))}
    todo = [r for r in data if _norm(r[ci]) not in judged]
    log(f"{args.market}: {len(data)} companies, {len(judged)} already judged, {len(todo)} to verify")

    parts: list[Path] = []
    if todo:
        size = -(-len(todo) // args.parts)
        for i in range(args.parts):
            chunk = todo[i * size:(i + 1) * size]
            if not chunk:
                break
            p = work / f"{stem}-part{i + 1}.xlsx"
            out = openpyxl.Workbook()
            out.active.append(head)
            for r in chunk:
                out.active.append(list(r))
            out.save(p)
            parts.append(p)
        procs = []
        for i, p in enumerate(parts):
            env = dict(os.environ, PYTHONIOENCODING="utf-8",
                       VERIFY_CHUNK=os.environ.get("VERIFY_CHUNK", "8"),
                       GOOGLE_AI_MODE_MAX_RESETS=os.environ.get("GOOGLE_AI_MODE_MAX_RESETS", "10"),
                       GOOGLE_AI_MODE_PROFILE_DIR=str(ROOT / "data" / f"ai_mode_pverify_{i + 1}"))
            logf = (work / f"part{i + 1}.log").open("a", encoding="utf-8")
            procs.append(subprocess.Popen(
                [sys.executable, "-u", str(VERIFIER), "--file", str(p), "--market-name", args.market],
                cwd=str(ROOT), env=env, stdout=logf, stderr=subprocess.STDOUT))
            log(f"part{i + 1}: {p.name} ({sum(1 for _ in openpyxl.load_workbook(p, read_only=True).active.iter_rows()) - 1} companies)")
            time.sleep(15)  # stagger browser launches
        for pr in procs:
            pr.wait()

    # Merge every part's verdicts into the market's own checkpoint.
    merged = _load_ckpt(main_ck)
    have = {_norm(v.get("Company")) for v in _good(merged.get("verdicts", []))}
    merged["verdicts"] = _good(merged.get("verdicts", []))
    added = 0
    for ck in sorted(work.glob("*_verify_checkpoint.json")):
        for v in _good(_load_ckpt(ck).get("verdicts", [])):
            k = _norm(v.get("Company"))
            if k and k not in have:
                merged["verdicts"].append(v)
                have.add(k)
                added += 1
    if main_ck.exists():
        shutil.copy2(main_ck, main_ck.with_suffix(f".before_parallel_{time.strftime('%Y%m%d_%H%M%S')}.json"))
    main_ck.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"merged {added} verdicts -> {main_ck.name} ({len(merged['verdicts'])} total)")

    # One normal verifier pass: nothing (or almost nothing) left to judge, so
    # it just writes the verified workbook the launcher reads.
    env = dict(os.environ, PYTHONIOENCODING="utf-8", VERIFY_CHUNK=os.environ.get("VERIFY_CHUNK", "8"),
               GOOGLE_AI_MODE_PROFILE_DIR=str(ROOT / "data" / "ai_mode_pverify_1"))
    subprocess.run([sys.executable, "-u", str(VERIFIER), "--file", str(src), "--market-name", args.market],
                   cwd=str(ROOT), env=env)
    log("done")


if __name__ == "__main__":
    main()
