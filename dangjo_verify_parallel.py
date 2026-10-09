#!/usr/bin/env python3
"""Verify the Dangjo list across N browsers at once instead of one.

A single verifier does 12 companies per query with a ~12s pace, so 739
companies is roughly an hour. Each browser drives its own Chromium
profile and a profile takes an exclusive lock, so N verifiers run
genuinely in parallel -- the same trick that lets 25 drone markets verify
simultaneously.

Companies are split into N contiguous slices, each written to its own
workbook so each worker gets its own checkpoint and they never write to
the same file. Already-judged companies are skipped up front, so
restarting costs nothing.

Results are merged back into the market's real checkpoint at the end,
which is what export_xlsx and the supervisor read.

Usage:
    python dangjo_verify_parallel.py --workers 5
    python dangjo_verify_parallel.py --workers 5 --merge-only
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

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from openpyxl import load_workbook  # noqa: E402

from run_dangjo import MARKET  # noqa: E402
from vendor_intel.pipeline.market_research.state import market_slug  # noqa: E402

OUT = ROOT / "output" / "market_research"
SLUG = market_slug(MARKET)
SRC = OUT / f"{SLUG}.xlsx"
CKPT = OUT / f"{SLUG}_verify_checkpoint.json"
MARKET_NAME = "Dangjo Pepper Ingredients Market"


def judged_names() -> set[str]:
    if not CKPT.is_file():
        return set()
    try:
        data = json.loads(CKPT.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return set()
    return {str(v.get("Company") or "").strip().lower()
            for v in (data.get("verdicts") or [])}


def build_slices(workers: int) -> list[Path]:
    """One workbook per worker, holding only unjudged companies."""
    done = judged_names()
    wb = load_workbook(SRC)
    ws = wb.active
    hrow = next(r for r in range(1, 15) if ws.cell(r, 1).value == "#")
    ncol = ws.max_column

    rows = []
    for r in range(hrow + 1, ws.max_row + 1):
        name = ws.cell(r, 2).value
        if not name:
            continue
        if str(name).strip().lower() in done:
            continue
        rows.append([ws.cell(r, c).value for c in range(1, ncol + 1)])

    print(f"{len(rows)} companies still need a verdict "
          f"({len(done)} already judged)")
    if not rows:
        return []

    per = -(-len(rows) // workers)          # ceiling divide
    paths: list[Path] = []
    for i in range(workers):
        chunk = rows[i * per:(i + 1) * per]
        if not chunk:
            continue
        dst = OUT / f"dangjo-par-{i + 1}.xlsx"
        shutil.copy2(SRC, dst)
        w2 = load_workbook(dst)
        s2 = w2.active
        # Clear the data block, then write only this worker's slice.
        for r in range(hrow + 1, s2.max_row + 1):
            for c in range(1, ncol + 1):
                s2.cell(r, c).value = None
        for j, vals in enumerate(chunk, 1):
            s2.cell(hrow + j, 1).value = j
            for c in range(2, ncol + 1):
                s2.cell(hrow + j, c).value = vals[c - 1]
        w2.save(dst)
        # A stale checkpoint would make the worker skip its own slice.
        ck = OUT / f"{dst.stem}_verify_checkpoint.json"
        if ck.is_file():
            ck.unlink()
        paths.append(dst)
        print(f"  worker {i + 1}: {len(chunk)} companies -> {dst.name}")
    return paths


def launch(paths: list[Path]) -> list[subprocess.Popen]:
    procs = []
    for i, p in enumerate(paths, 1):
        env = dict(os.environ)
        # Each worker needs its OWN profile: a Chromium profile takes an
        # exclusive lock, and two workers on one profile produce
        # about:blank tabs and no results.
        prof = ROOT / "data" / f"dangjo_par{i}"
        env["GOOGLE_AI_MODE_PROFILE_DIR"] = str(prof)
        env["GOOGLE_AI_MODE_QUERY_DELAY"] = "14"
        env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "150"
        # Stage the captcha solver as a sibling of the profile, which is
        # where the loader looks and where it survives a profile reset.
        ext_src = ROOT / "data" / "ai_mode_chrome_profile_extensions"
        ext_dst = ROOT / "data" / f"dangjo_par{i}_extensions"
        if ext_src.is_dir() and not ext_dst.is_dir():
            shutil.copytree(ext_src, ext_dst)
        log = (ROOT / f"logs_dangjo_par{i}.log").open(
            "w", encoding="utf-8", errors="replace")
        procs.append(subprocess.Popen(
            [sys.executable, "-u", str(ROOT / "verify_distributors_ai_mode.py"),
             "--file", str(p), "--market-name", MARKET_NAME],
            cwd=str(ROOT), env=env, stdout=log, stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL))
        print(f"  worker {i} pid {procs[-1].pid} -> logs_dangjo_par{i}.log")
        time.sleep(3)
    return procs


def merge(workers: int) -> None:
    """Fold every worker's verdicts into the market's real checkpoint."""
    data = {"offset": 0, "verdicts": []}
    if CKPT.is_file():
        try:
            data = json.loads(CKPT.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    verdicts = data.get("verdicts") or []
    have = {str(v.get("Company") or "").strip().lower() for v in verdicts}

    added = 0
    for i in range(1, workers + 1):
        ck = OUT / f"dangjo-par-{i}_verify_checkpoint.json"
        if not ck.is_file():
            continue
        try:
            part = json.loads(ck.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            continue
        for v in part.get("verdicts") or []:
            key = str(v.get("Company") or "").strip().lower()
            if key and key not in have:
                verdicts.append(v)
                have.add(key)
                added += 1

    if CKPT.is_file():
        shutil.copy2(CKPT, str(CKPT) + ".pre_merge")
    CKPT.write_text(json.dumps({"offset": len(verdicts),
                                "verdicts": verdicts},
                               ensure_ascii=False, indent=1),
                    encoding="utf-8")
    kept = sum(1 for v in verdicts if v.get("keep"))
    print(f"merged +{added} verdicts -> {len(verdicts)} judged, {kept} kept")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--workers", type=int, default=5)
    ap.add_argument("--merge-only", action="store_true")
    args = ap.parse_args()

    if args.merge_only:
        merge(args.workers)
        return

    paths = build_slices(args.workers)
    if not paths:
        print("nothing to do")
        return
    print("")
    launch(paths)
    print("")
    print("  python dangjo_verify_parallel.py --merge-only   # when done")


if __name__ == "__main__":
    main()
