#!/usr/bin/env python3
"""Verify the 501 AP Automation companies across 10 parallel browsers.

WHY SHARDING BY FILE, NOT BY SHARED-CHECKPOINT CONCURRENCY
------------------------------------------------------------
verify_distributors_ai_mode.py's save_checkpoint() is a blind overwrite:

    def save_checkpoint(ckpt_path, data):
        ckpt_path.write_text(json.dumps(data, ...))

No file lock, no merge-with-what's-on-disk. That is NOT the same
protected path as the discovery pipeline's state.py (which merges under
a file lock specifically because concurrent discovery workers write the
SAME market). Here, 10 processes pointed at the SAME xlsx + SAME
checkpoint would each load the full "pending" list, verify different
companies, and then overwrite each other's saved verdicts on every
chunk -- exactly the bug state.py was fixed for, except nothing fixes
it on this path.

So this script shards at the FILE level instead: split the 501 rows into
10 separate single-sheet workbooks up front (each gets its OWN
checkpoint file by construction, so no two processes ever touch the same
JSON), run verify_ap_automation.py's machinery against each shard in its
own process with its own Chromium profile, then merge the ten
*_verified_distributors.xlsx outputs back into one file matching the
original row order.

Usage:
    python verify_ap_automation_10_browsers.py
    python verify_ap_automation_10_browsers.py --workers 5
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import verify_ap_automation as A  # noqa: E402
import verify_distributors_ai_mode as V  # noqa: E402

SHARD_DIR = ROOT / "output" / "market_research" / "ap_shards"
LOGS = ROOT / "logs_ap_verify"
LOGS.mkdir(exist_ok=True)
_EXTENSION_MASTER = ROOT / "data" / "ai_mode_chrome_profile_extensions"


def log(msg: str) -> None:
    print(msg, flush=True)


def stage_extension(suffix: str) -> Path:
    dst = ROOT / "data" / f"ap_verify_profile_{suffix}"
    ext = ROOT / "data" / f"ap_verify_profile_{suffix}_extensions"
    if not ext.is_dir() and _EXTENSION_MASTER.is_dir():
        shutil.copytree(_EXTENSION_MASTER, ext)
    return dst


def make_shards(n_shards: int) -> list[Path]:
    """Split the source workbook into N single-sheet shard files.

    Round-robin, not contiguous blocks: the source rows are in roughly
    the order workers discovered them (clustered by country), so a
    contiguous split would hand one worker "every UK row" and another
    "every Singapore row" -- round-robin spreads country mix evenly,
    which matters if different countries take different CAPTCHA/latency
    profiles to verify.
    """
    SHARD_DIR.mkdir(parents=True, exist_ok=True)
    rows, headers, sheet_name = V.load_rows(A.XLSX_PATH)
    log(f"source: {len(rows)} companies -> {n_shards} shards")

    shards: list[list[dict]] = [[] for _ in range(n_shards)]
    for i, row in enumerate(rows):
        shards[i % n_shards].append(row)

    paths = []
    for i, shard_rows in enumerate(shards):
        out = SHARD_DIR / f"shard_{i}.xlsx"
        V.write_output(shard_rows, headers, sheet_name, out,
                       market_name=A.MARKET_NAME)
        paths.append(out)
        log(f"  shard {i}: {len(shard_rows)} companies -> {out.name}")
    return paths


def worker(idx: int, shard_path: Path) -> None:
    profile_dir = stage_extension(str(idx))
    log_path = LOGS / f"w{idx}.log"
    handle = log_path.open("a", encoding="utf-8", errors="replace")
    env = {
        **__import__("os").environ,
        "GOOGLE_AI_MODE_PROFILE_DIR": str(profile_dir),
        "AP_VERIFY_FILE_OVERRIDE": str(shard_path),
    }
    cmd = [sys.executable, "-u", str(ROOT / "verify_ap_automation.py")]
    try:
        subprocess.run(cmd, cwd=str(ROOT), env=env, stdout=handle,
                       stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                       timeout=3600)
    except subprocess.TimeoutExpired:
        pass
    except Exception as exc:  # noqa: BLE001
        handle.write(f"\n[launcher] worker {idx} error: {exc}\n")
    finally:
        handle.close()
    log(f"  [w{idx}] shard done")


def merge_results(n_shards: int) -> None:
    """Combine the N shard verified-xlsx outputs into one file, in the
    ORIGINAL source row order (not shard order), so the deliverable
    reads the same as before verification.
    """
    rows, headers, sheet_name = V.load_rows(A.XLSX_PATH)
    order = {V._norm(r.get("Company")): i for i, r in enumerate(rows)}

    merged: list[dict] = []
    total_kept = total_closed = total_acquired = 0
    for i in range(n_shards):
        verified_path = SHARD_DIR / f"shard_{i}_verified_distributors.xlsx"
        ckpt_path = SHARD_DIR / f"shard_{i}_verify_checkpoint.json"
        if verified_path.is_file():
            kept_rows, _, _ = V.load_rows(verified_path)
            merged.extend(kept_rows)
        if ckpt_path.is_file():
            data = json.loads(ckpt_path.read_text(encoding="utf-8"))
            verdicts = data.get("verdicts", [])
            total_closed += sum(
                1 for v in verdicts
                if str(v.get("status") or "").lower() == "closed")
            total_acquired += sum(1 for v in verdicts if v.get("acquired_by"))

    merged.sort(key=lambda r: order.get(V._norm(r.get("Company")), 1 << 30))

    out = A.XLSX_PATH.with_name(
        A.XLSX_PATH.stem + "_verified_distributors.xlsx")
    V.write_output(merged, headers, sheet_name, out,
                  market_name=A.MARKET_NAME)
    log("")
    log(f"MERGED: {len(merged)} kept of {len(rows)} "
        f"({total_closed} closed-dropped, {total_acquired} acquired) "
        f"-> {out.name}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--workers", type=int, default=10)
    args = ap.parse_args()

    # Same registration verify_ap_automation.py's main() does, and for the
    # same reason: set_distributor_definition() is in-memory per-process,
    # and THIS process (not that one) is what calls write_output() for
    # make_shards()/merge_results() below. Without it, every shard and the
    # final merged file gets the wrong "physical goods / EXCLUDED: channel
    # partners" banner despite being judged under the correct rule.
    from vendor_intel.pipeline.market_research import discovery as _discovery
    _discovery.set_distributor_definition(A.MARKET_NAME, "software")

    # Also replace the banner's leftover "EXCLUDED: ... system
    # integrators and consultancies..." clause (wrong for this market) --
    # see verify_ap_automation.py's _patched_definition_banner docstring
    # for why export_xlsx is patched directly rather than V.
    import vendor_intel.pipeline.market_research.export_xlsx as _export_xlsx
    _export_xlsx._real_definition_banner = _export_xlsx._definition_banner
    _export_xlsx._definition_banner = A._patched_definition_banner

    shard_paths = make_shards(args.workers)

    threads = []
    for i, path in enumerate(shard_paths):
        t = threading.Thread(target=worker, args=(i, path), daemon=True)
        threads.append(t)
        t.start()
        time.sleep(3)
    for t in threads:
        t.join()

    merge_results(args.workers)


if __name__ == "__main__":
    main()
