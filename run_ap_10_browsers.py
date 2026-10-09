#!/usr/bin/env python3
"""Ten parallel browsers sweeping the AP Automation channel-partner market
country by country.

Two things make this faster than one browser:

  * COUNTRY ROTATION. An unrestricted global query keeps returning the same
    few dozen household names. Restricting each query to ONE country makes
    the model answer "partners IN GERMANY", surfacing the regional
    resellers and ERP integrators a global query never reaches. Same
    state/city rotation that took the fire and drone markets past 250.
  * TEN PROCESSES, NOT TEN THREADS. ask() holds a process-wide browser
    lock ("one query at a time per page"), so threads inside one process
    serialise and gain nothing. Each worker is its own process with its
    own GOOGLE_AI_MODE_PROFILE_DIR, because a Chromium profile takes an
    exclusive lock -- two processes sharing one produce about:blank tabs.

Every worker writes to the SAME market state file. That is safe: state.py
merges on save under a file lock rather than overwriting, which was built
and stress-tested for exactly this.

Usage:
    python run_ap_10_browsers.py --target 250
    python run_ap_10_browsers.py --target 250 --workers 10
"""
from __future__ import annotations

import argparse
import os
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

from ap_automation_market import MARKET, SLUG  # noqa: E402
from run_ap_by_country import COUNTRIES  # noqa: E402

LOGS = ROOT / "logs_ap"
LOGS.mkdir(exist_ok=True)
_EXTENSION_MASTER = ROOT / "data" / "ai_mode_chrome_profile_extensions"


def log(msg: str) -> None:
    print(msg, flush=True)


def stage_extension(suffix: str) -> None:
    """Give this worker's profile its own captcha-raptor copy.

    The extension must live in a SIBLING directory of the profile
    (<profile>_extensions), not inside it, so a profile reset -- which
    the scraper does on repeated failures -- cannot delete it.
    """
    dst = ROOT / "data" / f"ap_profile_{SLUG}_{suffix}_extensions"
    if dst.is_dir():
        return
    if _EXTENSION_MASTER.is_dir():
        shutil.copytree(_EXTENSION_MASTER, dst)


def discovered() -> int:
    from vendor_intel.pipeline.market_research.state import MarketResearchState
    return len(MarketResearchState.load(MARKET).companies)


def worker(idx: int, countries: list[str], target: int,
           batch_size: int) -> None:
    """One worker: its own profile, its own slice of the country list."""
    suffix = f"w{idx}"
    stage_extension(suffix)
    log_path = LOGS / f"{suffix}.log"

    for country in countries:
        if discovered() >= target:
            return
        handle = log_path.open("a", encoding="utf-8", errors="replace")
        cmd = [
            sys.executable, "-u", str(ROOT / "ap_automation_market.py"),
            "--target", str(target),
            "--batch-size", str(batch_size),
            "--region", country,
            "--profile-suffix", suffix,
        ]
        try:
            subprocess.run(cmd, cwd=str(ROOT), stdout=handle,
                           stderr=subprocess.STDOUT,
                           stdin=subprocess.DEVNULL, timeout=1800)
        except subprocess.TimeoutExpired:
            pass
        except Exception as exc:  # noqa: BLE001
            handle.write(f"\n[launcher] {country}: {exc}\n")
        finally:
            handle.close()
        log(f"  [{suffix}] {country} done — total {discovered()}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=250)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--workers", type=int, default=10)
    args = ap.parse_args()

    # Deal countries round-robin so each worker gets a mix of dense
    # (US, UK, Germany) and thin (Brazil, South Africa) markets rather
    # than one worker drawing every slow country.
    slices: list[list[str]] = [[] for _ in range(args.workers)]
    for i, c in enumerate(COUNTRIES):
        slices[i % args.workers].append(c)

    log(f"market : {MARKET}")
    log(f"target : {args.target} discovered")
    log(f"workers: {args.workers}, {len(COUNTRIES)} countries")
    log(f"start  : {discovered()} already discovered")
    log("")

    threads = []
    for i in range(args.workers):
        t = threading.Thread(
            target=worker,
            args=(i, slices[i], args.target, args.batch_size),
            daemon=True,
        )
        threads.append(t)
        t.start()
        time.sleep(3)  # stagger Chromium launches

    for t in threads:
        t.join()

    log("")
    log(f"DONE — {discovered()} companies discovered")

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx
    log(export_xlsx(MARKET))


if __name__ == "__main__":
    main()
