#!/usr/bin/env python3
"""Keep topping all six markets up to 400 VERIFIED, surviving crashes.

A Python loop rather than a PowerShell one: two PowerShell watchdogs died
on parse errors tonight before doing any work, and a wrapper whose own
syntax can silently kill an overnight run is the wrong tool. This runs the
same command and re-enters it, which is all the wrapper ever needed to do.

Everything underneath checkpoints — discovery per batch, verification per
chunk of 12 — so re-entering costs nothing and resumes where it stopped.

Refusals and rate limits no longer need a human: discovery and
verification both rebuild the browser profile when AI Mode declines,
because the refusal lives in the session cookies and a browser restart
alone never cleared it. captcha-raptor lives outside the profile and
survives every reset.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

MAX_ATTEMPTS = 10


def shortfall() -> int:
    out = subprocess.run(
        [sys.executable, str(ROOT / "verified_shortfall.py")],
        cwd=str(ROOT), capture_output=True, text=True,
    ).stdout.strip()
    try:
        return int(out.splitlines()[-1])
    except (ValueError, IndexError):
        return -1


def main() -> None:
    # --only narrows the retry loop to the markets still short. The four
    # markets already past 400 re-verify nothing but still pay the cost of
    # a full discovery pass, which is ~40 minutes of the attempt budget
    # spent to print "already at target" four times.
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--only", action="append", default=[], metavar="TEXT")
    args = ap.parse_args()
    only = [a for text in args.only for a in ("--only", text)]

    for attempt in range(1, MAX_ATTEMPTS + 1):
        print("", flush=True)
        print("#" * 64, flush=True)
        print(f"# WATCHDOG attempt {attempt}/{MAX_ATTEMPTS}  "
              f"{datetime.now():%H:%M:%S}", flush=True)
        print("#" * 64, flush=True)

        subprocess.run(
            [sys.executable, str(ROOT / "top_up_to_target.py"),
             "--target", "400", *only],
            cwd=str(ROOT),
        )

        left = shortfall()
        if left == 0:
            print("\nAll six markets reached 400 verified.", flush=True)
            break
        print(f"\n{left} market(s) still short — retrying (progress is saved)",
              flush=True)
        time.sleep(60)

    print("", flush=True)
    print("-" * 62, flush=True)
    print("FINAL VERIFIED COUNTS", flush=True)
    print("-" * 62, flush=True)
    subprocess.run(
        [sys.executable, str(ROOT / "verified_shortfall.py"), "--report"],
        cwd=str(ROOT),
    )


if __name__ == "__main__":
    main()
