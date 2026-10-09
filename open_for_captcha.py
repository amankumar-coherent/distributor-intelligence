#!/usr/bin/env python3
"""Open the automation profile in a VISIBLE Chromium so a human can solve
a CAPTCHA, leaving the clearance cookies behind for the pipeline.

Why this exists: google.com/sorry/index blocks the profile until someone
passes the challenge. captcha-raptor clears most automatically, but when
Google escalates to "unusual traffic from your computer network" it wants
a real interaction.

Solving it in the SAME profile the pipeline uses is the whole point --
the clearance lives in that profile's cookies, so a CAPTCHA solved in
your everyday Chrome does nothing for the run.

Never run this alongside the pipeline: two processes on one profile
corrupt it. Stop the watchdog first.

Usage:
    python open_for_captcha.py
    python open_for_captcha.py --minutes 20
"""
from __future__ import annotations

import argparse
import os
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

from vendor_intel.scraping import google_ai_mode as g  # noqa: E402

# A real AI Mode query (udm=50), so passing the challenge proves the
# profile can actually reach the backend the pipeline uses -- not just
# google.com, which can pass while the AI Mode endpoint still blocks.
URL = "https://www.google.com/search?q=polyurethane+distributors&udm=50"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--minutes", type=float, default=15.0,
                    help="how long to hold the window open (default 15)")
    args = ap.parse_args()

    # Launch through the pipeline's OWN launcher so profile, flags and
    # staged extensions are identical -- a differently-configured browser
    # would write cookies the real run does not trust. Headed is already
    # the default (ai_mode_headless is opt-in), but force it explicitly:
    # the whole point is a window a human can click in.
    os.environ["GOOGLE_AI_MODE_HEADLESS"] = "false"
    page = g._get_page()  # noqa: SLF001
    profile = g.active_profile_dir()
    print(f"profile: {profile}")

    try:
        page.goto(URL, wait_until="domcontentloaded", timeout=60_000)
    except Exception as exc:  # noqa: BLE001
        print(f"navigation warning: {type(exc).__name__}: {exc}")

    print("")
    print("=" * 64)
    print("A Chromium window is open on the automation profile.")
    print("Solve the 'I'm not a robot' CAPTCHA there.")
    print("When normal results load, CLOSE THE WINDOW to save and exit.")
    print(f"Waiting up to {args.minutes:.0f} minutes.")
    print("=" * 64)

    deadline = time.monotonic() + args.minutes * 60
    try:
        while time.monotonic() < deadline:
            time.sleep(3)
            ctx = getattr(g, "_context", None)
            if ctx is None:
                break
            try:
                if not ctx.pages:      # window closed by hand
                    break
            except Exception:          # noqa: BLE001
                break
    except KeyboardInterrupt:
        print("interrupted — saving the profile")

    try:
        g.shutdown()
    except Exception:                  # noqa: BLE001
        pass
    print("session saved. Restart the pipeline now.")


if __name__ == "__main__":
    main()
