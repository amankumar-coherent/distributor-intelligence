#!/usr/bin/env python3
"""Keep all 11 markets working toward 400, unattended.

Written to run overnight with nobody watching. It handles the three
failure modes seen repeatedly in this project, in the order they cost
the most:

  1. A MARKET STOPS. Discovery exits when it decides the market is
     exhausted. That verdict has been wrong every time it mattered:
     Polyurethane stopped at 309 and reached 401; Anti-Drone stopped at
     86 and produced seven more from one Latin America query. So a
     stopped market is always restarted.

  2. A MARKET STALLS. It runs but stops accepting companies, because the
     exclusion list has saturated and a broad query has no new names to
     give. The fix is to ask a NARROWER question -- city by city for the
     India markets, country by country for Dangjo. Each stall advances
     that market one step through its list, so it works through the
     whole country rather than repeating one query.

  3. THE BACKEND REFUSES. "AI cannot generate a response", a rate limit,
     or a CAPTCHA wall. These live in the session cookies, so a browser
     restart alone never clears them -- only a profile reset does. The
     per-market processes already self-reset, but a market that is stuck
     in a reset loop gets its profile rebuilt from here and is restarted
     clean.

Every layer underneath checkpoints, so restarting costs nothing: a
market resumes from its last completed batch.

Usage:
    python overnight_supervisor.py
    python overnight_supervisor.py --interval 180 --status-every 10
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from drone_markets import MARKETS  # noqa: E402
import run_dangjo  # noqa: E402

TARGET = 400
LOGS = ROOT / "logs_drones"
DATA = ROOT / "data"

# India's drone industry by concentration; used when a drone market
# saturates on the plain "India" question.
CITIES = (
    "Bengaluru, Karnataka", "Hyderabad, Telangana",
    "Delhi NCR (Delhi, Noida, Gurugram)", "Pune, Maharashtra",
    "Chennai, Tamil Nadu", "Mumbai, Maharashtra",
    "Ahmedabad and Gujarat", "Coimbatore, Tamil Nadu",
    "Kolkata and West Bengal", "Jaipur and Rajasthan",
    "Kochi and Kerala", "Indore and Madhya Pradesh",
    "Chandigarh, Punjab and Haryana", "Lucknow and Uttar Pradesh",
    "Bhubaneswar and Odisha", "Nagpur and Vidarbha",
    "Visakhapatnam and Andhra Pradesh", "Surat and south Gujarat",
    "Thiruvananthapuram, Kerala", "Guwahati and north-east India",
)

# Text that means the session itself is poisoned, not the market empty.
REFUSAL_MARKERS = (
    "ai cannot generate",
    "can't generate a response",
    "couldn't generate a response",
    "no response available for this search",
    "request limit for ai responses",
    "try again in a little while",
    "unusual traffic",
)


def _ps_count(like: str) -> int:
    ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | "
        f"Where-Object {{ $_.CommandLine -like '{like}' }} | "
        "Measure-Object | Select-Object -ExpandProperty Count"
    )
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True, timeout=45).stdout
        return int(out.strip() or 0)
    except (ValueError, subprocess.SubprocessError):
        return 0


def discovered(market: str) -> int:
    from vendor_intel.pipeline.market_research.state import (
        STATE_DIR, market_slug,
    )
    path = STATE_DIR / f"{market_slug(market)}.json"
    if not path.is_file():
        return 0
    try:
        return len(json.loads(path.read_text(encoding="utf-8"))
                   .get("companies") or [])
    except (json.JSONDecodeError, OSError):
        return 0


def tail(path: Path, n: int = 4000) -> str:
    if not path.is_file():
        return ""
    try:
        data = path.read_text(encoding="utf-8", errors="replace")
        return data[-n:].lower()
    except OSError:
        return ""


def session_is_poisoned(log: Path) -> bool:
    """True when the log tail shows a refusal/limit rather than progress."""
    text = tail(log)
    if not text:
        return False
    if any(m in text for m in REFUSAL_MARKERS):
        # Only if it has not recovered since: a marker followed by an
        # accept means the reset already worked.
        last_marker = max(text.rfind(m) for m in REFUSAL_MARKERS
                          if m in text)
        return text.rfind("accepted #") < last_marker
    return False


def reset_profile(profile_base: Path) -> None:
    """Rebuild a profile, preserving its staged extensions.

    The staged extensions live in a SIBLING directory, so deleting the
    profile leaves captcha-raptor intact and the next launch reinstalls
    it automatically.
    """
    for suffix in ("_chromium", ""):
        p = Path(str(profile_base) + suffix)
        if p.is_dir():
            try:
                shutil.rmtree(p, ignore_errors=True)
            except OSError:
                pass


def launch_drone(slug: str, region: str | None) -> None:
    cmd = [sys.executable, "-u", str(ROOT / "run_drone_market.py"),
           "--slug", slug, "--target", str(TARGET), "--batch-size", "20"]
    if region:
        cmd += ["--region", region]
    env = dict(os.environ)
    env["GOOGLE_AI_MODE_QUERY_DELAY"] = "12"
    env["MARKET_RESEARCH_BLOCK_COOLOFF"] = "60"
    env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "120"
    LOGS.mkdir(exist_ok=True)
    handle = (LOGS / f"{slug}.log").open("a", encoding="utf-8",
                                         errors="replace")
    subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=handle,
                     stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL)


def launch_dangjo() -> None:
    cmd = [sys.executable, "-u", str(ROOT / "run_dangjo_by_country.py"),
           "--passes", "2", "--batch-size", "10"]
    env = dict(os.environ)
    env["GOOGLE_AI_MODE_QUERY_DELAY"] = "12"
    env["MARKET_RESEARCH_BLOCK_COOLOFF"] = "60"
    env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "120"
    handle = (ROOT / "logs_dangjo_overnight.txt").open(
        "a", encoding="utf-8", errors="replace")
    subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=handle,
                     stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--interval", type=int, default=240)
    ap.add_argument("--status-every", type=int, default=5,
                    help="print a full table every N cycles")
    args = ap.parse_args()

    city_idx: dict[str, int] = {}
    last_count: dict[str, int] = {}
    stalls: dict[str, int] = {}
    cycle = 0

    print(f"supervisor started {datetime.now():%Y-%m-%d %H:%M:%S}", flush=True)
    print(f"target {TARGET} per market, checking every {args.interval}s",
          flush=True)

    while True:
        cycle += 1
        now = datetime.now().strftime("%H:%M:%S")

        for slug, market, _ in MARKETS:
            have = discovered(market)
            if have >= TARGET:
                continue

            alive = _ps_count(f"*slug {slug}*") > 0
            log = LOGS / f"{slug}.log"
            poisoned = session_is_poisoned(log)

            # 3. Poisoned session: rebuild the profile before relaunching,
            #    because the refusal lives in the cookies.
            if poisoned:
                print(f"[{now}] {slug}: refusal/limit in log -> profile "
                      f"reset + restart (at {have})", flush=True)
                subprocess.run(
                    ["powershell", "-NoProfile", "-Command",
                     "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
                     f"| Where-Object {{ $_.CommandLine -like '*slug {slug}*' }} "
                     "| ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
                     "-ErrorAction SilentlyContinue }"],
                    capture_output=True, timeout=60)
                time.sleep(3)
                reset_profile(DATA / f"drone_profile_{slug}")
                idx = city_idx.get(slug, 0)
                launch_drone(slug, CITIES[idx % len(CITIES)])
                city_idx[slug] = idx + 1
                last_count[slug] = have
                time.sleep(4)
                continue

            # 1. Stopped: always restart, because "exhausted" has been
            #    wrong every time it mattered.
            if not alive:
                idx = city_idx.get(slug, -1)
                if idx < 0:
                    region, why = None, "restart"
                    city_idx[slug] = 0
                else:
                    region = CITIES[idx % len(CITIES)]
                    why = f"city: {region}"
                    city_idx[slug] = idx + 1
                print(f"[{now}] {slug}: stopped at {have} -> {why}",
                      flush=True)
                launch_drone(slug, region)
                last_count[slug] = have
                time.sleep(4)
                continue

            # 2. Running but not gaining: narrow the question.
            if last_count.get(slug) == have:
                stalls[slug] = stalls.get(slug, 0) + 1
                if stalls[slug] >= 3:      # ~12 minutes of no progress
                    idx = city_idx.get(slug, 0)
                    city = CITIES[idx % len(CITIES)]
                    print(f"[{now}] {slug}: stalled at {have} -> city: {city}",
                          flush=True)
                    subprocess.run(
                        ["powershell", "-NoProfile", "-Command",
                         "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
                         f"| Where-Object {{ $_.CommandLine -like '*slug {slug}*' }} "
                         "| ForEach-Object { Stop-Process -Id $_.ProcessId "
                         "-Force -ErrorAction SilentlyContinue }"],
                        capture_output=True, timeout=60)
                    time.sleep(3)
                    launch_drone(slug, city)
                    city_idx[slug] = idx + 1
                    stalls[slug] = 0
                    time.sleep(4)
            else:
                stalls[slug] = 0
            last_count[slug] = have

        # Dangjo: restart whenever it is not running. Its own sweep
        # already rotates countries, so no extra steering is needed.
        dj_have = discovered(run_dangjo.MARKET)
        if dj_have < TARGET and _ps_count("*run_dangjo*") == 0:
            print(f"[{now}] dangjo: stopped at {dj_have} -> restart",
                  flush=True)
            launch_dangjo()

        if cycle % args.status_every == 0:
            total = 0
            print(f"\n--- {now} ---", flush=True)
            for slug, market, _ in MARKETS:
                d = discovered(market)
                total += d
                print(f"  {slug:22}{d:>6}/{TARGET}", flush=True)
            print(f"  {'dangjo':22}{dj_have:>6}/{TARGET}", flush=True)
            print(f"  {'TOTAL':22}{total + dj_have:>6}\n", flush=True)

        time.sleep(args.interval)


if __name__ == "__main__":
    main()
