#!/usr/bin/env python3
"""Unattended watchdog for a rolling batch run (run_fire_10_browsers.py --rolling).

Every --every seconds:
  * appends per-market found / verified counts to logs/<batch>_progress.log;
  * if the launcher process is gone and markets are still short of target,
    relaunches it through its .ps1 (detached, so it survives this process);
  * if a market's discovery workers are alive but its log has been silent
    for --stale minutes, the browsers are stuck (a block page that never
    clears, a hung tab): kills those workers and archives the market's
    profiles, so the launcher's next round starts on fresh cookies.
When every market is finished it runs the last pipeline step itself
(status + ownership check, packaging) unless the launcher already did,
then exits.

Usage:
    python scripts/batch_watchdog.py --batch wound --ps1 logs\\run_wound_batch.ps1 --target 400
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

import run_fire_10_browsers as L  # noqa: E402
import supervisor_all as S  # noqa: E402
import market_batches  # noqa: E402


def _ps(command: str) -> str:
    try:
        return subprocess.run(["powershell", "-NoProfile", "-Command", command],
                              capture_output=True, text=True, timeout=90).stdout.strip()
    except subprocess.SubprocessError:
        return ""


def _count(pattern: str) -> int:
    out = _ps("(Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | "
              f"Where-Object {{ $_.CommandLine -match '{pattern}' }} | Measure-Object).Count")
    try:
        return int(out or 0)
    except ValueError:
        return -1


def _start_ps1(path: Path) -> None:
    """Start a .ps1 detached (WMI), so it outlives this watchdog. Under a
    headless conhost: Windows 11 otherwise opens a Windows Terminal window
    per launch (it ignores -WindowStyle Hidden), and closing it kills the run."""
    _ps("Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ "
        f"CommandLine = 'conhost.exe --headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"{path}\"'; "
        f"CurrentDirectory = '{ROOT}' }} | Out-Null")


def _launcher_log_from_ps1(ps1: Path) -> Path:
    """The log file the .ps1 actually appends to, read from its own
    `Out-File ... "logs\\...log"` line.

    Guessing `run_{batch}_batch.log` from the --batch key broke silently:
    run_lacto_batch.ps1 (--batch ingr) writes to run_lacto_batch.log, not
    run_ingr_batch.log, so the watchdog's "is it finished / finalized"
    checks always read an empty file. It never saw the real FINALIZE/"markets
    processed" lines, so it kept treating a genuinely finished, fully
    status-checked market as still running and relaunched it every --every
    seconds for hours (harmless here — a finished run's own finalize step
    is a fast no-op re-check with no new queries — but still wrong and the
    watchdog never exited on its own)."""
    text = ps1.read_text(encoding="utf-8", errors="replace")
    m = re.search(r'Out-File[^\n]*?"((?:logs\\|logs/)[^"]+\.log)"', text)
    if not m:
        raise RuntimeError(f"{ps1}: no Out-File log path found — cannot track completion")
    return (ROOT / m.group(1)).resolve()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--batch", required=True)
    ap.add_argument("--ps1", required=True, help="launcher .ps1 to relaunch with")
    ap.add_argument("--target", type=int, required=True)
    ap.add_argument("--every", type=int, default=600, help="seconds between checks")
    ap.add_argument("--stale", type=int, default=25, help="minutes of log silence = stuck")
    ap.add_argument("--status-workers", type=int, default=10,
                    help="browsers for the final status + ownership check")
    ap.add_argument("--then", action="append", default=[], metavar="PS1",
                    help="start this .ps1 once the batch is fully finished "
                    "and checked (repeatable, e.g. the next batch + its watchdog)")
    ap.add_argument("--start-delay", type=int, default=180,
                    help="seconds before the first check, so a launcher started "
                    "together with this watchdog is up before it is looked for")
    args = ap.parse_args()
    time.sleep(args.start_delay)

    L.select_batch(args.batch)
    slugs = list(L.FIRE_SLUGS)
    runner = Path(L.RUNNER).stem
    progress = ROOT / "logs" / f"{args.batch}_progress.log"
    ps1 = (ROOT / args.ps1).resolve()
    launcher_log = _launcher_log_from_ps1(ps1)

    def note(msg: str) -> None:
        with progress.open("a", encoding="utf-8") as fh:
            fh.write(f"[{datetime.now():%Y-%m-%d %H:%M}] {msg}\n")

    note(f"watchdog started: {len(slugs)} markets, target {args.target}")
    while True:
        rows, done = [], 0
        for slug in slugs:
            market = market_batches.by_slug(slug)[1]
            d, v = S.discovered(market), S.verified(market)
            done += v >= args.target
            rows.append(f"{slug} {v}/{args.target} verified ({d} found)")
        note(" | ".join(rows))

        finished = launcher_log.is_file() and f"markets processed" in launcher_log.read_text(
            encoding="utf-8", errors="replace")[-4000:]
        if done == len(slugs) or finished:
            # Last pipeline step: status + ownership check and packaging. A
            # launcher started before finalize() existed never runs it, so
            # the watchdog does unless the launcher already did / is doing it.
            tail = launcher_log.read_text(encoding="utf-8", errors="replace")[-20000:] \
                if launcher_log.is_file() else ""
            if f"FINALIZE {args.batch}: done" not in tail:
                if _count("run_fire_10_browsers.*--batch") > 0 and "FINALIZE" in tail:
                    note("launcher is running the final check — waiting")
                    time.sleep(args.every)
                    continue
                if _count("status_check_folder") <= 0:
                    note("all markets finished — running final status + ownership check")
                    subprocess.run([sys.executable, "-u", str(ROOT / "run_fire_10_browsers.py"),
                                    "--batch", args.batch, "--finalize-only",
                                    "--status-workers", str(args.status_workers)],
                                   cwd=str(ROOT), stdout=launcher_log.open("a", encoding="utf-8"),
                                   stderr=subprocess.STDOUT)
                    # Finished -> the next pass sees "done" and exits.
                    # Incomplete (Google blocks) -> pause, then resume it.
                    tail = launcher_log.read_text(encoding="utf-8", errors="replace")[-20000:]
                    if f"FINALIZE {args.batch}: done" not in tail:
                        time.sleep(args.every)
                    continue
            for nxt in args.then:
                note(f"starting next: {nxt}")
                _start_ps1((ROOT / nxt).resolve())
                time.sleep(20)
            note("all markets finished and checked — watchdog exiting")
            return

        if _count(f"run_fire_10_browsers.*--batch {args.batch}") == 0:
            note("launcher not running — relaunching")
            _start_ps1(ps1)
        else:
            for slug in slugs:
                log = L.LOGS / f"{slug}.log"
                if not log.is_file():
                    continue
                silent = (time.time() - log.stat().st_mtime) / 60
                if silent < args.stale or _count(f"{runner}.*--slug {slug}\\b") <= 0:
                    continue
                note(f"{slug}: workers silent {silent:.0f} min — killing them and resetting profiles")
                _ps("Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { "
                    f"$_.CommandLine -match '{runner}.*--slug {slug}\\b' }} | "
                    "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }")
                L.reset_profile(slug, f"watchdog: log silent {silent:.0f} min")
        time.sleep(args.every)


if __name__ == "__main__":
    main()
