#!/usr/bin/env python3
"""Launch 10 discovery browsers across the 4 markets under 150, rotating
states automatically -- unlike the first version of this script, a
worker does not stop when its assigned state runs dry, it moves to the
next unswept state and keeps going.

Each of the 4 markets gets 2-3 parallel workers. Workers on the SAME
market write to the same discovery state file; state.py's save() merges
with whatever is on disk under a file lock instead of overwriting, so
this is safe for concurrent same-market workers (verified live with 10
concurrent processes, zero data loss, zero duplicates).

THE BUG THIS FIXES (measured live on the first version): each worker was
given ONE fixed state and simply stopped once that state's supply ran
out -- "3 consecutive empty batches ... treating as genuinely exhausted"
after well under a minute of real work, because a single state has far
fewer companies than the whole country. 7 of 10 workers exited within
minutes and only 6 new companies were found total. supervisor_all.py's
own normal single-worker-per-market runs never hit this, because it
rotates through all 26 states (then individual cities) automatically
when one runs dry -- this script now does the same thing per worker,
with workers on the same market coordinating so two of them never pick
the same state at once.

--target passed to run_drone2_market.py --target is a DISCOVERED-count
target, not verified (see run_until_target's actual loop condition), so
each worker's per-region run is asked for a modest, achievable chunk
above whatever the market currently has -- not the whole remaining
verified gap in one go -- and is recomputed on every rotation so it
tracks progress from sibling workers too.

Usage:
    python run_10_browsers.py
    python run_10_browsers.py --target 160
"""
from __future__ import annotations

import argparse
import re
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

from india_states import INDIA_STATES  # noqa: E402
from india_cities import INDIA_CITIES  # noqa: E402
import supervisor_all as S  # noqa: E402

LOGS = ROOT / "logs_drones"
LOGS.mkdir(exist_ok=True)

PLAN: tuple[tuple[str, int], ...] = (
    ("drone_testing", 3),
    ("drone_inspection", 3),
    ("drone_surveillance", 2),
    ("drone_training", 2),
)

BATCH_SIZE = 20
MAX_ROTATIONS = 40


def stage_extension(slug: str, suffix: str) -> None:
    base = ROOT / "data" / f"drone2_profile_{slug}_extensions"
    dst = ROOT / "data" / f"drone2_profile_{slug}_{suffix}_extensions"
    if base.is_dir() and not dst.is_dir():
        shutil.copytree(base, dst)


def covered_regions(slug: str, worker_suffixes: list[str]) -> set[str]:
    """Every state/city ANY worker on this market has already queried,
    pooled across all sibling workers so two of them never pick the
    same region at once."""
    used: set[str] = set()
    for suffix in worker_suffixes:
        log = LOGS / f"{slug}_{suffix}.log"
        if not log.is_file():
            continue
        try:
            text = log.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        used |= {m.split("(")[0].strip()
                 for m in re.findall(r"coverage: ([^\n]+)", text)}
    return used


def next_region(slug: str, worker_suffixes: list[str]) -> str | None:
    used = covered_regions(slug, worker_suffixes)
    for state in INDIA_STATES:
        if state.split("(")[0].strip() not in used:
            return state
    for city in INDIA_CITIES:
        if city not in used:
            return city
    return None


def run_one_region(slug: str, suffix: str, region: str,
                   goal_padding: int) -> int:
    market = dict((s, m) for s, m, _s, _p in S.GROUPS)[slug]
    before = S.discovered(market)
    goal = before + goal_padding

    log_path = LOGS / f"{slug}_{suffix}.log"
    handle = log_path.open("a", encoding="utf-8", errors="replace")
    cmd = [
        sys.executable, "-u", str(ROOT / "run_drone2_market.py"),
        "--slug", slug,
        "--target", str(goal),
        "--batch-size", str(BATCH_SIZE),
        "--region", region,
        "--profile-suffix", suffix,
    ]
    subprocess.run(cmd, cwd=str(ROOT), stdout=handle,
                   stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                   timeout=1800)
    handle.close()
    after = S.discovered(market)
    return after - before


def worker_loop(slug: str, suffix: str, worker_suffixes: list[str],
                verify_target: int, results: dict) -> None:
    stage_extension(slug, suffix)
    market = dict((s, m) for s, m, _s, _p in S.GROUPS)[slug]

    for rotation in range(1, MAX_ROTATIONS + 1):
        if S.verified(market) >= verify_target:
            results[(slug, suffix)] = f"market at target ({S.verified(market)})"
            return

        region = next_region(slug, worker_suffixes)
        if region is None:
            results[(slug, suffix)] = "no unswept region left"
            return

        gained = 0
        try:
            gained = run_one_region(slug, suffix, region, goal_padding=60)
        except subprocess.TimeoutExpired:
            pass
        except Exception as exc:  # noqa: BLE001
            log_path = LOGS / f"{slug}_{suffix}.log"
            with log_path.open("a", encoding="utf-8", errors="replace") as fh:
                fh.write(f"\n[launcher] worker error: {exc}\n")

        results[(slug, suffix)] = (
            f"rotation {rotation}: {region.split(chr(40))[0].strip()} +{gained}"
        )
        time.sleep(2)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=160)
    args = ap.parse_args()

    results: dict[tuple[str, str], str] = {}
    threads = []
    for slug, n_workers in PLAN:
        worker_suffixes = [f"w{w}" for w in range(1, n_workers + 1)]
        for suffix in worker_suffixes:
            t = threading.Thread(
                target=worker_loop,
                args=(slug, suffix, worker_suffixes, args.target, results),
                daemon=True,
            )
            threads.append((slug, suffix, t))
            t.start()
            time.sleep(2)

    print(f"launched {len(threads)} rotating workers across {len(PLAN)} "
          f"markets, verify-target={args.target}")
    print("each worker rotates to a new state/city when its current one "
          "runs dry -- check progress with:")
    print("  python verify_all_drones.py --status")
    print("or watch a specific worker's rotations:")
    print("  tail -f logs_drones/<slug>_<w1|w2|w3>.log")

    for slug, suffix, t in threads:
        t.join()

    print("")
    print("all workers finished (target reached or every region swept):")
    for (slug, suffix), status in sorted(results.items()):
        print(f"  {slug:22} {suffix}  {status}")


if __name__ == "__main__":
    main()
