#!/usr/bin/env python3
"""Run the 7-role India leadership lookup across N parallel AI Mode browsers.

Each worker is a separate OS process with its OWN Chrome profile
(GOOGLE_AI_MODE_PROFILE_DIR), so the profile lock and cookie jar are not
shared. captcha-raptor is staged into every worker's extension directory,
so each browser can solve its own challenges.

AUTO-RESET: when a worker sees Google's "You've reached the request limit
for AI responses" page, it wipes its own profile and cookies, relaunches
and retries, instead of stalling until someone notices.

MEASURED CAVEAT, stated plainly because it shapes what to expect: Google's
AI-response limit is per-IP, not per-profile. All workers share one IP and
therefore one quota (~200-250 queries/day observed on this machine), so N
workers exhaust it roughly N times faster rather than multiplying
throughput. Earlier runs in this project also saw CAPTCHAs appear after
~25 queries at concurrency 10, versus a clean run at 3. Parallelism here
buys wall-clock speed until the shared quota is hit, not more total data.

Usage:
    python run_ceo_chro_parallel.py --companies-file list.txt --workers 10
    python run_ceo_chro_parallel.py --companies-file list.txt --workers 3
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

ROLES = ["CHRO", "CEO", "Founder", "Owner", "MD", "Co-founder", "Co-owner", "OTHER HEAD"]

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
OUT = ROOT / "output" / "retail_pharma"


def seed_extension(profile_base: Path, source_ext: Path) -> None:
    """Give this worker its own copy of the staged extension.

    google_ai_mode.py derives the staging root as
    "<profile_dir>_extensions" (a sibling of the profile), so each worker
    needs its own copy there for captcha-raptor to load.
    """
    staging = profile_base.parent / f"{profile_base.name}_extensions"
    target = staging / source_ext.name
    if target.is_dir():
        return
    staging.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source_ext, target, dirs_exist_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies-file", required=True)
    parser.add_argument("--workers", type=int, default=10)
    parser.add_argument("--batch-size", type=int, default=5)
    parser.add_argument("--retry-empty", action="store_true")
    parser.add_argument("--stagger", type=float, default=8.0,
                        help="Seconds between worker launches, so 10 browsers do not start at once")
    args = parser.parse_args()

    names = [l.strip() for l in Path(args.companies_file).read_text(encoding="utf-8").splitlines() if l.strip()]
    seen: set[str] = set()
    unique = []
    for n in names:
        if n.lower() in seen:
            continue
        seen.add(n.lower())
        unique.append(n)

    # Which companies still need work, so shards only cover real gaps.
    merged = OUT / "retail_pharma.json"
    done: set[str] = set()
    if merged.exists():
        data = json.loads(merged.read_text(encoding="utf-8"))
        for k, v in data.items():
            has = any((v.get(f) or {}).get("name") for f in ROLES)
            if has or not args.retry_empty:
                if has or k in data:
                    done.add(k)
        if args.retry_empty:
            done = {k for k, v in data.items()
                    if any((v.get(f) or {}).get("name") for f in ROLES)}

    todo = [c for c in unique if c.lower() not in done]
    print(f"companies total : {len(unique)}")
    print(f"already have    : {len(done)}")
    print(f"to process      : {len(todo)}")
    if not todo:
        print("nothing to do")
        return

    # Round-robin so every worker gets a mix, not one region of the list.
    shards: list[list[str]] = [[] for _ in range(args.workers)]
    for i, c in enumerate(todo):
        shards[i % args.workers].append(c)

    source_ext = DATA / "ai_mode_chrome_profile_extensions" / "captcha-raptor"
    if not source_ext.is_dir():
        print(f"WARNING: captcha-raptor not found at {source_ext}; workers will run without a solver")

    shard_dir = OUT / "shards"
    shard_dir.mkdir(parents=True, exist_ok=True)

    procs = []
    for w, shard in enumerate(shards):
        if not shard:
            continue
        shard_file = shard_dir / f"shard_{w}.txt"
        shard_file.write_text("\n".join(shard), encoding="utf-8")

        profile_base = DATA / f"ai_mode_rp{w}"
        if source_ext.is_dir():
            seed_extension(profile_base, source_ext)

        env = dict(os.environ)
        env["GOOGLE_AI_MODE_PROFILE_DIR"] = str(profile_base)
        env["GOOGLE_AI_MODE_ENABLED"] = "true"
        env["PYTHONIOENCODING"] = "utf-8"

        log = open(OUT / f"worker_{w}.log", "a", encoding="utf-8")
        cmd = [
            sys.executable, str(ROOT / "run_retail_pharma.py"),
            "--companies-file", str(shard_file),
            "--checkpoint", str(shard_dir / f"ckpt_{w}.json"),
            "--output", str(shard_dir / f"out_{w}.xlsx"),
            "--batch-size", str(args.batch_size),
            "--auto-reset",
        ]
        if args.retry_empty:
            cmd.append("--retry-empty")
        p = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=log, stderr=subprocess.STDOUT)
        procs.append((w, p, len(shard)))
        print(f"  worker {w}: {len(shard)} companies, profile {profile_base.name}", flush=True)
        time.sleep(args.stagger)

    print(f"\n{len(procs)} workers running. Merge with: python merge_retail_pharma_shards.py")
    for w, p, _ in procs:
        p.wait()
        print(f"worker {w} finished (exit {p.returncode})", flush=True)


if __name__ == "__main__":
    main()
