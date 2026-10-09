#!/usr/bin/env python3
"""Top each market up to N VERIFIED companies, not N discovered ones.

The discovery loop counts what it collects; verification then removes a
share of that. Measured after the third-party-distributor rule landed:

    Silicon Carbide  325 discovered -> 269 verified   (17% removed)
    CPaaS            377 discovered -> 132 verified   (65% removed)

So "--target 400" stops at 400 discovered and leaves ~330 verified, and
for CPaaS ~140. Reaching 400 VERIFIED means discovering past 400 and
re-verifying, repeatedly, until the verified count is the one that hits
the target.

Two things this adds beyond looping the two existing commands:

  1. REJECTED COMPANIES BECOME EXCLUSIONS. Discovery only excludes what
     is in state; it has never seen the verification verdicts, so it
     re-suggests companies already dropped as manufacturers and the same
     rejects cycle forever. Before each discovery pass the dropped names
     are written into the state's exclusion list, so the prompt says "do
     not repeat these" about them too.

  2. A HEADROOM ESTIMATE. It discovers to
         target / (verified rate so far)
     rather than to the target, so one pass usually suffices instead of
     five. The rate is measured per market from its own history, because
     it ranges from 1.17x (Silicon Carbide) to 2.9x (CPaaS).

Everything is resumable: discovery checkpoints per batch and verification
per chunk, so re-running continues rather than restarting.

Usage:
    python top_up_to_target.py --target 400
    python top_up_to_target.py --target 400 --only "silicon carbide"
    python top_up_to_target.py --target 400 --dry-run
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
)

OUT = ROOT / "output" / "market_research"

MARKETS: tuple[tuple[str, str], ...] = (
    ("Global Silicon Carbide Market - Distributors Only", "Global Silicon Carbide Market"),
    ("Global Wearable Glucometer Market - Distributors Only", "Global Wearable Glucometer Market"),
    ("Global Dry Eyes Treatment Market - Distributors Only", "Global Dry Eyes Treatment Market"),
    ("Global Solar Rooftop Market - Distributors Only", "Global Solar Rooftop Market"),
    ("Global Polyurethane Market - Distributors Only", "Global Polyurethane Market"),
    ("Global CPaaS (Communication Platform as a Service) Market - Distributors Only",
     "Global CPaaS (Communication Platform as a Service) Market"),
)

# Give up on a market after this many discover+verify rounds. A market
# that is genuinely exhausted stops yielding new companies long before
# this, and the low-yield stop inside run_market_research ends each
# discovery pass on its own.
_MAX_ROUNDS = 6


def _norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name or "").lower())


def _state_count(slug: str) -> int:
    path = STATE_DIR / f"{slug}.json"
    if not path.is_file():
        return 0
    return len(json.loads(path.read_text(encoding="utf-8")).get("companies") or [])


def _verdicts(slug: str) -> list[dict]:
    path = OUT / f"{slug}_verify_checkpoint.json"
    if not path.is_file():
        return []
    try:
        return json.loads(path.read_text(encoding="utf-8")).get("verdicts") or []
    except (json.JSONDecodeError, OSError):
        return []


def verified_count(slug: str) -> int:
    return sum(1 for v in _verdicts(slug) if v.get("keep"))


def rejected_names(slug: str) -> list[str]:
    """Companies verification judged NOT to be third-party distributors."""
    return [
        str(v.get("Company") or "")
        for v in _verdicts(slug)
        if not v.get("keep") and not v.get("backend_failure") and v.get("Company")
    ]


def seed_exclusions(market: str, slug: str) -> int:
    """Teach discovery about the rejects, so it stops re-suggesting them.

    state.excluded_names() is built from the collected companies, and a
    rejected company IS still in state -- it was collected, then dropped
    at verification. So the exclusion list already covers it, and the
    prompt already says "do not repeat". What it does NOT say is WHY, and
    a model re-offers a manufacturer it considers a good answer.

    Writing the rejects into a dedicated key lets the discovery prompt
    name them as rejected rather than merely seen. Kept as data on the
    state file so it survives restarts.
    """
    path = STATE_DIR / f"{slug}.json"
    if not path.is_file():
        return 0
    data = json.loads(path.read_text(encoding="utf-8"))
    rejects = rejected_names(slug)
    if not rejects:
        return 0
    known = {_norm(n) for n in data.get("rejected_names") or []}
    merged = list(data.get("rejected_names") or [])
    for name in rejects:
        if _norm(name) not in known:
            known.add(_norm(name))
            merged.append(name)
    data["rejected_names"] = merged
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(merged)


def discovery_target(slug: str, target: int) -> int:
    """How many to DISCOVER so that `target` survive verification.

    Uses the market's own measured pass rate, which ranges from 1.17x
    (Silicon Carbide) to 2.9x (CPaaS) -- a single global multiplier would
    either overshoot cheap markets or undershoot expensive ones.
    """
    judged = [v for v in _verdicts(slug) if not v.get("backend_failure")]
    kept = sum(1 for v in judged if v.get("keep"))
    if not judged or not kept:
        return target * 2          # no history yet: assume half survive
    rate = kept / len(judged)
    # +10% margin, and never ask for less than the target itself.
    return max(target, int(target / rate * 1.1) + 1)


def run(cmd: list[str]) -> int:
    print(f"    $ {' '.join(str(c) for c in cmd[1:3])} ...", flush=True)
    try:
        return subprocess.run(cmd, cwd=str(ROOT)).returncode
    except KeyboardInterrupt:
        raise
    except Exception as exc:  # noqa: BLE001
        print(f"    ! {type(exc).__name__}: {exc}")
        return -1


def top_up(market: str, market_name: str, *, target: int, dry_run: bool) -> dict:
    slug = market_slug(market)
    start_verified = verified_count(slug)

    print("")
    print("=" * 70)
    print(f"{market_name}")
    print(f"  state {_state_count(slug)} | verified {start_verified} | target {target}")
    print("=" * 70)

    if start_verified >= target:
        print("  already at target")
        return {"market": market_name, "before": start_verified,
                "after": start_verified, "rounds": 0}

    if dry_run:
        need = target - start_verified
        print(f"  would discover to {discovery_target(slug, target)} "
              f"to gain {need} verified")
        return {"market": market_name, "before": start_verified,
                "after": start_verified, "rounds": 0}

    for round_no in range(1, _MAX_ROUNDS + 1):
        have = verified_count(slug)
        if have >= target:
            break

        n_excluded = seed_exclusions(market, slug)
        disc_target = discovery_target(slug, target)
        print(f"  round {round_no}: verified {have}/{target} — "
              f"discovering to {disc_target} "
              f"({n_excluded} rejected names excluded)")

        run([sys.executable, str(ROOT / "run_market_research.py"),
             "--market", market, "--target", str(disc_target),
             "--batch-size", "20", "--export-xlsx"])

        run([sys.executable, str(ROOT / "verify_distributors_ai_mode.py"),
             "--file", str(OUT / f"{slug}.xlsx"),
             "--market-name", market_name])

        after = verified_count(slug)
        if after <= have:
            print(f"  round {round_no} added no verified companies — "
                  "the market is exhausted at this quality bar")
            break

    end = verified_count(slug)
    print(f"  --> {start_verified} -> {end} verified")
    return {"market": market_name, "before": start_verified,
            "after": end, "rounds": round_no}


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--target", type=int, default=400)
    ap.add_argument("--only", action="append", default=[], metavar="TEXT")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    queue = list(MARKETS)
    if args.only:
        needles = [n.lower() for n in args.only]
        queue = [m for m in queue if any(n in m[0].lower() for n in needles)]

    results = []
    try:
        for market, name in queue:
            results.append(top_up(market, name, target=args.target,
                                  dry_run=args.dry_run))
    except KeyboardInterrupt:
        print("\ninterrupted — progress is saved; re-run to continue")

    print("")
    print("=" * 70)
    print("SUMMARY")
    print("=" * 70)
    for r in results:
        gained = r["after"] - r["before"]
        print(f"  {r['market'][:48]:50} {r['before']:4} -> {r['after']:4} "
              f"(+{gained})")


if __name__ == "__main__":
    main()
