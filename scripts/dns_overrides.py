#!/usr/bin/env python3
"""Move rows whose website is dead (DNS does not resolve after 3 slow retries)
or is only a directory / social page out of the main list, through the status
check's overrides.json (flagged_for_review.xlsx gets them), before --build.

Why: fast parallel DNS gives false "dead" answers, so failures are re-checked
slowly with retries before anything is removed (esu / moz / jpf / iww lessons).
Keys are company_key(name, website) of the STAGED verified workbooks.

Usage:
  python scripts/dns_overrides.py --stage output/final/adh/adh_verified
Then rebuild:
  python scripts/status_check_folder.py --input <stage> --output <out> --build
"""
from __future__ import annotations

import argparse
import json
import os
import re
import socket
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

import openpyxl  # noqa: E402

from scripts.status_check_folder import _domain, company_key, run_dir  # noqa: E402

DIRECTORY = re.compile(r"indiamart|tradeindia|justdial|exportersindia|linkedin|facebook|"
                       r"alibaba|yellowpages|instagram|twitter|youtube|wikipedia|rocketreach|zoominfo|"
                       r"crunchbase|tofler|zaubacorp|sulekha|glassdoor|ambitionbox")


def resolves(host: str) -> bool:
    try:
        socket.getaddrinfo(host, None)
        return True
    except Exception:  # noqa: BLE001 - OSError, and UnicodeError ("label empty or too long") for malformed domains
        return False


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--stage", required=True, type=Path)
    args = ap.parse_args()
    stage = args.stage if args.stage.is_absolute() else ROOT / args.stage

    rows = []  # (company, website)
    for f in sorted(stage.glob("*_verified_distributors.xlsx")):
        ws = openpyxl.load_workbook(f, read_only=True).active
        header = None
        for i, r in enumerate(ws.iter_rows(values_only=True), 1):
            if header is None:
                if r and r[0] == "#":
                    header = list(r)
                continue
            if not any(r):
                continue
            rec = dict(zip(header, r))
            if rec.get("Company"):
                rows.append((rec["Company"], rec.get("Website")))
    doms = {}
    for n, w in rows:
        doms.setdefault(_domain(w), []).append(n)
    doms.pop("", None)
    print(f"{len(rows)} rows, {len(doms)} unique domains", flush=True)

    with ThreadPoolExecutor(40) as ex:
        res = dict(zip(doms, ex.map(resolves, list(doms))))
    fast_bad = [d for d, ok in res.items() if not ok]
    print(f"fast failures: {len(fast_bad)}", flush=True)
    def slow_check(d):
        for _ in range(3):
            time.sleep(1.5)
            if resolves(d) or resolves("www." + d):
                return True
        return False

    with ThreadPoolExecutor(40) as ex:
        alive = list(ex.map(slow_check, fast_bad))
    dead = [d for d, ok in zip(fast_bad, alive) if not ok]
    directory = [d for d in doms if DIRECTORY.search(d)]
    print(f"dead after 3 slow retries: {len(dead)} | directory/social sites: {len(directory)}", flush=True)

    path = run_dir(stage) / "overrides.json"
    cur = json.loads(path.read_text(encoding="utf-8-sig")) if path.exists() else {}
    added = 0
    for n, w in rows:
        d = _domain(w)
        key = company_key(n, w)
        if key in cur:
            continue
        if d in dead:
            cur[key] = {"note": "website does not resolve (checked 3x with retries) - company cannot be confirmed"}
            added += 1
        elif d in directory:
            cur[key] = {"note": "website is a directory / social page, not the company site"}
            added += 1
    os.makedirs(path.parent, exist_ok=True)
    path.write_text(json.dumps(cur, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"overrides added: {added} (total {len(cur)}) -> {path}", flush=True)


if __name__ == "__main__":
    main()
