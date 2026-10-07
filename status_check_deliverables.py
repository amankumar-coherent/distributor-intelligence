#!/usr/bin/env python3
"""Add active/closed status and acquired-by annotation to the 50 finished
deliverable workbooks, WITHOUT touching the production pipeline.

This is a one-off pass over files that were already verified and exported:
  C:/Users/aman.kumar/Desktop/Drone Market   (25 files)
  C:/Users/aman.kumar/Desktop/Fire Market    (25 files)

It never edits those files. For each input it writes a sibling copy named
<name>_status_checked.xlsx with two changes:

  * a company confirmed CLOSED is dropped from the copy;
  * a company confirmed ACQUIRED is renamed "ABC (acquired by XYZ)".

Everything else -- the banner rows, the header, every other column and
value -- is copied through untouched.

WHY THIS RESUMES RATHER THAN RUNS TO COMPLETION
-----------------------------------------------
~13,200 companies at 12 per query is ~1,100 Google AI Mode queries, and
AI Mode stops answering after roughly 200-250 queries from one IP in a
day (measured; see the project memory note). So a single run CANNOT
finish this. The script therefore:

  * checkpoints every chunk to data/status_check/<file>.json, so a day's
    work is never redone;
  * detects the daily wall and stops cleanly instead of burning the
    remaining chunks on empty answers.

Re-run it the next day and it picks up exactly where it stopped.

Usage:
    python status_check_deliverables.py               # run until the wall
    python status_check_deliverables.py --status      # progress only, no queries
    python status_check_deliverables.py --limit 50    # cap queries this run
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
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

import openpyxl  # noqa: E402

FOLDERS = [
    Path(r"C:\Users\aman.kumar\Desktop\Drone Market"),
    Path(r"C:\Users\aman.kumar\Desktop\Fire Market"),
]

CKPT_DIR = ROOT / "data" / "status_check"
CKPT_DIR.mkdir(parents=True, exist_ok=True)

HEADER_ROW = 5          # row 1-4 are banner text, row 5 is the header
FIRST_DATA_ROW = 6
COMPANY_COL = 2         # column B
CHUNK = 12              # same chunk size the verifier uses

# Stop the whole run after this many consecutive empty answers. AI Mode
# returns "" for a one-off failure AND for the daily per-IP wall, so the
# two are told apart by persistence: a single bad chunk recovers, the
# wall does not. Three in a row has always meant the wall in this project.
MAX_CONSECUTIVE_EMPTY = 3

_PROMPT = (
    'For each company below, answer two things.\n'
    'status: "closed" ONLY if the company is confirmed shut down, defunct, '
    'liquidated or ceased trading; otherwise "active".\n'
    'acquired_by: if it was ACQUIRED and no longer trades independently, '
    'the acquirer\'s name; else "".\n'
    'Never skip a name. One result per company.\n'
    'Reply with ONLY this JSON, no prose and no markdown:\n'
    '{"results":[{"Company":"","status":"active","acquired_by":""}]}'
)

# Values a model returns to mean "no acquirer". Writing any of these into
# the Company column would produce "ABC (acquired by None)".
_NO_ACQUIRER = {
    "", "none", "n/a", "na", "null", "nil", "unknown", "not acquired",
    "no", "false", "not applicable", "independent", "-", "--",
}


def log(msg: str) -> None:
    print(msg, flush=True)


def internet_up(timeout: float = 8.0) -> bool:
    """True if the network is reachable right now.

    Used to tell a network outage apart from the daily AI Mode limit:
    both produce an empty answer, but one should pause and retry while
    the other should stop for the day. A DNS-level check is the right
    probe -- the failure seen live was ERR_NAME_NOT_RESOLVED, i.e. DNS
    itself failing, not Google refusing.

    Fails SAFE: if the probe itself errors we return True, so a flaky
    probe can never masquerade as "network down" and keep the run
    spinning in a wait loop forever.
    """
    import socket
    try:
        socket.setdefaulttimeout(timeout)
        socket.getaddrinfo("www.google.com", 443)
        return True
    except OSError:
        return False
    except Exception:  # noqa: BLE001 - probe must never decide by crashing
        return True


def _norm(name: str) -> str:
    return "".join(ch for ch in str(name).lower() if ch.isalnum())


def clean_acquirer(value: object, company: object) -> str:
    """The acquirer's name, or "" when there isn't a usable one.

    Same guards as the pipeline's own helper: placeholder-for-empty, and
    the company's own name echoed back (which would render
    "ABC (acquired by ABC)").
    """
    name = str(value or "").strip().strip('"\'')
    if name.lower() in _NO_ACQUIRER:
        return ""
    if _norm(name) and _norm(name) == _norm(str(company or "")):
        return ""
    return name


def extract_json(text: str) -> dict | None:
    """First complete JSON object in the answer, by brace matching.

    AI Mode wraps the JSON in conversation text, so json.loads() on the
    whole answer fails; scanning for a balanced object is what the
    verifier does too.
    """
    if not text:
        return None
    best = None
    depth = 0
    start = None
    for i, ch in enumerate(text):
        if ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            if depth:
                depth -= 1
                if depth == 0 and start is not None:
                    try:
                        parsed = json.loads(text[start:i + 1])
                        if isinstance(parsed, dict) and "results" in parsed:
                            best = parsed
                    except json.JSONDecodeError:
                        pass
                    start = None
    return best


def ckpt_path(xlsx: Path) -> Path:
    return CKPT_DIR / f"{xlsx.stem}.json"


def load_ckpt(xlsx: Path) -> dict:
    p = ckpt_path(xlsx)
    if p.is_file():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            pass
    return {"verdicts": {}}


def save_ckpt(xlsx: Path, data: dict) -> None:
    tmp = ckpt_path(xlsx).with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1),
                   encoding="utf-8")
    os.replace(tmp, ckpt_path(xlsx))


def read_companies(xlsx: Path) -> list[str]:
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    ws = wb.active
    names = []
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW, values_only=True):
        if len(row) >= COMPANY_COL:
            v = row[COMPANY_COL - 1]
            if v and str(v).strip():
                names.append(str(v).strip())
    wb.close()
    return names


def write_copy(xlsx: Path, verdicts: dict) -> tuple[int, int, int]:
    """Write <name>_status_checked.xlsx. Returns (kept, dropped, renamed).

    Loads the original and edits the COPY in memory, so the user's file
    is never modified -- the whole point of writing alongside.
    """
    wb = openpyxl.load_workbook(xlsx)
    ws = wb.active

    drop_rows = []
    renamed = 0
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        cell = ws.cell(row=r, column=COMPANY_COL)
        name = str(cell.value or "").strip()
        if not name:
            continue
        v = verdicts.get(_norm(name))
        if not v:
            continue  # not yet judged -- leave exactly as-is
        if str(v.get("status") or "").strip().lower() == "closed":
            drop_rows.append(r)
            continue
        acq = v.get("acquired_by") or ""
        if acq and "acquired by" not in name.lower():
            cell.value = f"{name} (acquired by {acq})"
            renamed += 1

    # Delete bottom-up so earlier deletions don't shift later row indexes.
    for r in sorted(drop_rows, reverse=True):
        ws.delete_rows(r, 1)

    out = xlsx.with_name(f"{xlsx.stem}_status_checked.xlsx")
    wb.save(out)
    wb.close()
    kept = sum(1 for r in range(FIRST_DATA_ROW, ws.max_row + 1)
               if str(ws.cell(row=r, column=COMPANY_COL).value or "").strip())
    return kept, len(drop_rows), renamed


def all_files(shard: int = 0, shards: int = 1) -> list[Path]:
    """Workbooks to process, optionally sharded across parallel workers.

    Sharding is BY FILE, not by company, and that is the whole point:
    each workbook has one checkpoint JSON, and two processes writing the
    same checkpoint would overwrite each other's verdicts. Giving every
    worker a disjoint set of files means no checkpoint is ever shared,
    so no locking is needed and a worker dying can't corrupt another's
    progress.

    Round-robin (i % shards) rather than contiguous blocks, so the big
    and small workbooks spread evenly instead of one worker drawing all
    the 400-row files.
    """
    files: list[Path] = []
    for folder in FOLDERS:
        if not folder.is_dir():
            log(f"!! missing folder: {folder}")
            continue
        for f in sorted(folder.glob("*.xlsx")):
            if f.stem.endswith("_status_checked"):
                continue  # our own output
            files.append(f)
    if shards > 1:
        files = [f for i, f in enumerate(files) if i % shards == shard]
    return files


def show_status() -> None:
    """Progress, counted against DISTINCT companies.

    Counting raw rows overstated the work left: verdicts are keyed by
    normalised name, so two rows that normalise alike ("ABC Ltd" and
    "ABC Ltd.") share one verdict. Comparing len(verdicts) to the row
    count therefore reported a shortfall that no amount of querying
    could ever close -- measured at 113 phantom companies across the 50
    files, exactly the number of duplicate-name collisions in them.
    What actually matters is whether any company lacks a verdict, so
    that is what this counts.
    """
    tot_c = tot_j = 0
    for f in all_files():
        names = read_companies(f)
        verdicts = load_ckpt(f).get("verdicts", {})
        distinct = {_norm(n) for n in names}
        judged = len(distinct & set(verdicts))
        tot_c += len(distinct)
        tot_j += judged
        flag = "DONE" if judged >= len(distinct) else ""
        log(f"  {f.parent.name[:12]:12} {f.stem[:44]:44} "
            f"{judged:>5}/{len(distinct):<5} {flag}")
    log("")
    remaining = max(0, tot_c - tot_j)
    log(f"  judged {tot_j} of {tot_c} distinct companies "
        f"({remaining} left, ~{(remaining + CHUNK - 1)//CHUNK} queries)")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--status", action="store_true",
                    help="show progress and exit, no queries")
    ap.add_argument("--limit", type=int, default=0,
                    help="max queries this run (0 = until the daily wall)")
    ap.add_argument("--shard", type=int, default=0,
                    help="this worker's index, 0-based (parallel runs)")
    ap.add_argument("--shards", type=int, default=1,
                    help="total number of parallel workers")
    args = ap.parse_args()

    if args.status:
        show_status()
        return

    if not 0 <= args.shard < args.shards:
        raise SystemExit(f"--shard must be 0..{args.shards - 1}")

    from vendor_intel.scraping.google_ai_mode import ask

    queries = 0
    consecutive_empty = 0

    my_files = all_files(args.shard, args.shards)
    if args.shards > 1:
        log(f"[shard {args.shard}/{args.shards}] {len(my_files)} files: "
            f"{', '.join(f.stem[:24] for f in my_files[:4])}...")

    for xlsx in my_files:
        names = read_companies(xlsx)
        ck = load_ckpt(xlsx)
        verdicts = ck.get("verdicts", {})

        todo = [n for n in names if _norm(n) not in verdicts]
        if todo:
            log(f"[{xlsx.stem}] {len(todo)} of {len(names)} still to judge")

        for i in range(0, len(todo), CHUNK):
            if args.limit and queries >= args.limit:
                log(f"-- query limit {args.limit} reached, stopping")
                save_ckpt(xlsx, {"verdicts": verdicts})
                write_copy(xlsx, verdicts)
                return

            batch = todo[i:i + CHUNK]
            payload = [{"Company": n} for n in batch]
            prompt = (f"{_PROMPT}\n\nCompanies:\n"
                      f"{json.dumps(payload, ensure_ascii=False)}")

            wanted = {_norm(n) for n in batch}

            def _covered(blob: dict | None) -> int:
                """How many of THIS batch's companies the answer judged."""
                if not blob:
                    return 0
                got = {_norm(str(r.get("Company") or ""))
                       for r in (blob.get("results") or [])
                       if isinstance(r, dict)}
                return len(wanted & got)

            answer = ask(prompt)
            queries += 1
            parsed = extract_json(answer or "")

            # Retry when the answer is missing OR covers almost none of
            # the batch. Checking only "did we get JSON" was not enough:
            # measured on three Fire files, one query per file came back
            # as perfectly valid JSON that named 0-1 of its 12 companies,
            # so it counted as success and stranded the other 11. Each
            # gap was exactly one chunk, aligned to a chunk boundary --
            # the signature of a silently under-covered response rather
            # than scattered name-matching failures.
            #
            # Half the batch is the threshold: a model legitimately drops
            # the odd name, but returning fewer than half means the
            # answer itself was truncated and is worth re-asking.
            if _covered(parsed) < len(wanted) / 2:
                answer = ask(prompt)
                queries += 1
                retry = extract_json(answer or "")
                # Keep whichever answer covered more -- a retry that comes
                # back worse must not discard a usable first response.
                if _covered(retry) > _covered(parsed):
                    parsed = retry

            if not parsed:
                # An empty answer has TWO causes that need opposite
                # responses, and they look identical from here:
                #   * the daily per-IP AI Mode limit  -> stop for today
                #   * the network being down          -> wait, then go on
                # Observed live: a Wi-Fi drop produced ERR_NAME_NOT_
                # RESOLVED and three empty answers in a row, which this
                # would have reported as "daily limit reached" -- a
                # confident wrong diagnosis. Ask the network directly
                # instead of inferring from the symptom.
                if not internet_up():
                    log("  .. no network — waiting 60s (NOT counted as the "
                        "daily limit)")
                    time.sleep(60)
                    continue

                consecutive_empty += 1
                log(f"  .. no JSON ({consecutive_empty}/"
                    f"{MAX_CONSECUTIVE_EMPTY}) after {queries} queries")
                if consecutive_empty >= MAX_CONSECUTIVE_EMPTY:
                    # The daily per-IP wall, not a one-off failure: save
                    # and stop rather than spend the rest of the run
                    # writing empty results over good companies.
                    log("")
                    log(f"STOPPING: {MAX_CONSECUTIVE_EMPTY} empty answers in a "
                        f"row — this is the daily AI Mode limit.")
                    log(f"  {queries} queries used this run.")
                    log("  Re-run tomorrow; it resumes from the checkpoint.")
                    save_ckpt(xlsx, {"verdicts": verdicts})
                    write_copy(xlsx, verdicts)
                    return
                continue

            consecutive_empty = 0
            by_norm = {_norm(n): n for n in batch}
            for r in parsed.get("results") or []:
                if not isinstance(r, dict):
                    continue
                key = _norm(str(r.get("Company") or ""))
                original = by_norm.get(key)
                if not original:
                    continue
                verdicts[key] = {
                    "Company": original,
                    "status": str(r.get("status") or "").strip().lower(),
                    "acquired_by": clean_acquirer(r.get("acquired_by"),
                                                  original),
                }
            save_ckpt(xlsx, {"verdicts": verdicts})

        # A chunk that returned no JSON, or a company the model silently
        # omitted from its answer, leaves that company unjudged. It then
        # passes through write_copy untouched -- which is the SAFE
        # outcome, but it means the copy is only partly checked. Say so
        # rather than printing a line that reads like full coverage:
        # drone_ai was reported "wrote copy: 403 kept" when 104 of its
        # 406 companies had never been looked at.
        unjudged = sum(1 for n in names if _norm(n) not in verdicts)
        kept, dropped, renamed = write_copy(xlsx, verdicts)
        note = f", {unjudged} UNJUDGED (re-run to finish)" if unjudged else ""
        log(f"[{xlsx.stem}] wrote copy: {kept} kept, {dropped} closed-dropped, "
            f"{renamed} acquired-renamed{note}")

    log("")
    log(f"ALL FILES COMPLETE — {queries} queries this run")


if __name__ == "__main__":
    main()
