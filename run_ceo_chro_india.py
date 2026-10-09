#!/usr/bin/env python3
"""Find the India CEO and CHRO of each company via Google AI Mode.

Output per company: website, CEO name/email, CHRO name/email.

Two project-specific details are baked in deliberately:

* browser=chrome, not chromium. The captcha-raptor extension only loads
  under bundled Chromium, but Google serves Chromium a "Something went
  wrong" page that no solver can clear, while real Chrome completes long
  AI Mode runs. Measured again before writing this: Chrome returned a
  2,824-char answer with real names; that is the configuration used here.
* Google's error page is NOT an answer. It comes back as a non-empty
  string, so treating it as a result would quietly write fake blanks into
  the checkpoint. Each batch retries when it is detected.

ONLY REAL DATA: a name or email is written back only if it appears
verbatim in the model's own reply. Anything unsupported is dropped and
counted, never guessed.

Note: AI Mode has a cumulative per-IP daily limit of roughly 200-250
queries. At the default batch size, 1,458 companies is ~146 queries.

Usage:
    python run_ceo_chro_india.py --companies-file list.txt --dry-run
    python run_ceo_chro_india.py --companies-file list.txt --limit 10
    python run_ceo_chro_india.py --companies-file list.txt
"""
import argparse
import json
import os
import re
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

# Chrome, for the reason in the module docstring. Set before the scraper
# module reads its configuration.
os.environ.setdefault("GOOGLE_AI_MODE_BROWSER", "chrome")

from run_leadership_aimode_batch import norm, parse_batch, tokens  # noqa: E402
from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

OUT_DIR = ROOT / "output" / "ceo_chro_india"
FIELDS = ["CEO", "CHRO"]

# Keep this prompt SHORT.
#
# Measured failure: the first version was 1,022 chars, and AI Mode echoes
# the prompt back TWICE in the page transcript — 2,044 chars of overhead
# before any answer. Responses were capping near 2,300 chars, leaving ~250
# for data, so the JSON array was truncated as it began and 30 batches
# (~300 companies) were abandoned as "unparseable" despite Google having
# answered. A shorter prompt leaves room for the array to finish.
PROMPT = """India CEO and India CHRO (HR head) of each company below, with \
website. For a multinational give the India head, not the global one.

JSON array only, no commentary, same order:
[{{"company":"","website":"","CEO":{{"name":"","email":""}},\
"CHRO":{{"name":"","email":""}}}}]

Empty field if unknown — a guess is wrong. Email only if really published.

{companies}"""


def reset_profile() -> bool:
    """Wipe this process's AI Mode profile and cookies, then relaunch.

    Google's AI-response quota follows the profile's cookies, so a reset is
    what actually clears "You've reached the request limit" — retrying the
    same profile only burns attempts. Both the Chrome and the Chromium
    profile directory are removed: the run switches to Chromium to load the
    CAPTCHA solver, and clearing only one leaves the other's cookies (and
    therefore the limit) in place, which is exactly how an earlier manual
    reset appeared to fail.
    """
    from vendor_intel.scraping import google_ai_mode as gam

    try:
        gam.shutdown()
    except Exception:  # noqa: BLE001
        pass
    time.sleep(2)

    base = Path(
        (os.getenv("GOOGLE_AI_MODE_PROFILE_DIR") or "").strip()
        or str(ROOT / "data" / "ai_mode_chrome_profile")
    )
    stamp = time.strftime("%H%M%S")
    removed = False
    for path in (base, Path(f"{base}_chromium")):
        if path.is_dir():
            try:
                shutil.rmtree(path, ignore_errors=True)
                removed = True
            except Exception:  # noqa: BLE001
                try:
                    path.rename(Path(f"{path}.stale_{stamp}"))
                    removed = True
                except Exception:  # noqa: BLE001
                    pass
    for owner in base.parent.glob(f"{base.name}*.owner"):
        owner.unlink(missing_ok=True)
    return removed


def load_names(args) -> list[str]:
    if args.companies_file:
        lines = Path(args.companies_file).read_text(encoding="utf-8").splitlines()
    else:
        import pandas as pd

        path = Path(args.companies)
        df = pd.read_excel(path) if path.suffix.lower().startswith(".xls") else pd.read_csv(path)
        column = next((c for c in df.columns if str(c).strip().lower() == "company"), None)
        if column is None:
            column = next((c for c in df.columns if df[c].notna().any()), df.columns[0])
        lines = [str(v) for v in df[column].tolist()]
    seen: set[str] = set()
    out: list[str] = []
    for raw in lines:
        name = str(raw).strip()
        if not name or name.lower() in ("nan", "company") or name.lower() in seen:
            continue
        seen.add(name.lower())
        out.append(name)
    return out


def write_excel(data: dict, companies: list[str], output: Path) -> None:
    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    rows = []
    for company in companies:
        entry = data.get(company.lower()) or {}
        row = {"Company": company, "Website": entry.get("website", "")}
        for field in FIELDS:
            person = entry.get(field) or {}
            row[f"{field} Name"] = person.get("name", "")
            row[f"{field} Email"] = person.get("email", "")
        row["Found"] = sum(1 for f in FIELDS if (entry.get(f) or {}).get("name"))
        rows.append(row)

    df = pd.DataFrame(rows)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="CEO and CHRO India")
        ws = writer.sheets["CEO and CHRO India"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        body = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)] if lengths else 0
            ws.column_dimensions[get_column_letter(idx)].width = min(
                max(typical + 2, len(str(column)) + 2, 12), 40
            )
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies", default=None, help=".xlsx/.csv list")
    parser.add_argument("--companies-file", default=None, help="plain text, one company per line")
    parser.add_argument("--checkpoint", default=str(OUT_DIR / "ceo_chro_india.json"))
    parser.add_argument("--output", default=str(OUT_DIR / "ceo_chro_india.xlsx"))
    parser.add_argument(
        "--batch-size",
        type=int,
        default=6,
        help=(
            "Companies per query. 6, not 10: AI Mode caps the response "
            "around 2,300-3,900 chars including the echoed prompt, and at "
            "10 the JSON array was being truncated before it finished."
        ),
    )
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--delay", type=float, default=4.0)
    parser.add_argument("--retry-empty", action="store_true", help="Re-query companies recorded with no names")
    parser.add_argument(
        "--auto-reset",
        action="store_true",
        help=(
            "On Google's 'request limit' page, wipe this profile's cookies "
            "and relaunch instead of waiting. Required for unattended runs."
        ),
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    apply_env_overrides()
    os.environ["GOOGLE_AI_MODE_BROWSER"] = "chrome"

    companies = load_names(args)
    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    def has_any(name: str) -> bool:
        entry = data.get(name.lower()) or {}
        return any((entry.get(f) or {}).get("name") for f in FIELDS)

    if args.retry_empty:
        todo = [c for c in companies if not has_any(c)]
    else:
        todo = [c for c in companies if c.lower() not in data]
    if args.limit is not None:
        todo = todo[: args.limit]

    batches = [todo[i : i + args.batch_size] for i in range(0, len(todo), args.batch_size)]
    print(f"companies : {len(companies)}   to query: {len(todo)}   batches: {len(batches)} x {args.batch_size}")
    if len(batches) > 200:
        print("NOTE: AI Mode allows roughly 200-250 queries per IP per day; this needs more than one day.")

    if args.dry_run:
        if batches:
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batches[0]))
            print("\n--- prompt ---")
            print(PROMPT.format(companies=listing))
        return
    if not batches:
        write_excel(data, companies, Path(args.output))
        print(f"nothing to do; Excel refreshed: {args.output}")
        return

    from vendor_intel.scraping.google_ai_mode import ai_mode_enabled, ask, shutdown

    if not ai_mode_enabled():
        raise SystemExit("Set GOOGLE_AI_MODE_ENABLED=true in .env")

    added = rejected = 0
    try:
        for bi, batch in enumerate(batches, start=1):
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batch))
            print(f"\n[batch {bi}/{len(batches)}] {len(batch)} companies ...", flush=True)

            answer = ""
            records: list[dict] = []
            for attempt in range(4):
                answer = ask(PROMPT.format(companies=listing)) or ""
                # Google's error page is a non-empty string, not an answer.
                low = answer.lower()
                # "You've reached the request limit for AI responses at the
                # moment." Retrying does nothing for this — the quota is
                # tied to the profile's cookies — so wipe the profile and
                # relaunch rather than burning the remaining attempts.
                if "request limit" in low or "try again in a little while" in low:
                    print(f"   RATE LIMITED (try {attempt + 1}/4)", flush=True)
                    if args.auto_reset and reset_profile():
                        print("   profile + cookies reset — relaunching", flush=True)
                        time.sleep(5)
                        continue
                    print("   waiting 120s", flush=True)
                    time.sleep(120)
                    continue
                if "wasn't generated" in low or "was not generated" in low:
                    print(f"   Google declined (try {attempt + 1}/4) — retrying")
                    time.sleep(6 * (attempt + 1))
                    continue
                records = [r for r in parse_batch(answer) if "<" not in str(r.get("company") or "")]
                if records:
                    break
                print(f"   unparseable ({len(answer)} chars, try {attempt + 1}/4) — retrying")
                time.sleep(5)
            if not records:
                print("   giving up on this batch — will retry on resume")
                continue

            wanted = {norm(c): c for c in batch}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                company = wanted.get(norm(str(rec.get("company") or "")))
                if company is None:
                    continue
                entry = {"website": " ".join(str(rec.get("website") or "").split())}
                for field in FIELDS:
                    slot = rec.get(field) or {}
                    if not isinstance(slot, dict):
                        slot = {}
                    name = " ".join(str(slot.get("name") or "").split())
                    email = " ".join(str(slot.get("email") or "").split())
                    if name:
                        words = re.findall(r"[a-z0-9]{2,}", name.lower())
                        if words and any(w not in answer_tokens for w in words):
                            rejected += 1
                            name = email = ""
                    if email and email.lower() not in answer.lower():
                        email = ""
                    if name:
                        batch_added += 1
                    entry[field] = {"name": name, "email": email}
                data[company.lower()] = entry

            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
            added += batch_added
            print(f"   +{batch_added} names   (rejected {rejected})")
            if bi < len(batches) and args.delay:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    write_excel(data, companies, Path(args.output))
    withname = sum(1 for c in companies if has_any(c))
    emails = sum(1 for v in data.values() for f in FIELDS if (v.get(f) or {}).get("email"))
    print(f"\nnames added   : {added}")
    print(f"rejected      : {rejected} (not grounded in the answer)")
    print(f"companies with a name: {withname} of {len(companies)}")
    print(f"emails found  : {emails}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
