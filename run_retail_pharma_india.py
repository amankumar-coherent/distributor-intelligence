#!/usr/bin/env python3
"""India-specific leadership for the retail/pharma list, into a NEW workbook.

The first run captured whoever Google surfaced, which for a multinational is
usually the global CEO sitting in the US. This pass asks instead for the
person who actually runs the India business -- Skechers India's CEO, not
Skechers' global founder -- plus the company website.

The global sheet is kept untouched at
output/retail_pharma/backup/retail_pharma_global_*.xlsx.

Websites carry over from the global run where we already have them, so this
pass only has to find the ones still missing.

Same anti-fabrication rule throughout: a name is written back only if it
appears verbatim in the model's own reply. An Indian subsidiary with no
separate CHRO should come back blank, not filled with the global one.
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from run_leadership_aimode_batch import norm, parse_batch, tokens  # noqa: E402
from run_retail_pharma import ROLES, reset_profile  # noqa: E402
from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

OUT_DIR = ROOT / "output" / "retail_pharma_india"

PROMPT = """India leadership of each company below, plus website domain.

For a multinational give the INDIA person (India CEO, India MD, Country \
Head of the India business), NOT the global one.

JSON array only, same order:
[{{"company":"","website":"","CHRO":{{"name":"","email":""}},\
"CEO":{{"name":"","email":""}},"Founder":{{"name":"","email":""}},\
"Owner":{{"name":"","email":""}},"MD":{{"name":"","email":""}},\
"Co-founder":{{"name":"","email":""}},"Co-owner":{{"name":"","email":""}},\
"OTHER HEAD":{{"name":"","title":"","email":""}}}}]

OTHER HEAD = most senior India person, with their real title. If the India \
arm has no separate holder of a title, leave it empty -- do not substitute \
the global person. A guess is wrong.

{companies}"""


def write_excel(data: dict, companies: list[str], output: Path) -> None:
    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    rows = []
    for company in companies:
        entry = data.get(company.lower()) or {}
        row = {"Company": company, "Website": entry.get("website", "")}
        for role in ROLES:
            person = entry.get(role) or {}
            row[f"INDIA {role}"] = person.get("name", "")
            if role == "OTHER HEAD":
                row["INDIA OTHER HEAD TITLE"] = person.get("title", "")
            row[f"INDIA {role} EMAIL"] = person.get("email", "")
        row["Found"] = sum(1 for r in ROLES if (entry.get(r) or {}).get("name"))
        rows.append(row)

    df = pd.DataFrame(rows)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="India Leadership")
        ws = writer.sheets["India Leadership"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        body = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)] if lengths else 0
            ws.column_dimensions[get_column_letter(idx)].width = min(
                max(typical + 2, len(str(column)) + 2, 11), 34
            )
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24


def load_names(path: str) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for raw in Path(path).read_text(encoding="utf-8").splitlines():
        name = raw.strip()
        if not name or name.lower() in ("nan", "company") or name.lower() in seen:
            continue
        seen.add(name.lower())
        out.append(name)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--companies-file", required=True)
    ap.add_argument("--checkpoint", default=str(OUT_DIR / "retail_pharma_india.json"))
    ap.add_argument("--output", default=str(OUT_DIR / "retail_pharma_india.xlsx"))
    ap.add_argument("--global-source", default=str(ROOT / "output" / "retail_pharma" / "retail_pharma.json"),
                    help="Global run, read only to carry over websites we already know.")
    ap.add_argument("--batch-size", type=int, default=5)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--delay", type=float, default=3.0)
    ap.add_argument("--retry-empty", action="store_true")
    ap.add_argument("--auto-reset", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    apply_env_overrides()

    companies = load_names(args.companies_file)
    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    known_sites: dict[str, str] = {}
    gsrc = Path(args.global_source)
    if gsrc.exists():
        for key, entry in json.loads(gsrc.read_text(encoding="utf-8")).items():
            if entry.get("website"):
                known_sites[key] = entry["website"]

    def has_any(name: str) -> bool:
        entry = data.get(name.lower()) or {}
        return any((entry.get(r) or {}).get("name") for r in ROLES)

    if args.retry_empty:
        todo = [c for c in companies if not has_any(c)]
    else:
        todo = [c for c in companies if c.lower() not in data]
    if args.limit is not None:
        todo = todo[: args.limit]

    batches = [todo[i: i + args.batch_size] for i in range(0, len(todo), args.batch_size)]
    print(f"companies : {len(companies)}   to query: {len(todo)}   batches: {len(batches)} x {args.batch_size}")

    if args.dry_run:
        if batches:
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batches[0]))
            p = PROMPT.format(companies=listing)
            print("\n--- prompt ---")
            print(p)
            print(f"\nprompt chars: {len(p)}")
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
                low = answer.lower()
                if "request limit" in low or "try again in a little while" in low:
                    print(f"   RATE LIMITED (try {attempt+1}/4)", flush=True)
                    if args.auto_reset and reset_profile():
                        print("   profile + cookies reset - relaunching", flush=True)
                        time.sleep(5)
                        continue
                    time.sleep(120)
                    continue
                if "wasn't generated" in low or "was not generated" in low:
                    print(f"   Google declined (try {attempt+1}/4)", flush=True)
                    time.sleep(6 * (attempt + 1))
                    continue
                records = [r for r in parse_batch(answer) if "<" not in str(r.get("company") or "")]
                if records:
                    break
                print(f"   unparseable ({len(answer)} chars, try {attempt+1}/4)", flush=True)
                time.sleep(5)
            if not records:
                print("   giving up on this batch - will retry on resume", flush=True)
                continue

            wanted = {norm(c): c for c in batch}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                company = wanted.get(norm(str(rec.get("company") or "")))
                if company is None:
                    continue
                site = " ".join(str(rec.get("website") or "").split())
                entry = {"website": site or known_sites.get(company.lower(), "")}
                for role in ROLES:
                    slot = rec.get(role) or {}
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
                    out = {"name": name, "email": email}
                    if role == "OTHER HEAD":
                        out["title"] = " ".join(str(slot.get("title") or "").split())
                    entry[role] = out
                data[company.lower()] = entry

            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
            added += batch_added
            print(f"   +{batch_added} names   (rejected {rejected})", flush=True)
            if bi < len(batches) and args.delay:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    write_excel(data, companies, Path(args.output))
    print(f"\nnames added: {added}   rejected: {rejected}")
    print(f"companies with a name: {sum(1 for c in companies if has_any(c))} of {len(companies)}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
