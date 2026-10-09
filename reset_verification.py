#!/usr/bin/env python3
"""Clear the verification results but KEEP the discovered Kompass candidates.

Discovery is the expensive, slow half: sweeping the global listing plus 60
country listings produced 9,460 candidates over many hours. The screening
verdicts built on top of them, however, were produced before today's fixes:

  * the market-relevance gate (a distributor of the WRONG industry, e.g.
    Fritz Minke, was accepted as a Wearable Medical Devices distributor),
  * the bot-block guard (Google's CAPTCHA page echoes the query, so it was
    parsed as a confident "distributor" verdict for every blocked company),
  * Employees / Ownership / Website / Email / phone extraction.

So this resets the verify stage only. `kompass_candidates` is preserved,
which means the resumed run skips discovery entirely and goes straight to
re-screening all 9,460 with the corrected logic.

A timestamped backup is written first — this discards 1,505 companies'
worth of screening, and that is not something to do without an undo.

Usage:
    .\\.venv\\Scripts\\python.exe reset_verification.py --query "Wearable Medical Devices"
    .\\.venv\\Scripts\\python.exe reset_verification.py --query "..." --apply
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

# Everything discovery produced — must survive the reset.
_KEEP = ("kompass_candidates", "kompass_product_url", "kompass_product_id")

# Everything the verify stage produced — cleared so it is redone.
_CLEAR = {
    "kompass_processed_urls": [],
    "kompass_accepted": [],
    "kompass_rejected_count": 0,
    "kompass_unreachable_count": 0,
    "kompass_rescreen_rejected": [],
    "kompass_rescreened_urls": [],
}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", required=True)
    ap.add_argument("--country", default="global")
    ap.add_argument("--out-dir", default=None)
    ap.add_argument("--final", default=None)
    ap.add_argument(
        "--apply",
        action="store_true",
        help="actually write the reset (without this it only reports)",
    )
    args = ap.parse_args()

    import run_cursor_expand as rce

    rce._load_env()

    from vendor_intel.pipeline.openai_expand_checkpoint import ExpandCheckpoint
    from vendor_intel.pipeline.web_expand import resolve_final_path

    final_path = resolve_final_path(
        args.query, args.country, args.out_dir and Path(args.out_dir), args.final
    )
    ckpt_path = ExpandCheckpoint.path_for(final_path.parent, "kompass")
    if not ckpt_path.exists():
        print(f"No Kompass checkpoint at {ckpt_path}")
        return 1

    raw = json.loads(ckpt_path.read_text(encoding="utf-8"))
    data = raw.get("data") or {}

    candidates = data.get("kompass_candidates") or []
    print(f"checkpoint: {ckpt_path}")
    print(f"\nKEEPING  {len(candidates):>6}  discovered candidates")
    for key in _KEEP[1:]:
        print(f"         {str(data.get(key) or '')[:60]:>6}  {key}")
    print("\nCLEARING")
    for key in _CLEAR:
        cur = data.get(key)
        n = len(cur) if hasattr(cur, "__len__") else cur
        print(f"         {str(n):>6}  {key}")

    if not candidates:
        print("\nNo candidates to keep — refusing to reset, this would lose everything.")
        return 1

    if not args.apply:
        print("\nDRY RUN — nothing written. Re-run with --apply to reset.")
        return 0

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = ckpt_path.with_name(f"{ckpt_path.stem}.before_reset_{stamp}.json")
    shutil.copy2(ckpt_path, backup)
    print(f"\nbackup -> {backup.name}")

    for key, empty in _CLEAR.items():
        data[key] = empty
    raw["data"] = data
    # Send the resume back to the start of the verify stage. Discovery is
    # skipped because kompass_candidates is still populated.
    raw["step"] = "kompass_3_verify"
    raw["substep"] = "reset_for_reverification"
    raw["status"] = "running"
    raw["completed"] = {}
    raw["updated_at"] = datetime.now().isoformat(timespec="seconds")
    raw.setdefault("history", []).append(
        {
            "at": raw["updated_at"],
            "step": "kompass_3_verify",
            "substep": "reset_for_reverification",
            "note": f"cleared verification, kept {len(candidates)} candidates",
        }
    )

    tmp = ckpt_path.with_suffix(".tmp")
    tmp.write_text(json.dumps(raw, indent=2), encoding="utf-8")
    tmp.replace(ckpt_path)

    print(f"reset done — {len(candidates)} candidates kept, verification cleared")
    print("\nresume with:")
    print(f'  .\\.venv\\Scripts\\python.exe run_cursor_expand.py --query "{args.query}" --backend kompass')
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
