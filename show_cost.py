#!/usr/bin/env python3
"""Print OpenAI cost summary in the terminal."""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.pipeline import openai_cost  # noqa: E402


def main() -> int:
    totals = openai_cost.totals()
    print("=== OpenAI cost totals ===")
    print(f"  runs   : {totals['runs']}")
    print(f"  calls  : {totals['calls']}")
    print(f"  tokens : {totals['total_tokens']:,}")
    print(f"  est $  : ${totals['estimated_cost_usd']:.4f}")
    print(f"  folder : {totals['cost_dir']}")
    live = Path(totals["cost_dir"]) / "live.json"
    if live.exists():
        data = json.loads(live.read_text(encoding="utf-8"))
        print("\n=== Live / last run ===")
        print(f"  market : {data.get('market')}")
        print(f"  run_id : {data.get('run_id')}")
        print(f"  calls  : {data.get('calls')}")
        print(f"  tokens : {data.get('total_tokens')}")
        print(f"  est $  : ${float(data.get('estimated_cost_usd') or 0):.4f}")
    runs = openai_cost.load_runs(10)
    if runs:
        print("\n=== Last 10 runs ===")
        for r in reversed(runs[-10:]):
            print(
                f"  ${float(r.get('estimated_cost_usd') or 0):.4f}  "
                f"tok={int(r.get('total_tokens') or 0):,}  "
                f"calls={r.get('calls')}  "
                f"{(r.get('market') or '')[:50]}"
            )
    print("\nDashboard UI:")
    print("  .\\.venv\\Scripts\\python.exe -m streamlit run cost_dashboard.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
