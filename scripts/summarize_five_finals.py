"""Summarize cleaned FINALs and optionally overwrite originals."""
from __future__ import annotations

import json
import shutil
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from clean_verify_five_finals import MARKETS, _empty, load_sheet_rows  # noqa: E402


def main() -> None:
    summary = []
    for m in MARKETS:
        cleaned = m["path"].with_name(m["path"].stem + "_CLEANED.xlsx")
        src = cleaned if cleaned.exists() else m["path"]
        _, rows, _ = load_sheet_rows(src, m["sheet"])

        def e(c: str) -> int:
            return sum(1 for r in rows if _empty(r.get(c)))

        item = {
            "market": m["query"],
            "key": m["key"],
            "file": str(src),
            "company_count": len(rows),
            "empty_contact": e("Contact Person"),
            "empty_role": e("Role"),
            "empty_email": e("Email"),
            "empty_founded": e("Founded"),
            "empty_hq": e("Headquarters"),
            "empty_website": e("Website"),
            "business_types": Counter(r.get("Business Type") or "" for r in rows).most_common(10),
            "contact_fill_rate_pct": round(
                100.0 * (len(rows) - e("Contact Person")) / max(1, len(rows)), 1
            ),
        }
        summary.append(item)
        print(
            f"{m['key']:12} companies={len(rows):4}  "
            f"contact_empty={e('Contact Person'):3} ({item['contact_fill_rate_pct']}% filled)  "
            f"founded_empty={e('Founded'):3}  hq_empty={e('Headquarters'):3}"
        )

        # Sync CLEANED → ORIGINAL FINAL (backup first)
        if cleaned.exists() and cleaned.resolve() != m["path"].resolve():
            bak = m["path"].with_suffix(
                m["path"].suffix + f".pre_clean_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            )
            if not list(m["path"].parent.glob(m["path"].name + ".pre_clean_*")):
                shutil.copy2(m["path"], bak)
                print(f"  backup original → {bak.name}")
            shutil.copy2(cleaned, m["path"])
            print(f"  synced CLEANED → {m['path'].name}")

    out = Path("output/five_markets_final_status.json")
    out.write_text(
        json.dumps(
            {
                "generated_at": datetime.now(timezone.utc).isoformat(),
                "total_companies": sum(x["company_count"] for x in summary),
                "markets": summary,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    print(f"\nTOTAL companies: {sum(x['company_count'] for x in summary)}")
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
