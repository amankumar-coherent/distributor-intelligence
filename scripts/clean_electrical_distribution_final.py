"""One-shot cleanup of Global Electrical Distribution FINAL Excel.

Drops hard junk only: OEM blocklist, '(as distributor)' OEM arms, fake templated
country wholesalers. Keeps ambiguous real-looking distributors (LLM handles those
on future full pipeline runs with CHANNEL_PURITY_MODE=strict + CHANNEL_PURITY_LLM).
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))

from openpyxl import load_workbook

from vendor_intel.intelligence.channel_purity import classify_channel_purity
from vendor_intel.pipeline.web_expand import HEADERS, write_final_xlsx

XLSX = (
    ROOT
    / "output"
    / "web_expand"
    / "global_electrical_distribution_market_global"
    / "global_electrical_distribution_market_global_FINAL.xlsx"
)
OUT_DIR = XLSX.parent
BAK = OUT_DIR / "global_electrical_distribution_market_global_FINAL_ORIGINAL_573.xlsx"
# Legacy misnamed backup from earlier cleanup attempt
_LEGACY_BAK = XLSX.with_name(XLSX.name + ".bak_pre_purity_clean")
QUERY = "Global Electrical Distribution Market"
FAMILY = "electrical"

_HARD_REASONS = frozenset(
    {
        "known_oem_blocklist",
        "fake_templated_channel",
        "oem_as_distributor_label",
    }
)


def _is_hard_drop(reason: str) -> bool:
    if reason in _HARD_REASONS:
        return True
    return reason.startswith(("oem_signals", "oem_only", "oem_leaning"))


def _load_rows(path: Path) -> list[dict[str, str]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_raw = list(ws.iter_rows(values_only=True))
    headers = [str(h or "").strip() for h in rows_raw[0]]
    out: list[dict[str, str]] = []
    for r in rows_raw[1:]:
        row = {h: "" for h in HEADERS}
        for i, h in enumerate(headers):
            if h in row:
                row[h] = str(r[i] or "").strip() if i < len(r) else ""
        if row.get("Company"):
            out.append(row)
    wb.close()
    return out


def main() -> None:
    # Prefer full original 573-row backup
    if not BAK.exists() and _LEGACY_BAK.exists():
        BAK.write_bytes(_LEGACY_BAK.read_bytes())
        print(f"normalized backup -> {BAK}")

    if BAK.exists():
        source = BAK
        print(f"restoring from backup {BAK}")
    else:
        source = XLSX
        BAK.write_bytes(XLSX.read_bytes())
        print(f"backup created {BAK}")

    rows = _load_rows(source)
    before = len(rows)
    kept: list[dict[str, str]] = []
    dropped: list[dict] = []
    for row in rows:
        v = classify_channel_purity(
            row.get("Company") or "",
            family=FAMILY,
            specialty=row.get("Specialty Focus") or "",
            core_categories=row.get("Core Categories") or "",
            summary=row.get("Summary") or "",
            distribution_type=row.get("Distribution Type") or "",
            website=row.get("Website") or "",
        )
        if (not v.keep) and _is_hard_drop(v.reason):
            dropped.append(
                {"company": row.get("Company"), "verdict": v.label, "reason": v.reason}
            )
        else:
            kept.append(row)

    kept.sort(key=lambda r: (r.get("Company") or "").lower())
    write_final_xlsx(
        XLSX,
        kept,
        "Companies",
        audit={
            "cleaned": True,
            "mode": "hard_oem_fake_only",
            "query": QUERY,
            "family": FAMILY,
            "before": before,
            "after": len(kept),
            "dropped": len(dropped),
        },
    )
    dump = OUT_DIR / "chatgpt_rejected_purity_cleanup.json"
    dump.write_text(json.dumps(dropped, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"before={before} after={len(kept)} dropped={len(dropped)}")
    print(f"wrote {XLSX}")
    for d in dropped:
        print(f"  DROP {d.get('company')} | {d.get('reason')}")


if __name__ == "__main__":
    main()
