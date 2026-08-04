"""Dump companies needing summary rewrite with current text."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")


def main() -> None:
    for mid in [
        "global_smartwatch_market_global",
        "global_green_chemical_market_global",
        "global_firewall_market_global",
    ]:
        data = json.loads((ROOT / mid / f"{mid}.json").read_text(encoding="utf-8"))
        print(f"\n=== {mid} ===")
        for r in data.get("relevant_companies") or []:
            co = str(r.get("company") or "")
            sm = str(
                r.get("polished_summary")
                or r.get("company_summary")
                or r.get("summary")
                or (r.get("_crawl") or {}).get("linkedin_about")
                or ""
            ).strip()
            # flag interesting ones
            flags = []
            if not sm:
                flags.append("blank")
            if re.search(r"[\u00C0-\u024F].*[àáâãäåæçèéêëìíîïñòóôõöùúûüý]", sm, re.I) and not re.search(
                r"\b(the|and|company|distributor|retail)\b", sm, re.I
            ):
                flags.append("maybe_non_en")
            if re.search(r"(mais de|soluções|para a indústria|wir sind|notre|компания)", sm, re.I):
                flags.append("non_en")
            if "xiaomistores" in sm.lower() and "cover" in co.lower():
                flags.append("stale_name")
            if any(x in sm.lower() for x in ("digital marketing agency", "financial group", "not-for-profit society", "lumber", "hardware and lumber")):
                flags.append("wrong_identity")
            if sm.count("•") >= 3:
                flags.append("bullets")
            if flags:
                print(f"  {co[:40]:40} | {flags} | {sm[:160]!r}")


if __name__ == "__main__":
    main()
