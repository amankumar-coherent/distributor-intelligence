"""Compare names more carefully; classify run3-only as add vs skip."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT3 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run3")
ROOT4 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")

# Deliberately removed / OEM / portal / SEO junk — do not re-add
SKIP = {
    # smartwatch
    "garmin",
    "oemwatchco",
    "sunlitesystems",
    "sunlite systems",
    "xiaomistores",  # already Cover Company in run4
    "getdistributors",  # portal
    "dafitapp",
    "wearfitpro",
    # green — manufacturers / junk previously removed
    "evonik industries",
    "evonik",
    "sabic",
    "trishulent india",
    "trilliumchemicals",
    "getdistributors",
    "greenchemicaldesign",
    "skytechgreen",
    "smartchemical",
    # firewall — OEMs / portals / SEO
    "fortinet, inc.",
    "fortinet",
    "mcafee",
    "qualys",
    "trellix",
    "zonealarm",
    "huntress",
    "redseal",
    "nordlayer",
    "comodo",
    "ipfire",
    "expertinsights",
    "firewalldealer",
    "itmonteur",
    "wroffy",
    "firewall",
    "firewalla",
    "next generation firewall service",
    "software firewall platinum support",
    "partner",  # generic listing
    "accenture plc",  # consulting giant, not channel partner
    "verizon communications",
    "encse",
    "ariacybersecurity",  # verify — may be Aria Cybersecurity product vendor
}


def norm(name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", " ", (name or "").lower()).strip()
    return re.sub(r"\s+", " ", s)


ALIASES = {
    "xiaomistores": "cover company",
    "synnex": "td synnex",
    "oqema": "oqema group",
    "thegreenchemicalstore": "green chemical store",
    "tech data": "td synnex",  # often same family; check separately
}


def load(mid: str, root: Path) -> list[dict]:
    data = json.loads((root / mid / f"{mid}.json").read_text(encoding="utf-8"))
    return list(data.get("relevant_companies") or [])


def main() -> None:
    for mid in [
        "global_smartwatch_market_global",
        "global_green_chemical_market_global",
        "global_firewall_market_global",
    ]:
        c3 = load(mid, ROOT3)
        c4 = load(mid, ROOT4)
        keys4 = {norm(r.get("company") or r.get("brand") or "") for r in c4}
        # also alias targets present
        keys4_exp = set(keys4)
        for a, b in ALIASES.items():
            if b in keys4:
                keys4_exp.add(a)
            if a in keys4:
                keys4_exp.add(b)

        add, skip = [], []
        for r in c3:
            name = str(r.get("company") or r.get("brand") or "").strip()
            k = norm(name)
            if not k:
                continue
            alias = ALIASES.get(k, k)
            if k in keys4 or alias in keys4 or k in keys4_exp:
                continue
            if k in SKIP or any(s in k for s in SKIP if len(s) > 4):
                skip.append(name)
                continue
            # OEM-ish role heuristics for firewall tech partners that are product vendors
            role = str(r.get("role") or "").lower()
            summary = str(r.get("summary") or r.get("company_summary") or "").lower()
            if mid.startswith("global_firewall") and any(
                x in summary for x in ("we develop", "our firewall", "endpoint protection platform", "saas security platform")
            ):
                skip.append(name + " (oem-summary)")
                continue
            add.append(name)
        print(f"\n=== {mid} === run3={len(c3)} run4={len(c4)}")
        print(f"  ADD ({len(add)}):")
        for n in add:
            print(f"    + {n}")
        print(f"  SKIP ({len(skip)}):")
        for n in skip:
            print(f"    - {n}")


if __name__ == "__main__":
    main()
