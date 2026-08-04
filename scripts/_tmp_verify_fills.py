import json
from pathlib import Path

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
checks = [
    "TD SYNNEX",
    "Softcat",
    "Softchoice",
    "Ingram Micro",
    "Wearabledevices",
    "A&S Creations",
    "Nagase",
    "Quimdis",
    "Biesterfeld",
    "Rahi Systems",
    "ACPL",
    "ComGuard",
    "Redington",
    "Acceron",
    "Connection",
    "ScanSource",
    "DigiKey",
    "D&H",
    "Macys",
    "Target",
]
for mid in [
    "global_smartwatch_market_global",
    "global_green_chemical_market_global",
    "global_firewall_market_global",
]:
    d = json.loads((ROOT / mid / f"{mid}.json").read_text(encoding="utf-8"))
    print("==", mid)
    for r in d.get("relevant_companies") or []:
        co = str(r.get("company") or "")
        if any(c.lower() in co.lower() for c in checks):
            c = r.get("_crawl") or {}
            print(
                "  ",
                co[:36].ljust(36),
                "|",
                str(c.get("contact_person") or "")[:32].ljust(32),
                "|",
                str(c.get("contact_role") or "")[:28].ljust(28),
                "|",
                str(c.get("phone") or "")[:24],
            )
