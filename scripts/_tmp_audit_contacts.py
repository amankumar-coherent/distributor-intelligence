import json
from pathlib import Path

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
KEYS = [
    "comguard",
    "acceron",
    "scansource",
    "softchoice",
    "connection",
    "maroon",
    "zales",
    "rahi",
    "gps",
    "b&h",
    "jcpenney",
    "synnex",
    "td synnex",
    "redington",
    "vianet",
    "wildberries",
    "softcat",
    "nagase",
    "a&s",
    "digikey",
    "d&h",
    "esprinet",
    "mouser",
    "farnell",
    "arrow",
    "insight",
    "westcon",
    "imcd",
    "macys",
    "macy",
    "target",
    "quimdis",
    "arpadis",
    "eigenmann",
    "hsh",
    "acpl",
]

for mid in [
    "global_smartwatch_market_global",
    "global_green_chemical_market_global",
    "global_firewall_market_global",
]:
    path = ROOT / mid / f"{mid}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    print(f"\n=== {mid} ===")
    for c in data.get("companies", []):
        n = (c.get("name") or "").lower()
        if any(k in n for k in KEYS):
            print(
                f"{(c.get('name') or '')[:42]:42} | "
                f"{(c.get('contact_name') or '')[:36]:36} | "
                f"{(c.get('contact_role') or '')[:22]:22} | "
                f"{(c.get('phone') or '')[:22]:22} | "
                f"{(c.get('website') or '')[:45]}"
            )
