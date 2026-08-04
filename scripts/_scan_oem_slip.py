import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from clean_verify_five_finals import DIST_RE, MARKETS, OEM_BY_MARKET, _norm_name, load_sheet_rows

for m in MARKETS:
    cleaned = m["path"].with_name(m["path"].stem + "_CLEANED.xlsx")
    src = cleaned if cleaned.exists() else m["path"]
    _, rows, _ = load_sheet_rows(src, m["sheet"])
    oems = OEM_BY_MARKET.get(m["key"], set())
    suspects = []
    for r in rows:
        nl = _norm_name(r.get("Company") or "")
        for o in oems:
            if o in nl and not DIST_RE.search(r.get("Company") or ""):
                suspects.append(r.get("Company"))
                break
    flags = []
    for r in rows:
        n = (r.get("Company") or "").lower()
        if any(x in n for x in ("alibaba", "amazon", "ebay", "walmart", "flipkart", "apple ", "samsung ")):
            flags.append(r.get("Company"))
    print(m["key"], "n=", len(rows), "oem-ish=", suspects[:25], "flag=", flags[:25])
