"""Apply LinkedIn MCP-verified Founded / HQ / Cities updates to top-10 expanded Excels."""
from __future__ import annotations

from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel\output")

# Exact Excel company name -> updates (only fields to set)
# Source: LinkedIn company About pages via user-linkedin MCP (2026-07-29)
LI: dict[str, dict[str, str]] = {
    "SyFabrics": {
        "hq": "3430 W Pico Blvd, Los Angeles, CA 90019, USA",
        "cities": "Los Angeles, California, USA",
        "note": "LinkedIn /company/sy-fabrics HQ confirmed",
    },
    "Nail Tech Supply": {
        "hq": "2100 S Great Southwest Pkwy, Grand Prairie, TX 75051-3543, USA",
        "cities": "Grand Prairie, Texas; Kansas City / Gladstone, Missouri",
        "note": "LinkedIn /company/nail-tech-supply HQ confirmed",
    },
    "Nail Tech Supply / specialty wholesalers": {
        "hq": "2100 S Great Southwest Pkwy, Grand Prairie, TX 75051-3543, USA",
        "cities": "Grand Prairie, Texas; Kansas City / Gladstone, Missouri",
        "note": "LinkedIn /company/nail-tech-supply HQ confirmed",
    },
    "Unclefossil": {
        "hq": "520 E Valley Blvd, San Gabriel, CA 91776, USA",
        "cities": "San Gabriel, California 91776, USA",
        "note": "LinkedIn /company/uncle-fossil-wine-&-spirits HQ San Gabriel confirmed",
    },
    "Audiology Systems / regional dealers": {
        "founded": "2012",
        "hq": "Schaumburg, Illinois, USA (Otometrics/Audiology Systems)",
        "cities": "Schaumburg, Illinois, USA",
        "note": "LinkedIn /company/otometrics-audiology-systems Founded 2012; HQ Schaumburg",
    },
    "Sewing Parts Online": {
        # LinkedIn lists online launch 2008; brick store opened 1997 per LI About text
        "founded": "1997 (store); 2008 (online — LinkedIn)",
        "hq": "220 Tennsco Dr, Dickson, TN 37055, USA",
        "cities": "Dickson, Tennessee, USA; ships worldwide",
        "note": "LinkedIn /company/sewing-parts-online Founded 2008 online; HQ Dickson TN",
    },
    "Kaganails": {
        "hq": "Oldbury, B69 4NH, United Kingdom",
        "cities": "Oldbury, West Midlands, United Kingdom",
        "note": "LinkedIn /company/kaga-nails location Oldbury UK",
    },
    "Semiconductorinsight": {
        "founded": "2016",
        "hq": "Kharadi Road, City Vista A Wing, Pune, Maharashtra 411014, India",
        "cities": "Pune, Maharashtra, India; USA office listed on website (NY)",
        "note": "LinkedIn /company/semiconductor-insight Founded 2016; HQ Pune",
    },
    "The Nailhub / pro e-wholesalers": {
        "hq": "USA (thenailhub.com — LinkedIn /company/the-nail-hub; street HQ not on LI)",
        "cities": "USA online professional nail wholesale",
        "note": "LinkedIn /company/the-nail-hub website match; no street HQ on profile",
    },
    "Dearsundays": {
        "hq": "51 E 25th St, New York, NY 10010, USA; Hudson Yards studio NYC",
        "cities": "New York, NY, USA",
        "note": "No LI company page found; retained website contact addresses",
    },
}


def main() -> None:
    updated = 0
    for path in sorted(ROOT.rglob("*_distributors_expanded.xlsx")):
        wb = openpyxl.load_workbook(path)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=False))
        hdr_i = next(
            i
            for i, r in enumerate(rows[:12])
            if r and r[0].value and str(r[0].value).strip() == "Company"
        )
        header = [str(c.value or "").strip() for c in rows[hdr_i]]
        ci = header.index("Company")
        fi = header.index("Founded")
        hi = header.index("Headquarters")
        ri = header.index("Cities / Regions")
        si = header.index("Data Sources") if "Data Sources" in header else None

        changed = False
        for r in rows[hdr_i + 1 :]:
            company = str(r[ci].value or "").strip()
            if company not in LI:
                continue
            info = LI[company]
            if "founded" in info:
                r[fi].value = info["founded"]
            if "hq" in info:
                r[hi].value = info["hq"]
            if "cities" in info:
                r[ri].value = info["cities"]
            if si is not None and info.get("note"):
                prev = str(r[si].value or "").strip()
                tag = "LinkedIn MCP verified"
                if tag not in prev:
                    r[si].value = (prev + "; " if prev else "") + tag + f" ({info['note']})"
            updated += 1
            changed = True
            print(f"OK {path.parent.name}: {company}")

        if changed:
            wb.save(path)
        wb.close()
    print(f"\nUpdated rows: {updated}")


if __name__ == "__main__":
    main()
