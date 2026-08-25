#!/usr/bin/env python3
"""Apply web-verified office contact overrides; clear junk phones. Never invent."""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

XLSX = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global"
    r"\global_avocado_oil_market_OFFICE_SLIM.xlsx"
)
AUDIT = XLSX.parent / "_avocado_office_manual_overrides_audit.json"
NPD = "Not publicly disclosed"

# Only values verified from official contact/privacy pages or agent web fetch.
OVERRIDES = {
    "Scavuzzo's Foodservice": {
        "office_no": "816-231-1517",
        "office_email": "",  # form only on scavuzzos.com/contact-us
        "source": "https://www.scavuzzos.com/contact-us",
    },
    "Bozzuto's": {
        "office_no": "(203) 272-3511",
        "office_email": "",  # form only
        "source": "https://bozzutos.com/contact",
    },
    "Golden State Foods": {
        "office_no": "(949) 247-8000",
        "office_email": "info@goldenstatefoods.com",
        "source": "https://www.goldenstatefoods.com/contact + privacy policy",
    },
    "Cash & Carry Smart Foodservice": {
        "office_no": "(800) 662-4242",
        "office_email": "",
        "source": "chefstore.com / Smart Foodservice customer service",
    },
    "Al Maya Group": {
        "office_no": "+971 4 823 0000",
        "office_email": "info@almaya.ae",
        "source": "https://www.almaya.ae/contact-us.php",
    },
    "Cheney Brothers": {
        "office_no": "(800) 432-1341",
        "office_email": "",  # contact form only
        "source": "https://www.cheneybrothers.com/contact-us/",
    },
    "Castell Howell Foods": {
        "office_no": "01269 846060",
        "office_email": "",  # contact form; no public mailbox on contact page
        "source": "https://www.castellhowellfoods.co.uk/en/contact",
    },
    "Merchants Foodservice": {
        "office_no": "(804) 484-7700",
        "office_email": "mfs-info@pfgc.com",
        "source": "prior verified / PFG merchants",
    },
    "Oleon Distribution": {
        "office_no": "",  # contact form only; VAT mistaken for phone
        "office_email": "",
        "clear_junk_phone": True,
        "source": "https://www.oleon.com/contact/",
    },
    "Europa Agricult Product (EAP)": {
        "office_no": "",
        "office_email": "",
        "clear_junk_phone": True,
        "source": "website parked/for-sale; no reliable contact",
    },
    "Bulk by CHO": {
        "office_no": "",
        "office_email": "",
        "clear_junk_phone": True,
        "source": "junk phone cleared pending better source",
    },
    "Service Alimentaire Desco": {
        "office_no": "450-437-7182",
        "office_email": "info@groupedesco.com",
        "source": "https://www.groupedesco.com/contact",
    },
    "Medikonda Nutrients (Medikonda Australia)": {
        "office_no": "+1 904 430 0199",
        "office_email": "hello@medikonda.com",
        "source": "https://medikonda.com.au/pages/support + privacy policy",
    },
    "Bulk by CHO": {
        "office_no": "+1 281 712 1549",
        "office_email": "sales@bulkbycho.com",
        "source": "https://bulkbycho.com/request-pricing/ + need-bulk-olive-oil page",
    },
    "Gordon Food Service, Inc.": {
        "office_no": "(800) 968-4164",
        "office_email": "",  # cts@gfs.com is tech support, not general office
        "source": "https://gfs.com/en-us/about-us/contact-us/",
    },
    "Dot Foods, Inc.": {
        "office_no": "217-773-4411",
        "office_email": "",
        "source": "https://integration.dotfoods.com/contact-us/",
    },
    "Fuji Oil Asia Distribution": {
        "office_no": "",
        "office_email": "",
        "clear_junk_phone": True,
        "source": "cleared date mistaken as phone",
    },
}


def emptyish(v: object) -> bool:
    s = str(v or "").strip().lower()
    return s in {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}


def main() -> None:
    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}
    applied = []

    for row in range(3, ws.max_row + 1):
        co = str(ws.cell(row, ic["Company"]).value or "").strip()
        if co not in OVERRIDES:
            continue
        ov = OVERRIDES[co]
        cur_p = str(ws.cell(row, ic["Office No."]).value or "").strip()
        cur_e = str(ws.cell(row, ic["Office Email"]).value or "").strip()
        changed = {}

        if ov.get("office_no"):
            ws.cell(row, ic["Office No."]).value = ov["office_no"]
            changed["phone"] = ov["office_no"]
        elif ov.get("clear_junk_phone") and (not emptyish(cur_p)):
            ws.cell(row, ic["Office No."]).value = NPD
            changed["phone"] = NPD

        if ov.get("office_email"):
            ws.cell(row, ic["Office Email"]).value = ov["office_email"]
            changed["email"] = ov["office_email"]
        elif "office_email" in ov and ov["office_email"] == "" and emptyish(cur_e):
            ws.cell(row, ic["Office Email"]).value = NPD

        for col in ("Office No.", "Office Email"):
            ws.cell(row, ic[col]).alignment = Alignment(wrap_text=True, vertical="top")

        if changed:
            applied.append({"company": co, **changed, "source": ov.get("source")})

    wb.save(XLSX)
    wb.close()

    # recount
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    ic2 = {str(h): i for i, h in enumerate(hdr) if h}
    rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if str(r[ic2["Company"]] or "").strip()]
    phone_n = sum(1 for r in rows if not emptyish(r[ic2["Office No."]]))
    email_n = sum(1 for r in rows if not emptyish(r[ic2["Office Email"]]))
    wb.close()

    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "applied": applied,
        "total_rows": len(rows),
        "phone_filled": phone_n,
        "email_filled": email_n,
        "phone_npd": len(rows) - phone_n,
        "email_npd": len(rows) - email_n,
    }
    AUDIT.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(json.dumps(audit, indent=2))


if __name__ == "__main__":
    main()
