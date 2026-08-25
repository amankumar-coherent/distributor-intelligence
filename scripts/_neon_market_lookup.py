# -*- coding: utf-8 -*-
"""Lookup 7 market names in Neon (markets + vendor_intel_exports + cmi_reports)."""
from __future__ import annotations

import os
from pathlib import Path

import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))

NAMES = [
    "Global Electrical Distribution Market",
    "Global Avocado Oil Market",
    "Global Firewall Market",
    "Global Green Chemical Market",
    "Global Smartphone Market",
    "Global Antibody Drug Conjugates Market",
    "Global Smartwatch Market",
    "Electrical Distribution",
    "Avocado Oil",
    "Firewall",
    "Green Chemical",
    "Smartphone",
    "Antibody Drug Conjugates",
    "ADC",
    "Smartwatch",
]


def main() -> None:
    url = (os.environ.get("DATABASE_URL") or "").strip()
    if not url:
        raise SystemExit("DATABASE_URL missing")
    with psycopg.connect(url) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT table_name FROM information_schema.tables
                WHERE table_schema='public' AND table_type='BASE TABLE'
                ORDER BY 1
                """
            )
            tables = [r[0] for r in cur.fetchall()]
            print("TABLES", ", ".join(tables))

            if "markets" in tables:
                cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='markets' ORDER BY ordinal_position")
                print("markets cols:", [r[0] for r in cur.fetchall()])
                cur.execute("SELECT id, name FROM markets ORDER BY name")
                rows = cur.fetchall()
                print(f"markets count={len(rows)}")
                for r in rows:
                    name = str(r[1] or "")
                    if any(k.lower() in name.lower() for k in (
                        "electrical", "avocado", "firewall", "green chem", "smartphone",
                        "antibody", "conjugat", "smartwatch", "adc",
                    )):
                        print("  MARKET HIT", r[0], r[1])

            if "vendor_intel_exports" in tables:
                cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='vendor_intel_exports' ORDER BY ordinal_position")
                print("vie cols:", [r[0] for r in cur.fetchall()])
                cur.execute("SELECT id, market_name FROM vendor_intel_exports ORDER BY market_name")
                rows = cur.fetchall()
                print(f"vendor_intel_exports count={len(rows)}")
                for r in rows:
                    name = str(r[1] or "")
                    if any(k.lower() in name.lower() for k in (
                        "electrical", "avocado", "firewall", "green chem", "smartphone",
                        "antibody", "conjugat", "smartwatch", "adc",
                    )):
                        print("  VIE HIT", r[0], r[1])
                print("  ALL VIE:")
                for r in rows[:80]:
                    print("   ", r[0], r[1])

            if "cmi_reports" in tables:
                cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='cmi_reports' ORDER BY ordinal_position")
                cols = [r[0] for r in cur.fetchall()]
                print("cmi_reports cols:", cols)
                # try keyword/title
                for col in ("keyword", "title", "name", "slug"):
                    if col in cols:
                        cur.execute(
                            f"""
                            SELECT id, "{col}" FROM cmi_reports
                            WHERE LOWER(COALESCE("{col}",'')) LIKE ANY(%s)
                            LIMIT 40
                            """,
                            ([
                                "%electrical distribution%",
                                "%avocado oil%",
                                "%firewall%",
                                "%green chemical%",
                                "%smartphone%",
                                "%antibody drug conjugate%",
                                "%smartwatch%",
                            ],),
                        )
                        print(f"cmi_reports by {col}:")
                        for r in cur.fetchall():
                            print("  CMI", r[0], r[1])


if __name__ == "__main__":
    main()
