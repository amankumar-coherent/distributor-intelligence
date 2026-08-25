# -*- coding: utf-8 -*-
from pathlib import Path
import os
import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))
url = (os.environ.get("DATABASE_URL") or "").strip()

with psycopg.connect(url) as conn:
    with conn.cursor() as cur:
        print("=== cmi_reports distribution / wholesale / electrical equipment ===")
        cur.execute(
            """
            SELECT "newsId", keyword
            FROM cmi_reports
            WHERE LOWER(COALESCE(keyword,'')) LIKE '%distribut%'
               OR LOWER(COALESCE(keyword,'')) LIKE '%wholesale%'
               OR LOWER(COALESCE(keyword,'')) LIKE '%electrical equipment%'
               OR LOWER(COALESCE(keyword,'')) LIKE '%switchgear%'
               OR LOWER(COALESCE(keyword,'')) LIKE '%wire and cable%'
               OR LOWER(COALESCE(keyword,'')) LIKE '%power cable%'
            ORDER BY keyword
            """
        )
        for r in cur.fetchall():
            print(" ", r[0], r[1])

        print("=== vendor_intel_companies sample electrical names ===")
        cur.execute(
            """
            SELECT column_name FROM information_schema.columns
            WHERE table_name='vendor_intel_companies'
            ORDER BY ordinal_position
            """
        )
        print("cols", [r[0] for r in cur.fetchall()])

        cur.execute(
            """
            SELECT e.market_name, c.company_name
            FROM vendor_intel_companies c
            JOIN vendor_intel_exports e ON e.id = c.export_id
            WHERE LOWER(c.company_name) LIKE ANY(%s)
            LIMIT 30
            """,
            ([
                "%wesco%", "%rexel%", "%sonepar%", "%grainger%",
                "%anixter%", "%border states%",
            ],),
        )
        rows = cur.fetchall()
        print("company hits", len(rows))
        for r in rows:
            print(" ", r)

        print("=== market_company_link via companies name ===")
        cur.execute(
            """
            SELECT column_name FROM information_schema.columns
            WHERE table_name='companies' ORDER BY ordinal_position
            """
        )
        print("companies cols", [r[0] for r in cur.fetchall()])
