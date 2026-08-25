# -*- coding: utf-8 -*-
from pathlib import Path
import os
import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))
url = (os.environ.get("DATABASE_URL") or "").strip()

with psycopg.connect(url) as conn:
    with conn.cursor() as cur:
        print("=== markets ===")
        cur.execute(
            """
            SELECT id::text, name
            FROM markets
            WHERE LOWER(name) LIKE '%electrical%'
               OR LOWER(name) LIKE '%electric distributor%'
               OR LOWER(name) LIKE '%electrical distributor%'
               OR LOWER(name) LIKE '%power distribution%'
               OR LOWER(name) LIKE '%wire and cable%'
               OR LOWER(name) LIKE '%electrical wholesale%'
            ORDER BY name
            """
        )
        for r in cur.fetchall():
            print(" MARKET", r[0], "|", r[1])

        print("=== vendor_intel_exports ===")
        cur.execute(
            """
            SELECT id::text, market_name, company_count, export_file
            FROM vendor_intel_exports
            WHERE LOWER(COALESCE(market_name,'')) LIKE '%electric%'
               OR LOWER(COALESCE(market_name,'')) LIKE '%distribut%'
               OR LOWER(COALESCE(query,'')) LIKE '%electric%'
            ORDER BY market_name
            """
        )
        rows = cur.fetchall()
        if not rows:
            print("  none")
        for r in rows:
            print(" VIE", r)

        print("=== cmi_reports keyword/custom/subject ===")
        cur.execute(
            """
            SELECT "newsId", keyword, "CustomName"
            FROM cmi_reports
            WHERE LOWER(COALESCE(keyword,'')) LIKE '%electrical%'
               OR LOWER(COALESCE(keyword,'')) LIKE '%electric distributor%'
               OR LOWER(COALESCE("CustomName",'')) LIKE '%electrical%'
               OR LOWER(COALESCE("newsSubject",'')) LIKE '%electrical distribution%'
               OR LOWER(COALESCE(keyword,'')) LIKE '%distributor%'
                  AND (
                    LOWER(COALESCE(keyword,'')) LIKE '%electric%'
                    OR LOWER(COALESCE(keyword,'')) LIKE '%power%'
                  )
            ORDER BY "newsId"
            """
        )
        for r in cur.fetchall():
            print(" CMI", r[0], "|", r[1], "|", r[2])
