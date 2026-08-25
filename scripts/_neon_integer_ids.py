# -*- coding: utf-8 -*-
from pathlib import Path
import os
import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))
url = (os.environ.get("DATABASE_URL") or "").strip()

needles = [
    "Global Electrical Distribution Market",
    "Electrical Distribution Market",
    "Global Avocado Oil Market",
    "Avocado Oil Market",
    "Global Firewall Market",
    "Firewall Market",
    "Global Green Chemical Market",
    "Green Chemical Market",
    "Global Smartphone Market",
    "Smartphone Market",
    "Global Antibody Drug Conjugates Market",
    "Antibody Drug Conjugates Market",
    "Global Smartwatch Market",
    "Smartwatch Market",
    "Smartwatches Market",
]

with psycopg.connect(url) as conn:
    with conn.cursor() as cur:
        print("=== exact keyword ===")
        for n in needles:
            cur.execute(
                """
                SELECT "newsId", keyword
                FROM cmi_reports
                WHERE LOWER(COALESCE(keyword,'')) = LOWER(%s)
                """,
                (n,),
            )
            rows = cur.fetchall()
            print(f"EXACT {n!r} -> {rows}")

        print("=== ilike ===")
        for n in [
            "%electrical distribution%",
            "%avocado oil%",
            "%firewall%",
            "%green chemical%",
            "%smartphone%",
            "%antibody drug conjugate%",
            "%smartwatch%",
        ]:
            cur.execute(
                """
                SELECT "newsId", keyword
                FROM cmi_reports
                WHERE LOWER(COALESCE(keyword,'')) LIKE %s
                ORDER BY "newsId"
                """,
                (n,),
            )
            print(f"-- {n}")
            for r in cur.fetchall():
                print(" ", r[0], r[1])
