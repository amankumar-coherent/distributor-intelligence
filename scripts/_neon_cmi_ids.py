# -*- coding: utf-8 -*-
from pathlib import Path
import os
import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))
url = (os.environ.get("DATABASE_URL") or "").strip()

needles = [
    "electrical distribution",
    "avocado oil",
    "firewall",
    "green chemical",
    "smartphone",
    "antibody drug conjugate",
    "smartwatch",
]

with psycopg.connect(url) as conn:
    with conn.cursor() as cur:
        print("=== markets exact/ilike ===")
        for n in [
            "Global Electrical Distribution Market",
            "Electrical Distribution Market",
            "Global Avocado Oil Market",
            "Avocado Oil Market",
            "Global Firewall Market",
            "Global Green Chemical Market",
            "Global Smartphone Market",
            "Global Antibody Drug Conjugates Market",
            "Antibody Drug Conjugates Market",
            "Global Smartwatch Market",
        ]:
            cur.execute("SELECT id, name FROM markets WHERE LOWER(name)=LOWER(%s)", (n,))
            row = cur.fetchone()
            print("EXACT", n, "->", row)
        print("=== markets avocado/electrical ===")
        cur.execute(
            """
            SELECT id, name FROM markets
            WHERE LOWER(name) LIKE '%avocado%'
               OR LOWER(name) LIKE '%electrical distribution%'
               OR LOWER(name) LIKE '%power distribution%'
            ORDER BY name
            LIMIT 40
            """
        )
        for r in cur.fetchall():
            print(" ", r[0], r[1])

        print("=== cmi_reports keyword ===")
        for n in needles:
            cur.execute(
                """
                SELECT "newsId", keyword
                FROM cmi_reports
                WHERE LOWER(COALESCE(keyword,'')) LIKE %s
                ORDER BY "newsId"
                LIMIT 15
                """,
                (f"%{n}%",),
            )
            rows = cur.fetchall()
            print(f"-- {n} ({len(rows)})")
            for r in rows:
                print("  ", r[0], r[1])
