# -*- coding: utf-8 -*-
from pathlib import Path
import os
import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))
url = (os.environ.get("DATABASE_URL") or "").strip()

needles = [
    "%electrical distribution%",
    "%avocado oil%",
    "%firewall%",
    "%green chemical%",
    "%green chem%",
    "%smartphone%",
    "%antibody drug conjugate%",
    "%smartwatch%",
]

with psycopg.connect(url) as conn:
    with conn.cursor() as cur:
        for col in ("keyword", "CustomName", "newsSubject", "meta_title"):
            print(f"=== {col} ===")
            for n in needles:
                cur.execute(
                    f"""
                    SELECT "newsId", keyword, "CustomName", "newsSubject"
                    FROM cmi_reports
                    WHERE LOWER(COALESCE("{col}",'')) LIKE %s
                    ORDER BY "newsId"
                    LIMIT 20
                    """,
                    (n,),
                )
                rows = cur.fetchall()
                if rows:
                    print(f"  {n}")
                    for r in rows:
                        print("   ", r)
