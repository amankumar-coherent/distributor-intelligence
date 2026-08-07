# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path

import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))

TARGETS = [
    "Global Firewall Market",
    "Global Smartwatch Market",
    "Global Green Chemical Market",
    "Global Avocado Oil Market",
]


def main() -> None:
    url = os.environ["DATABASE_URL"]
    with psycopg.connect(url) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT "newsId", keyword, "newsSubject", "CustomName", "IsActive"
                FROM cmi_reports
                WHERE keyword ILIKE ANY(%s)
                   OR "newsSubject" ILIKE ANY(%s)
                   OR "CustomName" ILIKE ANY(%s)
                ORDER BY "newsId"
                """,
                (
                    ["%Firewall%", "%Smartwatch%", "%Green Chemical%", "%Avocado%"],
                    ["%Firewall%", "%Smartwatch%", "%Green Chemical%", "%Avocado%"],
                    ["%Firewall%", "%Smartwatch%", "%Green Chemical%", "%Avocado%"],
                ),
            )
            print("cmi_reports matches:")
            for r in cur.fetchall():
                print(" ", r)

            print("\nexact/near keyword match:")
            for t in TARGETS:
                cur.execute(
                    """
                    SELECT "newsId", keyword, "newsSubject"
                    FROM cmi_reports
                    WHERE lower(coalesce(keyword,'')) = lower(%s)
                       OR lower(coalesce("newsSubject",'')) = lower(%s)
                       OR lower(coalesce(keyword,'')) LIKE lower(%s)
                       OR lower(coalesce("newsSubject",'')) LIKE lower(%s)
                    ORDER BY "newsId"
                    """,
                    (t, t, f"%{t}%", f"%{t}%"),
                )
                rows = cur.fetchall()
                print(f"=== {t} ({len(rows)}) ===")
                for r in rows[:10]:
                    print(" ", r)

            cur.execute(
                "SELECT id, name FROM markets WHERE name ILIKE %s OR name ILIKE %s",
                ("%avocado%", "%Avocado%"),
            )
            print("\nneon markets avocado:")
            for r in cur.fetchall():
                print(" ", r)


if __name__ == "__main__":
    main()
