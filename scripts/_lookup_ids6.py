# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path

import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))


def main() -> None:
    with psycopg.connect(os.environ["DATABASE_URL"]) as conn:
        with conn.cursor() as cur:
            for term in [
                "Global Firewall Market",
                "Global Green Chemical Market",
                "Global Smartwatch Market",
                "Global Avocado Oil Market",
            ]:
                cur.execute(
                    """
                    SELECT "newsId", keyword, "CustomName"
                    FROM cmi_reports
                    WHERE keyword ILIKE %s
                       OR "newsSubject" ILIKE %s
                       OR "CustomName" ILIKE %s
                       OR meta_title ILIKE %s
                    LIMIT 5
                    """,
                    (f"%{term}%", f"%{term}%", f"%{term.replace(' ', '-').lower()}%", f"%{term}%"),
                )
                print(term, "=>", cur.fetchall())


if __name__ == "__main__":
    main()
