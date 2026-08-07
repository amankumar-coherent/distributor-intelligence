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
            cur.execute(
                """
                SELECT "newsId", keyword
                FROM cmi_reports
                WHERE lower(keyword) LIKE '%green%'
                ORDER BY "newsId"
                """
            )
            print("green keywords:")
            for r in cur.fetchall():
                print(" ", r)

            # fuzzy: slug customname
            for slug in [
                "global-firewall-market",
                "global-smartwatch-market",
                "global-green-chemical-market",
                "global-avocado-oil-market",
                "avocado-oil-market",
                "smartwatches-market",
                "green-chemical-market",
            ]:
                cur.execute(
                    """
                    SELECT "newsId", keyword, "CustomName"
                    FROM cmi_reports
                    WHERE lower(coalesce("CustomName",'')) = %s
                       OR lower(coalesce("CustomName",'')) LIKE %s
                    """,
                    (slug, f"%{slug}%"),
                )
                rows = cur.fetchall()
                print(f"slug {slug}: {rows}")


if __name__ == "__main__":
    main()
