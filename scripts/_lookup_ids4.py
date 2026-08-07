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
            for table in ["cmi_report_dynamic", "market_company_link", "market_metric_link"]:
                cur.execute(
                    """
                    SELECT column_name, data_type
                    FROM information_schema.columns
                    WHERE table_schema='public' AND table_name=%s
                    ORDER BY ordinal_position
                    """,
                    (table,),
                )
                print(table, cur.fetchall()[:20])

            # any link between markets uuid and newsId?
            cur.execute(
                """
                SELECT table_name, column_name
                FROM information_schema.columns
                WHERE table_schema='public'
                  AND (column_name ILIKE '%newsid%' OR column_name ILIKE '%market_id%' OR column_name ILIKE '%report%')
                ORDER BY table_name, column_name
                """
            )
            print("id-ish columns:")
            for r in cur.fetchall():
                print(" ", r)


if __name__ == "__main__":
    main()
