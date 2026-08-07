# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path

import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))

UUIDS = {
    "firewall": "fde0876d-0fd7-48a4-9e82-2766d5048558",
    "green": "1cc15bca-e35e-4f82-a20c-8220cac49c68",
    "smartwatch": "cd382d27-444d-40fa-b1cd-144f94abcae4",
    "avocado": "343d7375-1928-4f89-a556-3ac5398383ea",
}


def main() -> None:
    with psycopg.connect(os.environ["DATABASE_URL"]) as conn:
        with conn.cursor() as cur:
            for label, uid in UUIDS.items():
                cur.execute(
                    "SELECT count(*) FROM market_company_link WHERE market_id = %s::uuid",
                    (uid,),
                )
                print(label, "company links", cur.fetchone()[0])
                cur.execute(
                    "SELECT id, name, status FROM markets WHERE id = %s::uuid",
                    (uid,),
                )
                print(" ", cur.fetchone())

            # Does any table store integer alongside market uuid?
            cur.execute(
                """
                SELECT c.table_name, c.column_name, c.data_type
                FROM information_schema.columns c
                WHERE c.table_schema='public'
                  AND c.table_name IN (
                    SELECT table_name FROM information_schema.columns
                    WHERE table_schema='public' AND column_name='market_id'
                  )
                ORDER BY c.table_name, c.ordinal_position
                """
            )
            print("tables with market_id:")
            for r in cur.fetchall():
                print(" ", r)


if __name__ == "__main__":
    main()
