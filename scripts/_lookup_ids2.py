# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path

import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))


def main() -> None:
    url = os.environ["DATABASE_URL"]
    with psycopg.connect(url) as conn:
        with conn.cursor() as cur:
            for q in [
                "%avocado%",
                "%green chemical%",
                "%green chem%",
                "%smartwatch%",
                "%firewall%",
            ]:
                cur.execute(
                    'SELECT "newsId", keyword FROM cmi_reports WHERE lower(keyword) LIKE lower(%s) ORDER BY "newsId"',
                    (q,),
                )
                rows = cur.fetchall()
                print(f"cmi keyword {q!r}: {len(rows)}")
                for r in rows[:20]:
                    print(" ", r)

            print("\nall neon markets containing tokens:")
            cur.execute(
                """
                SELECT id::text, name
                FROM markets
                WHERE lower(name) LIKE ANY(%s)
                ORDER BY name
                """,
                (
                    [
                        "%firewall%",
                        "%smartwatch%",
                        "%green chemical%",
                        "%avocado%",
                        "%green chem%",
                    ],
                ),
            )
            for r in cur.fetchall():
                print(" ", r)

            # Is there a serial/integer surrogate on markets?
            cur.execute(
                """
                SELECT a.attname, format_type(a.atttypid, a.atttypmod)
                FROM pg_attribute a
                JOIN pg_class c ON a.attrelid = c.oid
                JOIN pg_namespace n ON c.relnamespace = n.oid
                WHERE n.nspname='public' AND c.relname='markets' AND a.attnum > 0 AND NOT a.attisdropped
                """
            )
            print("\nmarkets attrs:", cur.fetchall())


if __name__ == "__main__":
    main()
