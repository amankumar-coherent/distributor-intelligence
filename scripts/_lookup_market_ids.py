# -*- coding: utf-8 -*-
"""Lookup Neon + local batch IDs for the four FINAL markets."""
from __future__ import annotations

import os
import sqlite3
from pathlib import Path

import psycopg
from dotenv import load_dotenv

load_dotenv(Path(r"D:\vendor-intel\.env"))
load_dotenv(Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env"))

NEEDLES = ("firewall", "smartwatch", "green chemical", "avocado")


def main() -> None:
    print("=== batch.db ===")
    con = sqlite3.connect(r"D:\vendor-intel\data\batch.db")
    con.row_factory = sqlite3.Row
    tables = [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")]
    print("tables:", tables)
    if "markets" in tables:
        print("cols:", con.execute("PRAGMA table_info(markets)").fetchall())
        q = (
            "SELECT id, query, market_slug, batch_id, status FROM markets "
            "WHERE lower(coalesce(query,'')) LIKE ? "
            "OR lower(coalesce(query,'')) LIKE ? "
            "OR lower(coalesce(query,'')) LIKE ? "
            "OR lower(coalesce(query,'')) LIKE ? "
            "ORDER BY id"
        )
        rows = con.execute(
            q, ("%firewall%", "%smartwatch%", "%green chemical%", "%avocado%")
        ).fetchall()
        for row in rows:
            print(dict(row))
    con.close()

    print("\n=== Neon markets ===")
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        print("No DATABASE_URL")
        return
    with psycopg.connect(url) as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT column_name, data_type
                FROM information_schema.columns
                WHERE table_schema='public' AND table_name='markets'
                ORDER BY ordinal_position
                """
            )
            print("columns:")
            for r in cur.fetchall():
                print(" ", r)
            cur.execute(
                """
                SELECT id, name
                FROM markets
                WHERE name ILIKE '%Firewall%'
                   OR name ILIKE '%Smartwatch%'
                   OR name ILIKE '%Green Chemical%'
                   OR name ILIKE '%Avocado%'
                ORDER BY name
                """
            )
            print("rows:")
            for r in cur.fetchall():
                print(" ", r[0], "|", r[1])

            # also vendor_intel_exports
            cur.execute(
                """
                SELECT column_name, data_type
                FROM information_schema.columns
                WHERE table_schema='public' AND table_name='vendor_intel_exports'
                ORDER BY ordinal_position
                """
            )
            print("\nvendor_intel_exports columns:")
            for r in cur.fetchall():
                print(" ", r)
            cur.execute(
                """
                SELECT id, market_name, query, created_at
                FROM vendor_intel_exports
                WHERE market_name ILIKE '%Firewall%'
                   OR market_name ILIKE '%Smartwatch%'
                   OR market_name ILIKE '%Green Chemical%'
                   OR market_name ILIKE '%Avocado%'
                   OR query ILIKE '%Firewall%'
                   OR query ILIKE '%Smartwatch%'
                   OR query ILIKE '%Green Chemical%'
                   OR query ILIKE '%Avocado%'
                ORDER BY created_at DESC
                LIMIT 30
                """
            )
            print("export rows:")
            for r in cur.fetchall():
                print(" ", r)


if __name__ == "__main__":
    main()
