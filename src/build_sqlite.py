#!/usr/bin/env python3
"""P16 — load the SAMPLE sales file into SQLite (sql/schema.sql).

    python3 src/build_sqlite.py

The .db is NOT uploaded to GitHub (rebuildable from the pushed
scripts) — same rule as every other project in this portfolio.
"""
import csv
import sqlite3
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw" / "daily_sales_SAMPLE.csv"
DB = PROJECT / "data" / "processed" / "sales.db"


def main():
    rows = list(csv.DictReader(open(RAW, encoding="utf-8")))
    if DB.exists():
        DB.unlink()
    con = sqlite3.connect(DB)
    con.executescript((PROJECT / "sql" / "schema.sql").read_text())
    products = {}
    for r in rows:
        products[r["product"]] = r["category"]
    con.executemany("INSERT INTO dim_product VALUES (?, ?)",
                    sorted(products.items()))
    con.executemany(
        "INSERT INTO fact_daily_sales "
        "(sale_date, week_no, week_start, product, qty, "
        " unit_price_pkr, amount_pkr) VALUES (?,?,?,?,?,?,?)",
        [(r["date"], int(r["week_no"]), r["week_start"], r["product"],
          int(r["qty"]), int(r["unit_price_pkr"]), int(r["amount_pkr"]))
         for r in rows])
    con.commit()
    n = con.execute("SELECT COUNT(*) FROM fact_daily_sales").fetchone()[0]
    total = con.execute(
        "SELECT SUM(amount_pkr) FROM fact_daily_sales").fetchone()[0]
    con.close()
    print(f"loaded {n} lines, total PKR {total:,}")


if __name__ == "__main__":
    main()
