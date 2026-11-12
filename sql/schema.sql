-- P16 — star-lite schema for the weekly-report automation.
-- Grain: one row per sale line. The weekly report is a view over
-- this table — the same aggregation the Apps Script does with
-- SUMIFS, so the SQL and the script can be checked against each
-- other. Loaded by src/build_sqlite.py.

CREATE TABLE IF NOT EXISTS dim_product (
    product       TEXT PRIMARY KEY,
    category      TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS fact_daily_sales (
    sale_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    sale_date      TEXT NOT NULL,   -- yyyy-mm-dd
    week_no        INTEGER NOT NULL,
    week_start     TEXT NOT NULL,   -- Monday, yyyy-mm-dd
    product        TEXT NOT NULL REFERENCES dim_product(product),
    qty            INTEGER NOT NULL CHECK (qty > 0),
    unit_price_pkr INTEGER NOT NULL CHECK (unit_price_pkr > 0),
    amount_pkr     INTEGER NOT NULL CHECK (amount_pkr = qty * unit_price_pkr)
);

CREATE INDEX IF NOT EXISTS idx_sales_week
    ON fact_daily_sales (week_start, product);
