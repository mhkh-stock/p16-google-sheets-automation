-- P16 — weekly aggregation: the same report the Apps Script builds
-- with SUMIFS, expressed in SQL. One row per week, one column per
-- category, total and week-on-week growth via LAG().
-- Verified against data/processed/weekly_summary.csv by src/checks.py.

SELECT
    week_start,
    week_no,
    COUNT(*) AS lines,
    SUM(qty) AS units,
    SUM(CASE WHEN p.category = 'Groceries'     THEN amount_pkr ELSE 0 END) AS groceries_pkr,
    SUM(CASE WHEN p.category = 'Dairy'         THEN amount_pkr ELSE 0 END) AS dairy_pkr,
    SUM(CASE WHEN p.category = 'Bakery'        THEN amount_pkr ELSE 0 END) AS bakery_pkr,
    SUM(CASE WHEN p.category = 'Household'     THEN amount_pkr ELSE 0 END) AS household_pkr,
    SUM(CASE WHEN p.category = 'Beverages'     THEN amount_pkr ELSE 0 END) AS beverages_pkr,
    SUM(CASE WHEN p.category = 'Personal Care' THEN amount_pkr ELSE 0 END) AS personal_care_pkr,
    SUM(amount_pkr) AS total_pkr,
    ROUND(100.0 * (SUM(amount_pkr)
        - LAG(SUM(amount_pkr)) OVER (ORDER BY week_start))
        / LAG(SUM(amount_pkr)) OVER (ORDER BY week_start), 2) AS wow_growth_pct
FROM fact_daily_sales f
JOIN dim_product p ON p.product = f.product
GROUP BY week_start, week_no
ORDER BY week_start;
