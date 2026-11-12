# P16 — Weekly Report: DailyNeeds Mart (SAMPLE)

**Period:** Mon 2026-09-07 → Sun 2026-11-01 (8 full weeks, 56 days)
**Lines:** 295 sale lines · **Total revenue: PKR 570,445**
**Average week:** PKR 71,306

## What the automation found

1. **The shop grew 4.12% from week 1 to week 8** (PKR 61,465 → PKR 64,000). The growth is in the data, not in the script — the script only reports it.
2. **Best week: week 5 (2026-10-05) at PKR 86,875**; worst week: week 1 at PKR 61,465.
3. **Groceries is the biggest category** at PKR 178,235 (31.2% of revenue).
4. **Top product: Desi Ghee 1kg** — PKR 97,150 across 67 units.
5. **Weekends carry the shop:** Friday–Sunday sale lines average 36.0% more units than weekday lines — the script's weekend flag surfaces this on a real file too.

## Weekly table (recomputed from the raw file)

| Week | Starting | Lines | Revenue (PKR) | WoW | Top category |
|---|---|---|---|---|---|
| 1 | 2026-09-07 | 33 | 61,465 | — | Groceries |
| 2 | 2026-09-14 | 39 | 64,975 | +5.71% | Dairy |
| 3 | 2026-09-21 | 39 | 71,045 | +9.34% | Groceries |
| 4 | 2026-09-28 | 40 | 66,345 | -6.62% | Groceries |
| 5 | 2026-10-05 | 40 | 86,875 | +30.94% | Groceries |
| 6 | 2026-10-12 | 36 | 75,215 | -13.42% | Dairy |
| 7 | 2026-10-19 | 32 | 80,525 | +7.06% | Dairy |
| 8 | 2026-10-26 | 36 | 64,000 | -20.52% | Dairy |

## What the Apps Script does (in Google Sheets)

- `buildWeeklyReport()` rebuilds the Weekly Report sheet from the RawData sheet: one row per week, one column per category, totals and week-on-week growth as live `SUMIFS` formulas.
- `setupSheets()` creates RawData / Weekly Report / Top Products if they are missing, so a new client file works on first run.
- `createWeeklyTrigger()` installs a Monday 08:00 time trigger — the report rebuilds itself every week without anyone opening it.
- `emailWeeklySummary()` emails the week's headline numbers to the owner (recipient set in the CONFIG block; needs one authorisation).

> Every number in this report is recomputed from `data/raw/daily_sales_SAMPLE.csv` by `src/analyze.py` — nothing is typed by hand. Run `python3 src/checks.py` to re-verify.
