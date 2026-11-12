# P16 — Google Sheets Automation Pack: the weekly report that builds itself

**Client (SAMPLE):** DailyNeeds Mart, Lahore — general store,
8 weeks (Mon 2026-09-07 → Sun 2026-11-01) · **Status: SAMPLE delivery on DEMO data**
**Tools:** Google Apps Script · Google Sheets (SUMIFS) · SQL (SQLite) · Excel (live formulas)
**Deliverable:** `apps_script/weekly_report_automation.gs` + `outputs/weekly_report_SAMPLE.xlsx`

> ⚠️ **What this is and is not.** The sales file is a fictional
> shop's demo data, generated to show the method. No real shop,
> customer or payment is involved. What is real is the automation:
> the same script, the same SUMIFS logic and the same workbook run
> on a real client's file the same way, and every number in the
> report is recomputed from the data, never typed.

## Business question it answers

> "Stop re-typing the same report every week — build it for me."

A shop owner who rebuilds the same weekly table by hand every
Monday wastes an hour and mistypes a total sooner or later. This
delivery drops one script into their Google Sheet: the report
rebuilds itself from the raw sales rows, emails the owner the
headlines, and a Monday 08:00 trigger means nobody has to remember
to run it at all.

## What is inside

| File | What it does |
|---|---|
| `apps_script/weekly_report_automation.gs` | The deliverable: `onOpen` menu, `buildWeeklyReport()` (SUMIFS weekly table), `buildTopProducts()`, `emailWeeklySummary()`, `createWeeklyTrigger()` (Monday 08:00), `setupSheets()` |
| `outputs/weekly_report_SAMPLE.xlsx` | Offline demo of the exact same output: RawData (295 lines) + Weekly Report (8 weeks, live formulas) + Top Products + Method + weekly revenue chart |
| `outputs/weekly_report.md` | Findings report: headline numbers, weekly table, what the script does |
| `outputs/weekly_summary.json` | Machine-readable summary (KPIs + weekly detail + top products) |
| `data/raw/daily_sales_SAMPLE.csv` | The sales file: 295 lines, 56 days, 8 weeks, PKR 570,445 |
| `data/processed/weekly_summary.csv` | The 8-week summary the script reproduces in Sheets |
| `data/processed/sales.db` | SQLite: dim_product + fact_daily_sales (NOT uploaded — rebuildable) |
| `data/README.md` | Data dictionary + known quirks |
| `sql/schema.sql` | Star-lite schema for the two tables |
| `sql/weekly_aggregation.sql` | The weekly report in SQL — must agree with the script and the workbook |
| `src/generate_data.py` | Makes the demo file (seeded; NOT run on a real client file) |
| `src/analyze.py` | Recomputes every headline number from the raw CSV |
| `src/build_sqlite.py` | Loads the sales file into SQLite |
| `src/build_workbook.py` | Builds the workbook: 4 sheets, live formulas |
| `src/checks.py` | Data-quality + reconciliation checks as code — 61 checks, all passing |
| `requirements.txt` | Pinned: `openpyxl==3.1.2` (Python 3.12) |

## How it fits together

```
client's daily sales file ──▶ Apps Script (in Sheets) ──▶ Weekly Report sheet
        (295 lines)                 │                        (live SUMIFS)
                                    ▼
                          src/build_sqlite.py ──▶ data/processed/sales.db
                                    │              (NOT uploaded: rebuildable)
                                    ▼
                          src/analyze.py ──▶ outputs/weekly_report.md
                                    │              + weekly_summary.json/csv
                                    ▼
                          src/build_workbook.py ──▶ outputs/weekly_report_SAMPLE.xlsx
                                    ▲ src/checks.py verifies all of it (61 checks)
```

## Headline numbers (SAMPLE — DailyNeeds Mart, 8 weeks)

- **Total revenue: PKR 570,445** across 295 sale lines, 56 days
- **Average week: PKR 71,306** · best week 5 (from 2026-10-05): PKR 86,875 · worst week 1: PKR 61,465
- **Growth week 1 → week 8: +4.12%** (PKR 61,465 → PKR 64,000)
- **Groceries is the biggest category:** PKR 178,235 (31.2% of revenue); Dairy close behind at PKR 164,130
- **Top product: Desi Ghee 1kg** — PKR 97,150 across 67 units
- **Weekends carry the shop:** Friday–Sunday sale lines average 36.0% more units than weekday lines

## How the automation works

1. **Paste** the daily sales into the `RawData` sheet (same 8 columns).
2. **Run** Weekly Report → Build weekly report. The script finds the
   distinct week-start Mondays in the data and writes one row per
   week: line count (`COUNTIF`), one `SUMIFS` column per category,
   a Total column and a week-on-week growth column — all live
   formulas, so the sheet stays correct when new rows are pasted in.
3. **Trigger** (once): Weekly Report → Install Monday 08:00 trigger.
   The report rebuilds itself every Monday morning.
4. **Email** (optional): set `OWNER_EMAIL` in the CONFIG block and
   run Weekly Report → Email this week's summary. First run asks for
   a one-time authorisation; after that the owner gets the week's
   headlines without opening the sheet.

## Same logic, three ways — and they must agree

The weekly totals exist in three forms: the Apps Script `SUMIFS`,
the SQL in `sql/weekly_aggregation.sql`, and this workbook's
formulas. `src/checks.py` runs the SQL against the same CSV and
compares all three — a mistake in any one of them shows up as a
failure. The workbook's formulas were also recalculated in
LibreOffice: every weekly total matches the recompute exactly.

## Install (for the client, 2 minutes)

Extensions → Apps Script → paste `weekly_report_automation.gs` →
Save → reload the sheet → Weekly Report menu → Build weekly report.
