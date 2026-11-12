#!/usr/bin/env python3
"""P16 — build the deliverable workbook, the offline demo of exactly
what the Apps Script produces in Google Sheets:

  RawData       — all 295 sale lines (values, as pasted into Sheets)
  Weekly Report — 8 weeks x 6 categories, totals + WoW growth,
                  every number a live SUMIFS/COUNTIF formula
  Top Products  — top 5 by revenue, live SUMIF formulas
  Method        — what the script does, parameters, how to install

    python3 src/build_workbook.py
"""
import csv
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw" / "daily_sales_SAMPLE.csv"
SUMMARY = PROJECT / "outputs" / "weekly_summary.json"
OUT = PROJECT / "outputs" / "weekly_report_SAMPLE.xlsx"

CATEGORIES = ["Groceries", "Dairy", "Bakery", "Household",
              "Beverages", "Personal Care"]
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
MONEY_FMT = '#,##0'


def style_header(ws, ncols, row=1):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT


def main():
    rows = list(csv.DictReader(open(RAW, encoding="utf-8")))
    summary = json.loads(SUMMARY.read_text(encoding="utf-8"))
    n = len(rows)
    last = n + 1  # last RawData row (row 1 = header)

    wb = Workbook()

    # ── RawData ────────────────────────────────────────────────────────
    raw = wb.active
    raw.title = "RawData"
    raw.append(["date", "week_no", "week_start", "product", "category",
                "qty", "unit_price_pkr", "amount_pkr"])
    for r in rows:
        raw.append([r["date"], int(r["week_no"]), r["week_start"],
                    r["product"], r["category"], int(r["qty"]),
                    int(r["unit_price_pkr"]), int(r["amount_pkr"])])
    style_header(raw, 8)
    raw.freeze_panes = "A2"
    for col, w in zip("ABCDEFGH", [12, 9, 12, 22, 15, 7, 15, 13]):
        raw.column_dimensions[col].width = w
    for r in range(2, last + 1):
        raw.cell(row=r, column=6).number_format = MONEY_FMT
        raw.cell(row=r, column=7).number_format = MONEY_FMT
        raw.cell(row=r, column=8).number_format = MONEY_FMT

    # ── Weekly Report ──────────────────────────────────────────────────
    rep = wb.create_sheet("Weekly Report")
    rep["A1"] = "Weekly sales report — DailyNeeds Mart (SAMPLE)"
    rep["A1"].font = TITLE_FONT
    rep["A2"] = ("Every total below is a live formula over RawData — "
                 "the same SUMIFS logic the Apps Script writes in "
                 "Google Sheets. Paste new rows into RawData and the "
                 "report rebuilds itself.")
    rep["A2"].font = Font(italic=True, size=10, color="595959")
    headers = (["Week starting", "Lines"] + CATEGORIES +
               ["Total (PKR)", "WoW growth", "Top category"])
    rep.append([])  # row 3 blank
    for c, h in enumerate(headers, start=1):
        rep.cell(row=4, column=c, value=h)
    style_header(rep, len(headers), row=4)
    rep.freeze_panes = "A5"

    weeks = summary["weeks_detail"]
    for i, wk in enumerate(weeks):
        r = 5 + i
        rep.cell(row=r, column=1, value=wk["week_start"])
        rep.cell(row=r, column=2).value = (
            f"=COUNTIF(RawData!$C$2:$C${last},A{r})")
        for j, cat in enumerate(CATEGORIES):
            col = 3 + j
            rep.cell(row=r, column=col).value = (
                f"=SUMIFS(RawData!$H$2:$H${last},"
                f"RawData!$C$2:$C${last},A{r},"
                f"RawData!$E$2:$E${last},\"{cat}\")")
            rep.cell(row=r, column=col).number_format = MONEY_FMT
        rep.cell(row=r, column=9).value = f"=SUM(C{r}:H{r})"
        rep.cell(row=r, column=9).number_format = MONEY_FMT
        rep.cell(row=r, column=9).font = Font(bold=True)
        if i == 0:
            rep.cell(row=r, column=10, value="—")
        else:
            rep.cell(row=r, column=10).value = (
                f"=IF(I{r-1}=0,\"—\",I{r}/I{r-1}-1)")
            rep.cell(row=r, column=10).number_format = "0.00%"
        rep.cell(row=r, column=11).value = (
            f"=INDEX($C$4:$H$4,1,MATCH(MAX(C{r}:H{r}),C{r}:H{r},0))")
    t = 5 + len(weeks)
    rep.cell(row=t, column=1, value="TOTAL (all weeks)").font = Font(bold=True)
    rep.cell(row=t, column=2).value = f"=SUM(B5:B{t-1})"
    for j in range(6):
        col = 3 + j
        L = get_column_letter(col)
        rep.cell(row=t, column=col).value = f"=SUM({L}5:{L}{t-1})"
        rep.cell(row=t, column=col).number_format = MONEY_FMT
    rep.cell(row=t, column=9).value = f"=SUM(I5:I{t-1})"
    rep.cell(row=t, column=9).number_format = MONEY_FMT
    rep.cell(row=t, column=9).font = Font(bold=True)
    for col, w in zip(range(1, 12),
                      [14, 7, 12, 12, 10, 11, 11, 14, 13, 11, 15]):
        rep.column_dimensions[get_column_letter(col)].width = w

    chart = BarChart()
    chart.type = "col"
    chart.title = "Weekly revenue (PKR)"
    chart.y_axis.title = "PKR"
    chart.x_axis.title = "Week starting"
    chart.data = Reference(rep, min_col=9, min_row=4, max_row=t - 1)
    chart.categories = Reference(rep, min_col=1, min_row=5, max_row=t - 1)
    chart.height, chart.width = 7.5, 15
    rep.add_chart(chart, f"A{t + 2}")

    # ── Top Products ───────────────────────────────────────────────────
    top = wb.create_sheet("Top Products")
    top["A1"] = "Top 5 products by revenue — live SUMIF formulas"
    top["A1"].font = TITLE_FONT
    top.append([])
    for c, h in enumerate(["Product", "Category", "Units sold",
                           "Revenue (PKR)"], start=1):
        top.cell(row=3, column=c, value=h)
    style_header(top, 4, row=3)
    top.freeze_panes = "A4"
    for i, p in enumerate(summary["top_products"]):
        r = 4 + i
        top.cell(row=r, column=1, value=p["product"])
        top.cell(row=r, column=2, value=p["category"])
        top.cell(row=r, column=3).value = (
            f"=SUMIF(RawData!$D$2:$D${last},A{r},RawData!$F$2:$F${last})")
        top.cell(row=r, column=4).value = (
            f"=SUMIF(RawData!$D$2:$D${last},A{r},RawData!$H$2:$H${last})")
        top.cell(row=r, column=4).number_format = MONEY_FMT
    for col, w in zip("ABCD", [22, 15, 12, 15]):
        top.column_dimensions[col].width = w

    # ── Method ─────────────────────────────────────────────────────────
    m = wb.create_sheet("Method")
    m["A1"] = "How this report builds itself"
    m["A1"].font = TITLE_FONT
    lines = [
        ("What this is", "A Google Sheets automation pack. The Apps "
         "Script (apps_script/weekly_report_automation.gs) rebuilds "
         "the Weekly Report sheet from the RawData sheet — one row per "
         "week, one column per category, totals and week-on-week growth "
         "as live SUMIFS formulas."),
        ("Install", "Extensions > Apps Script, paste the .gs file, "
         "reload the sheet, then Weekly Report > Build weekly report. "
         "Install the Monday 08:00 trigger once and it rebuilds "
         "itself every week."),
        ("Email", "emailWeeklySummary() sends the latest week's "
         "headlines to the owner. Set OWNER_EMAIL in the CONFIG block; "
         "the first run asks for a one-time authorisation."),
        ("Data", "SAMPLE: DailyNeeds Mart, Lahore — 295 sale lines, "
         "8 weeks (2026-09-07 → 2026-11-01), PKR 570,445 total. "
         "Demo data, never claimed as real client work."),
        ("Check it", "Every number here is recomputed from the raw "
         "CSV by src/analyze.py and verified by src/checks.py — "
         "run it: python3 src/checks.py"),
        ("Same logic, three ways", "The weekly totals exist in three "
         "forms that must agree: the Apps Script SUMIFS, the SQL in "
         "sql/weekly_aggregation.sql, and this workbook's formulas. "
         "src/checks.py verifies all three against the raw file."),
    ]
    r = 3
    for title, text in lines:
        m.cell(row=r, column=1, value=title).font = Font(bold=True)
        m.cell(row=r, column=2, value=text)
        m.cell(row=r, column=2).alignment = Alignment(wrap_text=True,
                                                      vertical="top")
        r += 1
    m.column_dimensions["A"].width = 22
    m.column_dimensions["B"].width = 110

    wb.save(OUT)
    print(f"wrote {OUT.name}: sheets {wb.sheetnames}")


if __name__ == "__main__":
    main()
