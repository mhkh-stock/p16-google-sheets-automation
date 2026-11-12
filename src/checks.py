#!/usr/bin/env python3
"""Checks as code for P16. Recomputes every headline number from the
raw CSV and verifies the workbook, the SQL aggregation and the Apps
Script against it — three independent paths that must agree.

    python3 src/checks.py
Exit code 0 when every check passes, 1 otherwise.
"""
import csv
import json
import sqlite3
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw" / "daily_sales_SAMPLE.csv"
SUMMARY = PROJECT / "outputs" / "weekly_summary.json"
WEEKLY_CSV = PROJECT / "data" / "processed" / "weekly_summary.csv"
WORKBOOK = PROJECT / "outputs" / "weekly_report_SAMPLE.xlsx"
GS = PROJECT / "apps_script" / "weekly_report_automation.gs"
SQL = PROJECT / "sql" / "weekly_aggregation.sql"
SCHEMA = PROJECT / "sql" / "schema.sql"

CATEGORIES = ["Groceries", "Dairy", "Bakery", "Household",
              "Beverages", "Personal Care"]
fails = []
ran = []


def check(name, cond, detail=""):
    ran.append(name)
    print(("PASS " if cond else "FAIL ") + name +
          (f" — {detail}" if detail else ""))
    if not cond:
        fails.append(name)


rows = list(csv.DictReader(open(RAW, encoding="utf-8")))
for r in rows:
    r["qty"] = int(r["qty"])
    r["amount_pkr"] = int(r["amount_pkr"])
    r["week_no"] = int(r["week_no"])

# ── raw file ─────────────────────────────────────────────────────────────
check("row count = 295", len(rows) == 295, str(len(rows)))
check("dates span 2026-09-07..2026-11-01",
      min(r["date"] for r in rows) == "2026-09-07" and
      max(r["date"] for r in rows) == "2026-11-01")
check("56 distinct days", len({r["date"] for r in rows}) == 56)
check("8 distinct weeks", len({r["week_no"] for r in rows}) == 8)
check("week_no matches week_start",
      all(r["week_start"] == {
          1: "2026-09-07", 2: "2026-09-14", 3: "2026-09-21",
          4: "2026-09-28", 5: "2026-10-05", 6: "2026-10-12",
          7: "2026-10-19", 8: "2026-10-26"}[r["week_no"]] for r in rows))
raw2 = list(csv.DictReader(open(RAW, encoding="utf-8")))
check("amount_pkr = qty * unit_price_pkr (re-read)",
      all(int(r["amount_pkr"]) == int(r["qty"]) * int(r["unit_price_pkr"])
          for r in raw2))
check("6 categories only",
      {r["category"] for r in rows} == set(CATEGORIES))
check("25 distinct products", len({r["product"] for r in rows}) == 25)
total = sum(r["amount_pkr"] for r in rows)
check("grand total = 570,445", total == 570445, f"{total:,}")

# ── weekly recompute ─────────────────────────────────────────────────────
weeks = {}
for r in rows:
    w = weeks.setdefault(r["week_no"], {"total": 0, "lines": 0,
                                        "cat": defaultdict(int)})
    w["total"] += r["amount_pkr"]
    w["lines"] += 1
    w["cat"][r["category"]] += r["amount_pkr"]
wk = [weeks[n] for n in sorted(weeks)]
check("week 1 total = 61,465", wk[0]["total"] == 61465, str(wk[0]["total"]))
check("week 5 total = 86,875 (best)", wk[4]["total"] == 86875,
      str(wk[4]["total"]))
check("week 8 total = 64,000", wk[7]["total"] == 64000, str(wk[7]["total"]))
check("weekly totals sum to grand",
      sum(w["total"] for w in wk) == total)
check("category totals sum to grand",
      sum(sum(w["cat"].values()) for w in wk) == total)
cat_total = {c: sum(w["cat"][c] for w in wk) for c in CATEGORIES}
check("Groceries = 178,235", cat_total["Groceries"] == 178235)
check("Dairy = 164,130", cat_total["Dairy"] == 164130)
check("top category = Groceries",
      max(cat_total, key=cat_total.get) == "Groceries")
prod = defaultdict(int)
for r in rows:
    prod[r["product"]] += r["amount_pkr"]
top5 = sorted(prod, key=prod.get, reverse=True)[:5]
check("top product = Desi Ghee 1kg", top5[0] == "Desi Ghee 1kg")
check("Desi Ghee revenue = 97,150", prod["Desi Ghee 1kg"] == 97150)
we = [r["qty"] for r in rows if date.fromisoformat(r["date"]).weekday() >= 4]
wd = [r["qty"] for r in rows if date.fromisoformat(r["date"]).weekday() < 4]
lift = round(100.0 * (sum(we) / len(we)) / (sum(wd) / len(wd)) - 100.0, 1)
check("weekend lift = 36.0%", lift == 36.0, f"{lift}%")

# ── summary JSON + weekly CSV ────────────────────────────────────────────
s = json.loads(SUMMARY.read_text(encoding="utf-8"))
check("JSON total matches recompute", s["total_revenue_pkr"] == total)
check("JSON lines = 295", s["lines"] == 295)
check("JSON best week = 5 at 86,875",
      s["best_week"] == {"week_no": 5, "week_start": "2026-10-05",
                         "total_pkr": 86875}, str(s["best_week"]))
check("JSON worst week = 1 at 61,465",
      s["worst_week"] == {"week_no": 1, "week_start": "2026-09-07",
                          "total_pkr": 61465})
check("JSON growth W1->W8 = 4.12%",
      s["overall_growth_pct"] == 4.12, str(s["overall_growth_pct"]))
check("JSON weekend lift = 36.0%", s["weekend_lift_pct"] == 36.0)
check("JSON top products match recompute",
      [p["product"] for p in s["top_products"]] == top5)
check("JSON top product revenue = 97,150",
      s["top_products"][0]["revenue_pkr"] == 97150)
wcsv = list(csv.DictReader(open(WEEKLY_CSV, encoding="utf-8")))
check("weekly CSV has 8 rows", len(wcsv) == 8)
check("weekly CSV totals match recompute",
      all(int(wcsv[i]["total_pkr"]) == wk[i]["total"] for i in range(8)))
check("weekly CSV week 5 WoW present",
      wcsv[4]["wow_growth_pct"] != "")

# ── SQL aggregation (in-memory, same schema + same CSV) ──────────────────
con = sqlite3.connect(":memory:")
con.executescript(SCHEMA.read_text())
con.executemany("INSERT INTO dim_product VALUES (?, ?)",
                sorted({(r["product"], r["category"]) for r in rows}))
price_of = {x["product"]: int(x["unit_price_pkr"]) for x in raw2}
con.executemany(
    "INSERT INTO fact_daily_sales "
    "(sale_date, week_no, week_start, product, qty, unit_price_pkr, "
    " amount_pkr) VALUES (?,?,?,?,?,?,?)",
    [(r["date"], r["week_no"], r["week_start"], r["product"], r["qty"],
      price_of[r["product"]], r["amount_pkr"]) for r in rows])
sql_rows = con.execute(SQL.read_text()).fetchall()
check("SQL returns 8 weeks", len(sql_rows) == 8, str(len(sql_rows)))
check("SQL week totals match recompute",
      all(sql_rows[i][10] == wk[i]["total"] for i in range(8)),
      str([r[10] for r in sql_rows]))
check("SQL groceries week 1 matches",
      sql_rows[0][4] == wk[0]["cat"]["Groceries"],
      str(sql_rows[0][4]))
con.close()

# ── workbook ─────────────────────────────────────────────────────────────
wb = load_workbook(WORKBOOK)
check("sheets = RawData, Weekly Report, Top Products, Method",
      wb.sheetnames == ["RawData", "Weekly Report", "Top Products",
                        "Method"], str(wb.sheetnames))
check("RawData rows = 296 incl header",
      wb["RawData"].max_row == 296, str(wb["RawData"].max_row))
check("RawData total column sums to 570,445 (values)",
      sum(wb["RawData"].cell(row=r, column=8).value
          for r in range(2, 297)) == 570445)
rep = wb["Weekly Report"]
check("Weekly Report has 8 week rows + total row",
      rep.cell(row=13, column=1).value == "TOTAL (all weeks)",
      str(rep.cell(row=13, column=1).value))
check("Weekly Report B5 is a COUNTIF formula",
      isinstance(rep["B5"].value, str) and
      rep["B5"].value.startswith("=COUNTIF("), str(rep["B5"].value))
check("Weekly Report C5 is a SUMIFS formula",
      isinstance(rep["C5"].value, str) and
      rep["C5"].value.startswith("=SUMIFS("), str(rep["C5"].value))
check("Weekly Report I5 sums C5:H5",
      rep["I5"].value == "=SUM(C5:H5)", str(rep["I5"].value))
check("Weekly Report J6 is a WoW formula",
      isinstance(rep["J6"].value, str) and "I6/I5-1" in rep["J6"].value,
      str(rep["J6"].value))
check("Weekly Report J5 is the week-1 dash",
      rep["J5"].value == "—", str(rep["J5"].value))
check("Weekly Report has a chart", len(rep._charts) >= 1)
top = wb["Top Products"]
check("Top Products row 4 = Desi Ghee 1kg",
      top["A4"].value == "Desi Ghee 1kg", str(top["A4"].value))
check("Top Products D4 is a SUMIF formula",
      isinstance(top["D4"].value, str) and
      top["D4"].value.startswith("=SUMIF("), str(top["D4"].value))
check("Top Products has 5 products", top.max_row == 8,
      str(top.max_row))
check("Method sheet mentions Apps Script",
      "Apps Script" in (wb["Method"]["B3"].value or ""))

# ── Apps Script ──────────────────────────────────────────────────────────
gs = GS.read_text(encoding="utf-8")
for fn in ["onOpen", "setupSheets", "buildWeeklyReport",
           "buildTopProducts", "emailWeeklySummary",
           "createWeeklyTrigger", "deleteWeeklyTrigger"]:
    check(f".gs defines {fn}()", f"function {fn}(" in gs)
check(".gs uses SUMIFS", "SUMIFS" in gs)
check(".gs uses COUNTIF", "COUNTIF" in gs)
check(".gs installs a Monday trigger",
      "onWeekDay(ScriptApp.WeekDay.MONDAY)" in gs)
check(".gs emails via MailApp", "MailApp.sendEmail" in gs)
check(".gs lists all 6 categories",
      all(f"'{c}'" in gs for c in CATEGORIES))
check(".gs has an OWNER_EMAIL placeholder",
      "OWNER_EMAIL" in gs and "owner@example.com" in gs)

print(f"\n{len(ran) - len(fails)}/{len(ran)} checks passed")
if fails:
    print("FAILURES:", fails)
    sys.exit(1)
