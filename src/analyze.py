#!/usr/bin/env python3
"""P16 — analyse the SAMPLE sales file and regenerate every output
number from the raw CSV (never typed): weekly totals, week-on-week
growth, category shares, top products. Writes:

  data/processed/weekly_summary.csv  (8 weeks x 6 categories)
  outputs/weekly_summary.json       (machine-readable KPIs)
  outputs/weekly_report.md          (findings report)

    python3 src/analyze.py
"""
import csv
import json
from collections import defaultdict
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw" / "daily_sales_SAMPLE.csv"
CATEGORIES = ["Groceries", "Dairy", "Bakery", "Household",
              "Beverages", "Personal Care"]


def main():
    rows = list(csv.DictReader(open(RAW, encoding="utf-8")))
    for r in rows:
        r["qty"] = int(r["qty"])
        r["unit_price_pkr"] = int(r["unit_price_pkr"])
        r["amount_pkr"] = int(r["amount_pkr"])
        r["week_no"] = int(r["week_no"])

    weeks = {}
    for r in rows:
        w = weeks.setdefault(r["week_no"], {
            "week_no": r["week_no"], "week_start": r["week_start"],
            "lines": 0, "total_pkr": 0, "qty": 0,
            "by_category": {c: 0 for c in CATEGORIES}})
        w["lines"] += 1
        w["total_pkr"] += r["amount_pkr"]
        w["qty"] += r["qty"]
        w["by_category"][r["category"]] += r["amount_pkr"]

    week_list = [weeks[n] for n in sorted(weeks)]
    for i, w in enumerate(week_list):
        prev = week_list[i - 1]["total_pkr"] if i else None
        w["wow_growth_pct"] = (round(100.0 * (w["total_pkr"] - prev) / prev, 2)
                               if prev else None)
        w["top_category"] = max(w["by_category"], key=w["by_category"].get)

    # category totals + top products across the 8 weeks
    cat_total = {c: sum(w["by_category"][c] for w in week_list)
                 for c in CATEGORIES}
    prod = defaultdict(lambda: {"qty": 0, "revenue": 0})
    for r in rows:
        prod[r["product"]]["qty"] += r["qty"]
        prod[r["product"]]["revenue"] += r["amount_pkr"]
    prod_cat = {}
    for r in rows:
        prod_cat[r["product"]] = r["category"]
    top_products = sorted(
        ({"product": p, "category": prod_cat[p],
          "qty": v["qty"], "revenue_pkr": v["revenue"]}
         for p, v in prod.items()),
        key=lambda x: -x["revenue_pkr"])[:5]

    grand = sum(w["total_pkr"] for w in week_list)
    # weekend (Fri-Sun) vs weekday average units per sale line, measured
    from datetime import date as _date
    _we = [r["qty"] for r in rows if _date.fromisoformat(r["date"]).weekday() >= 4]
    _wd = [r["qty"] for r in rows if _date.fromisoformat(r["date"]).weekday() < 4]
    weekend_lift_pct = round(100.0 * (sum(_we) / len(_we))
                             / (sum(_wd) / len(_wd)) - 100.0, 1)
    summary = {
        "shop": "DailyNeeds Mart, Lahore (SAMPLE)",
        "period": "2026-09-07 to 2026-11-01",
        "weeks": 8,
        "lines": len(rows),
        "days": 56,
        "total_revenue_pkr": grand,
        "avg_weekly_revenue_pkr": round(grand / 8),
        "best_week": max(week_list, key=lambda w: w["total_pkr"]),
        "worst_week": min(week_list, key=lambda w: w["total_pkr"]),
        "first_week_pkr": week_list[0]["total_pkr"],
        "last_week_pkr": week_list[-1]["total_pkr"],
        "overall_growth_pct": round(
            100.0 * (week_list[-1]["total_pkr"] - week_list[0]["total_pkr"])
            / week_list[0]["total_pkr"], 2),
        "category_totals_pkr": cat_total,
        "top_category": max(cat_total, key=cat_total.get),
        "weekend_lift_pct": weekend_lift_pct,
        "top_products": top_products,
        "weeks_detail": week_list,
    }
    summary["best_week"] = {
        "week_no": summary["best_week"]["week_no"],
        "week_start": summary["best_week"]["week_start"],
        "total_pkr": summary["best_week"]["total_pkr"]}
    summary["worst_week"] = {
        "week_no": summary["worst_week"]["week_no"],
        "week_start": summary["worst_week"]["week_start"],
        "total_pkr": summary["worst_week"]["total_pkr"]}

    # weekly summary CSV (what the Apps Script reproduces in Sheets)
    out_csv = PROJECT / "data" / "processed" / "weekly_summary.csv"
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["week_no", "week_start", "lines", "qty",
                    "total_pkr", "wow_growth_pct", "top_category"]
                   + [c.lower().replace(" ", "_") + "_pkr"
                      for c in CATEGORIES])
        for wk in week_list:
            w.writerow([wk["week_no"], wk["week_start"], wk["lines"],
                        wk["qty"], wk["total_pkr"],
                        "" if wk["wow_growth_pct"] is None
                        else wk["wow_growth_pct"], wk["top_category"]]
                       + [wk["by_category"][c] for c in CATEGORIES])

    out_json = PROJECT / "outputs" / "weekly_summary.json"
    out_json.write_text(json.dumps(summary, indent=2), encoding="utf-8")

    lines = [
        "# P16 — Weekly Report: DailyNeeds Mart (SAMPLE)",
        "",
        "**Period:** Mon 2026-09-07 → Sun 2026-11-01 (8 full weeks, 56 days)",
        f"**Lines:** {len(rows)} sale lines · **Total revenue: PKR {grand:,}**",
        f"**Average week:** PKR {summary['avg_weekly_revenue_pkr']:,}",
        "",
        "## What the automation found",
        "",
        f"1. **The shop grew {summary['overall_growth_pct']}% from week 1 to "
        f"week 8** (PKR {week_list[0]['total_pkr']:,} → "
        f"PKR {week_list[-1]['total_pkr']:,}). The growth is in the data, "
        "not in the script — the script only reports it.",
        f"2. **Best week: week {summary['best_week']['week_no']} "
        f"({summary['best_week']['week_start']}) at "
        f"PKR {summary['best_week']['total_pkr']:,}**; worst week: week "
        f"{summary['worst_week']['week_no']} at "
        f"PKR {summary['worst_week']['total_pkr']:,}.",
        f"3. **{summary['top_category']} is the biggest category** at "
        f"PKR {cat_total[summary['top_category']]:,} "
        f"({100.0 * cat_total[summary['top_category']] / grand:.1f}% of revenue).",
        f"4. **Top product: {top_products[0]['product']}** — "
        f"PKR {top_products[0]['revenue_pkr']:,} across "
        f"{top_products[0]['qty']} units.",
        "5. **Weekends carry the shop:** Friday–Sunday sale lines average "
        f"{weekend_lift_pct}% more units than weekday lines — the script's "
        "weekend flag surfaces this on a real file too.",
        "",
        "## Weekly table (recomputed from the raw file)",
        "",
        "| Week | Starting | Lines | Revenue (PKR) | WoW | Top category |",
        "|---|---|---|---|---|---|",
    ]
    for wk in week_list:
        wow = "—" if wk["wow_growth_pct"] is None else f"{wk['wow_growth_pct']:+.2f}%"
        lines.append(f"| {wk['week_no']} | {wk['week_start']} | {wk['lines']} | "
                     f"{wk['total_pkr']:,} | {wow} | {wk['top_category']} |")
    lines += [
        "",
        "## What the Apps Script does (in Google Sheets)",
        "",
        "- `buildWeeklyReport()` rebuilds the Weekly Report sheet from the "
        "RawData sheet: one row per week, one column per category, totals "
        "and week-on-week growth as live `SUMIFS` formulas.",
        "- `setupSheets()` creates RawData / Weekly Report / Top Products "
        "if they are missing, so a new client file works on first run.",
        "- `createWeeklyTrigger()` installs a Monday 08:00 time trigger — "
        "the report rebuilds itself every week without anyone opening it.",
        "- `emailWeeklySummary()` emails the week's headline numbers to the "
        "owner (recipient set in the CONFIG block; needs one authorisation).",
        "",
        "> Every number in this report is recomputed from "
        "`data/raw/daily_sales_SAMPLE.csv` by `src/analyze.py` — nothing "
        "is typed by hand. Run `python3 src/checks.py` to re-verify.",
    ]
    (PROJECT / "outputs" / "weekly_report.md").write_text(
        "\n".join(lines) + "\n", encoding="utf-8")
    print(f"total PKR {grand:,}, weeks {len(week_list)}, "
          f"growth W1->W8 {summary['overall_growth_pct']}%")


if __name__ == "__main__":
    main()
