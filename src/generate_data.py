#!/usr/bin/env python3
"""P16 — generate the SAMPLE daily sales file for DailyNeeds Mart.

Fictional Lahore general store, 8 full weeks: Mon 2026-09-07 to
Sun 2026-11-01 (56 days). Seeded, so the file is reproducible.
Each day gets 4-7 sale lines; weekend days (Fri-Sun) sell more,
and the shop grows ~2% a week — the automation has a real trend
to find. Prices are fixed per product; amount = qty x unit price.

    python3 src/generate_data.py
"""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
OUT = PROJECT / "data" / "raw" / "daily_sales_SAMPLE.csv"

SEED = 20261112

# product -> (category, unit_price_pkr)
PRODUCTS = {
    "Basmati Rice 1kg": ("Groceries", 420),
    "Atta 10kg": ("Groceries", 1150),
    "Cooking Oil 1L": ("Groceries", 640),
    "Sugar 1kg": ("Groceries", 185),
    "Tea 500g": ("Groceries", 950),
    "Lentils (Daal) 1kg": ("Groceries", 320),
    "Fresh Milk 1L": ("Dairy", 240),
    "Yogurt 500g": ("Dairy", 180),
    "Desi Ghee 1kg": ("Dairy", 1450),
    "Cheese 200g": ("Dairy", 520),
    "White Bread": ("Bakery", 120),
    "Rusk 400g": ("Bakery", 260),
    "Cake Slice": ("Bakery", 150),
    "Biscuits 250g": ("Bakery", 210),
    "Surf Excel 1kg": ("Household", 580),
    "Dishwash Bar": ("Household", 90),
    "Tissue Box": ("Household", 160),
    "Floor Cleaner 1L": ("Household", 340),
    "Rooh Afza 800ml": ("Beverages", 420),
    "Soft Drink 1.5L": ("Beverages", 180),
    "Juice 1L": ("Beverages", 260),
    "Mineral Water 1.5L": ("Beverages", 110),
    "Soap Bar": ("Personal Care", 95),
    "Shampoo 200ml": ("Personal Care", 380),
    "Toothpaste 100g": ("Personal Care", 240),
}

START = date(2026, 9, 7)   # Monday
DAYS = 56                   # exactly 8 weeks


def main():
    rng = random.Random(SEED)
    rows = []
    names = list(PRODUCTS)
    for day_no in range(DAYS):
        day = START + timedelta(days=day_no)
        week_no = day_no // 7 + 1
        week_start = START + timedelta(days=(week_no - 1) * 7)
        # weekend lift (Fri=4, Sat=5, Sun=6) + slow weekly growth
        weekend = 1.35 if day.weekday() >= 4 else 1.0
        growth = 1.0 + 0.02 * (week_no - 1)
        n_lines = rng.randint(4, 7)
        for _ in range(n_lines):
            product = rng.choice(names)
            category, price = PRODUCTS[product]
            qty = max(1, round(rng.randint(1, 8) * weekend * growth))
            rows.append({
                "date": day.isoformat(),
                "week_no": week_no,
                "week_start": week_start.isoformat(),
                "product": product,
                "category": category,
                "qty": qty,
                "unit_price_pkr": price,
                "amount_pkr": qty * price,
            })
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    total = sum(r["amount_pkr"] for r in rows)
    print(f"wrote {len(rows)} lines, {DAYS} days, 8 weeks, total PKR {total:,}")


if __name__ == "__main__":
    main()
