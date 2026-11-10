"""Generate messy monthly sales exports, like the ones a small business gets from its POS system."""
import csv
import random
from datetime import date, timedelta
from pathlib import Path

random.seed(7)
OUT = Path(__file__).parent / "input"
OUT.mkdir(exist_ok=True)

BRANCHES = ["Ankara Kizilay", "Ankara Cankaya", "Istanbul Kadikoy", "Izmir Alsancak"]
PRODUCTS = {"Coffee Beans 1kg": 420.0, "Espresso Machine": 8900.0, "Grinder": 2150.0,
            "Milk Frother": 990.0, "Paper Filters": 85.0, "Travel Mug": 340.0}
DATE_FORMATS = ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"]  # different exports use different formats

receipt_no = 100000
for month in (8, 9, 10):
    rows = []
    day = date(2026, month, 1)
    while day.month == month:
        for _ in range(random.randint(6, 14)):
            product = random.choice(list(PRODUCTS))
            qty = random.randint(1, 5)
            price = PRODUCTS[product]
            branch = random.choice(BRANCHES)
            # messiness: random casing/whitespace, comma decimals, missing qty, duplicates
            if random.random() < 0.15:
                branch = f"  {branch.upper()} "
            price_txt = f"{price:.2f}".replace(".", ",") if random.random() < 0.3 else f"{price:.2f}"
            qty_txt = "" if random.random() < 0.02 else str(qty)
            receipt_no += 1
            row = [f"R{receipt_no}", day.strftime(random.choice(DATE_FORMATS)), branch, product,
                   qty_txt, price_txt]
            rows.append(row)
            if random.random() < 0.03:
                rows.append(row)  # accidental duplicate export line
        day += timedelta(days=1)
    with open(OUT / f"sales_2026_{month:02d}.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Receipt No", "Date", "Branch", "Product", "Qty", "Unit Price"])
        w.writerows(rows)
print("sample files written to", OUT)
