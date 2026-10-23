"""Generate realistic, deliberately messy sales exports for the demo.

All data is synthetic. It imitates a small Indian tea & coffee brand selling on
three channels, each exporting in its own inconsistent format - the kind of
files a business owner actually receives:

  * shopify_orders.csv - ISO timestamps, overlapping exports (duplicate rows), refunds
  * amazon_orders.csv  - DD/MM/YYYY dates, "₹1,299.00" price strings, typos, cancellations
  * store_pos.csv      - mixed date formats, abbreviated/lowercase item names, blank rows

Usage:
    py -3.10 make_messy_data.py            # writes into ./input
"""

import csv
import random
from datetime import date, datetime, timedelta
from pathlib import Path

random.seed(42)

START, END = date(2026, 5, 1), date(2026, 10, 31)

# canonical name -> (sku, price in ₹, popularity weight)
CATALOGUE = {
    "Masala Chai Blend 250g": ("TEA-MC-250", 349, 30),
    "Assam Gold Tea 500g": ("TEA-AG-500", 449, 18),
    "Kashmiri Kahwa 100g": ("TEA-KK-100", 399, 10),
    "Filter Coffee Powder 500g": ("COF-FC-500", 599, 20),
    "Cold Brew Coffee Bags (10)": ("COF-CB-010", 499, 14),
    "French Press 600ml": ("EQP-FP-600", 1299, 5),
    "Ceramic Chai Cups (Set of 4)": ("EQP-CC-004", 899, 7),
    "Tea Lovers Gift Hamper": ("GFT-TL-001", 1799, 4),
}

# How each channel mangles product names
AMAZON_VARIANTS = {
    "Masala Chai Blend 250g": ["Masala Chai Blend 250g", "Masala Chai Blend - 250 g", "Masala Chay Blend 250g"],
    "Filter Coffee Powder 500g": ["Filter Coffee Powder 500g", "Filter Cofee Powder 500g"],
    "Tea Lovers Gift Hamper": ["Tea Lovers Gift Hamper", "Tea Lover's Gift Hamper"],
}
POS_NAMES = {
    "Masala Chai Blend 250g": ["masala chai 250", "MASALA CHAI 250G", "Masala chai blend"],
    "Assam Gold Tea 500g": ["assam gold 500", "Assam Gold"],
    "Kashmiri Kahwa 100g": ["kahwa 100g", "Kashmiri Kahwa"],
    "Filter Coffee Powder 500g": ["filter coffee 500", "FILTER COFFEE"],
    "Cold Brew Coffee Bags (10)": ["cold brew bags", "Cold Brew x10"],
    "French Press 600ml": ["french press", "French Press 600"],
    "Ceramic Chai Cups (Set of 4)": ["chai cups set", "Cups (4)"],
    "Tea Lovers Gift Hamper": ["gift hamper", "Hamper - Tea Lovers"],
}


def season_weight(d: date, product: str) -> float:
    """Steady growth, a monsoon chai bump, and pre-Diwali gifting in October."""
    growth = 1 + (d - START).days / 365
    w = growth
    if d.month in (7, 8) and "Chai" in product:
        w *= 1.4
    if d.month == 10 and product in ("Tea Lovers Gift Hamper", "Ceramic Chai Cups (Set of 4)", "French Press 600ml"):
        w *= 3.0
    if d.month in (5, 6) and "Cold Brew" in product:
        w *= 1.8
    return w


def pick_product(d: date) -> str:
    names = list(CATALOGUE)
    weights = [CATALOGUE[n][2] * season_weight(d, n) for n in names]
    return random.choices(names, weights)[0]


def days():
    d = START
    while d <= END:
        yield d
        d += timedelta(days=1)


def make_shopify(out: Path) -> None:
    rows, order_no = [], 1001
    for d in days():
        for _ in range(random.randint(4, 9) + (3 if d.month == 10 else 0)):
            ts = datetime.combine(d, datetime.min.time()) + timedelta(minutes=random.randint(420, 1380))
            status = "refunded" if random.random() < 0.03 else "paid"
            # Shopify merges repeats of one product into a single line item
            for p in sorted({pick_product(d) for _ in range(random.choice([1, 1, 1, 2]))}):
                rows.append({
                    "Name": f"#{order_no}",
                    "Created at": ts.strftime("%Y-%m-%d %H:%M:%S +0530"),
                    "Financial Status": status,
                    "Lineitem name": p,
                    "Lineitem sku": CATALOGUE[p][0],
                    "Lineitem quantity": random.choice([1, 1, 1, 2, 3]),
                    "Lineitem price": f"{CATALOGUE[p][1]:.2f}",
                })
            order_no += 1
    # Two exports were pasted together, so a chunk of orders appears twice
    overlap = rows[len(rows) // 2: len(rows) // 2 + 60]
    rows = rows[: len(rows) // 2 + 60] + overlap + rows[len(rows) // 2 + 60:]
    write_csv(out / "shopify_orders.csv", rows)


def make_amazon(out: Path) -> None:
    rows = []
    for d in days():
        for _ in range(random.randint(3, 7)):
            p = pick_product(d)
            qty = random.choice([1, 1, 2])
            name = random.choice(AMAZON_VARIANTS.get(p, [p]))
            rows.append({
                "order-id": f"403-{random.randint(1000000, 9999999)}-{random.randint(1000000, 9999999)}",
                "purchase-date": d.strftime("%d/%m/%Y"),
                "order-status": "Cancelled" if random.random() < 0.04 else "Shipped",
                "sku": CATALOGUE[p][0],
                "product-name": name,
                "quantity-purchased": qty,
                "item-price": f"₹{CATALOGUE[p][1] * qty:,.2f}",  # line total, as Amazon reports it
            })
    write_csv(out / "amazon_orders.csv", rows)


def make_pos(out: Path) -> None:
    rows = []
    formats = ["%d %b %Y", "%Y/%m/%d", "%d-%m-%Y"]
    for d in days():
        if d.weekday() == 0:  # store closed Mondays
            continue
        for _ in range(random.randint(2, 6)):
            p = pick_product(d)
            qty = random.choice([1, 1, 2])
            amount = CATALOGUE[p][1] * qty
            rows.append({
                "Date": d.strftime(random.choice(formats)),
                "Item": "  " + random.choice(POS_NAMES[p]) + " " * random.randint(0, 2),
                "Qty": qty,
                # cashier sometimes skips the amount; it has to be recovered from the price list
                "Amount": "" if random.random() < 0.05 else amount,
            })
        if random.random() < 0.08:
            rows.append({"Date": "", "Item": "", "Qty": "", "Amount": ""})
    write_csv(out / "store_pos.csv", rows)


def write_csv(path: Path, rows: list[dict]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {path.name:<20} {len(rows):>5} rows")


if __name__ == "__main__":
    out = Path(__file__).parent / "input"
    out.mkdir(exist_ok=True)
    make_shopify(out)
    make_amazon(out)
    make_pos(out)
