# Messy exports → one clean report

![The script's fix log, then the report it builds](media/demo.gif)

**Problem.** A small brand sells on its website, on Amazon and in a shop. Each channel exports sales
in its own format, and none of them are clean. Building one monthly report means hours of
copy-pasting, retyping product names and hunting for duplicates, and the numbers are still hard to
trust.

**Solution.** `build_report.py` reads all three exports, fixes every problem it finds, merges them
into one dataset, and writes a formatted Excel report. It also writes a log of every fix it made.

**Result.** 3,186 raw rows become 2,961 clean sale lines (₹25.8 lakh in revenue), with 9 kinds of
problem fixed and every fix counted.

> The data is synthetic (`make_messy_data.py`), modelled on a small Indian tea and coffee brand.
> The problems in it are the ones real exports have.

## What it fixes

| Source | Problem | Rows | What the script does |
|---|---|---|---|
| Shopify | Duplicate rows from overlapping exports | 60 | Removed |
| Shopify | Refunded orders | 45 | Excluded from revenue |
| Amazon | Cancelled orders | 37 | Excluded from revenue |
| Amazon | Misspelt / inconsistent product names (`Filter Cofee`, `Masala Chai Blend - 250 g`) | 271 | Standardised using the SKU |
| Amazon | Prices stored as text (`"₹1,299.00"`) | 894 | Converted to numbers |
| Store till | Blank rows | 12 | Removed |
| Store till | Dates in 3 different formats | 403 | Unified |
| Store till | Abbreviated / lowercase item names (`masala chai 250`, `Cups (4)`) | 601 | Matched to the price list |
| Store till | Missing sale amounts | 27 | Recovered from price list × quantity |

## What you get

`output/sales_report.xlsx`:

| Sheet | Contents |
|---|---|
| **Summary** | Revenue, orders, units and average order value; monthly revenue by channel as a table and a stacked chart; top products with their share of revenue |
| **By product** | Product × month revenue |
| **Clean data** | Every sale line in one consistent format, filterable |
| **Data quality log** | Each problem found, how many rows it affected, and what was done |

Plus `output/clean_sales.csv`, which feeds the [dashboard demo](../03-live-dashboard).

## How it works

- **One loader per channel**, each converting its export to the same schema (date, channel, order,
  SKU, quantity, revenue). Adding a fourth channel means adding one function.
- **A price list as the single source of truth.** Amazon rows are standardised by SKU. Free-text
  till names are matched by comparing words, after normalising units (`250g` → `250`) and
  quantities (`x10` → `10`). Anything it can't match confidently is held back for review instead
  of being guessed.
- **Nothing is fixed silently.** Every fix is counted in the data quality log, so the person reading
  the report can see exactly what changed.

## Run it

```bash
pip install pandas openpyxl
python make_messy_data.py   # writes the three sample exports into input/
python build_report.py      # writes output/sales_report.xlsx and output/clean_sales.csv
```

## Adapting it to your business

Point the loaders at your real exports (Shopify, Amazon, Flipkart, Tally, a till system or a Google
Sheet) and swap in your price list. It can then run on a schedule, so the report is ready before
anyone opens their laptop.
