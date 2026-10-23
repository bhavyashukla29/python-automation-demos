# Website → spreadsheet

![The scraper running, then the workbook it writes](media/demo.gif)

**Problem.** Someone needs every product from a supplier's or competitor's site in a spreadsheet:
name, category, price, rating, stock. Copying a thousand listings by hand takes days and goes stale
the moment prices change.

**Solution.** `scrape_books.py` walks every category of the catalogue, follows the pagination, reads
each product card, and writes a workbook that is ready to hand to someone else: formatted,
filterable, and summarised.

**Result.** 1,000 products from 50 categories in 58 seconds, on
[books.toscrape.com](https://books.toscrape.com), a public site built for scraping practice.

## What you get

`books_catalogue.xlsx`, with three sheets:

| Sheet | Contents |
|---|---|
| **Summary** | Products, average price and rating, cheapest and priciest per category, plus a chart of the 15 largest categories |
| **Products** | All 1,000 rows: title, category, price, stars, stock, and a link to the product page. Prices are colour-coded from green (cheap) to red (expensive), with filters and a frozen header |
| **Best deals** | The 75 in-stock products rated 4–5 stars and priced under £20, best first |

## How it works

- **Crawls by category.** It reads the category list from the sidebar and then follows each
  category's "next" links, so nothing is missed and every row knows its category.
- **Is polite and resilient.** It waits between requests, retries failed requests with a back-off,
  and sends an identifying user agent.
- **Writes Excel properly.** It uses `openpyxl` for number formats, conditional colour scales,
  hyperlinks, auto-filters, frozen panes and a native Excel chart. These are real spreadsheet
  features, not just colours pasted into cells.

## Run it

```bash
pip install requests beautifulsoup4 openpyxl
python scrape_books.py                     # full catalogue -> books_catalogue.xlsx
python scrape_books.py --max-categories 3  # quick test run
python scrape_books.py --deal-price 15     # change the "best deals" price cut-off
```

## Adapting it to a real site

The structure stays the same and only the selectors change. Typical requests:

- Competitor price tracking, scheduled to run daily, with week-on-week price changes highlighted
- Supplier catalogues pulled into your own product sheet
- Directory or listing data collected into a lead list

I check a site's terms and `robots.txt` before scraping it. Pages that load their content with
JavaScript need a browser-based scraper (Playwright) instead of `requests`.
