# Catalog Export Tool (Python)

Reads a paginated product listing and writes a clean CSV and/or Excel file.
Demo target: https://books.toscrape.com/ (a public practice site made for scraping tests).

    pip install -r requirements.txt
    python export_catalog.py --url https://books.toscrape.com/ --pages 3 --out books --format both

What it does: follows "next" links, de-duplicates, parses price/currency/rating/stock, skips broken cards and counts them,
retries failed requests, waits between requests (--delay), checks robots.txt first, limits pages (1-50), refuses non-http URLs.
CSV is UTF-8 with BOM so Excel opens it correctly. The Excel file has a styled header, frozen row, filter and numeric prices.

Tests: `python tests/test_export.py` (18 checks, local fixture server, no internet needed).
Live run checked once on 24 Nov 2026: 3 pages, 60 unique items, 0 skipped (sample-output.csv / .xlsx).
Other sites: pass `--config my-site.json` with CSS selectors (item, title, link, price, rating, stock, next, in_stock_text). See `selectors.example.json`. Unknown keys are rejected. This is tested against a local fixture site with different markup, not against a real second website, so check the first page of any new site before trusting a full run.
Limits: static HTML pages only (no JavaScript-rendered shops, no logins). Respect each site's terms and robots.txt.
