#!/usr/bin/env python3
"""Catalog scraper: reads a paginated product listing and writes CSV and/or Excel.

Demo target: https://books.toscrape.com/ (a public practice site made for scraping tests).
Only use it on sites whose terms and robots.txt allow it.

Usage:
  python export_catalog.py --url https://books.toscrape.com/ --pages 3 --out books --format both
"""
import argparse, csv, sys, time, re
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser
import requests
from bs4 import BeautifulSoup

UA = "catalog-export-demo/1.0 (portfolio demo)"
RATING = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
FIELDS = ["title", "price", "currency", "rating", "in_stock", "url"]


def allowed(url, session):
    p = urlparse(url)
    rp = RobotFileParser()
    try:
        r = session.get(f"{p.scheme}://{p.netloc}/robots.txt", timeout=10)
        if r.status_code != 200:
            return True  # no robots file found
        rp.parse(r.text.splitlines())
    except requests.RequestException:
        return True
    return rp.can_fetch(UA, url)


def fetch(session, url, retries=3, delay=1.0):
    for i in range(retries):
        try:
            r = session.get(url, timeout=15)
            if r.status_code == 404:
                return None
            r.raise_for_status()
            return r.content.decode(r.encoding if r.encoding and 'charset' in r.headers.get('content-type', '').lower() else 'utf-8', errors='replace')
        except requests.RequestException as e:
            if i == retries - 1:
                raise SystemExit(f"Could not fetch {url}: {e}")
            time.sleep(delay * (i + 1))


def parse_price(text):
    m = re.search(r"([^\d\s.,]*)\s*([\d]+(?:[.,]\d+)?)", text or "")
    if not m:
        return None, ""
    return float(m.group(2).replace(",", ".")), m.group(1)


def parse_page(html, base):
    soup = BeautifulSoup(html, "html.parser")
    rows, bad = [], 0
    for card in soup.select("article.product_pod"):
        a = card.select_one("h3 a")
        price_el = card.select_one(".price_color")
        if not a or not price_el:
            bad += 1
            continue
        price, cur = parse_price(price_el.get_text())
        if price is None:
            bad += 1
            continue
        rating_el = card.select_one("p.star-rating")
        rating = next((RATING[c] for c in (rating_el.get("class", []) if rating_el else []) if c in RATING), None)
        stock = card.select_one(".availability")
        rows.append({
            "title": a.get("title") or a.get_text(strip=True),
            "price": price, "currency": cur, "rating": rating,
            "in_stock": "in stock" in (stock.get_text(strip=True).lower() if stock else ""),
            "url": urljoin(base, a.get("href", "")),
        })
    nxt = soup.select_one("li.next a")
    return rows, (urljoin(base, nxt["href"]) if nxt else None), bad


def write_csv(rows, path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig so Excel opens it correctly
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)


def write_xlsx(rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook(); ws = wb.active; ws.title = "Catalog"
    ws.append(FIELDS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="0F766E")
    for r in rows:
        ws.append([r[k] for k in FIELDS])
    for col, w in zip("ABCDEF", (48, 10, 10, 8, 10, 70)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "0.00"
    wb.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--url", required=True)
    ap.add_argument("--pages", type=int, default=3, help="max pages to read (1-50)")
    ap.add_argument("--out", default="catalog")
    ap.add_argument("--format", choices=["csv", "xlsx", "both"], default="both")
    ap.add_argument("--delay", type=float, default=1.0, help="seconds between requests")
    ap.add_argument("--ignore-robots", action="store_true", help="not recommended")
    a = ap.parse_args(argv)
    if not (1 <= a.pages <= 50):
        ap.error("--pages must be between 1 and 50")
    if urlparse(a.url).scheme not in ("http", "https"):
        ap.error("--url must start with http:// or https://")
    s = requests.Session(); s.headers["User-Agent"] = UA
    if not a.ignore_robots and not allowed(a.url, s):
        raise SystemExit("robots.txt does not allow this URL. Stopping.")
    rows, url, seen, skipped = [], a.url, set(), 0
    cur_url = a.url
    for n in range(a.pages):
        if not url or url in seen:
            break
        seen.add(url); cur_url = url
        html = fetch(s, url, delay=a.delay)
        if html is None:
            break
        page_rows, url, bad = parse_page(html, cur_url)
        rows += page_rows; skipped += bad
        print(f"page {n + 1}: {len(page_rows)} items" + (f", {bad} skipped" if bad else ""))
        if url:
            time.sleep(a.delay)
    if not rows:
        raise SystemExit("No products found. Check the URL or the page structure.")
    seen_u, uniq = set(), []
    for r in rows:
        if r["url"] not in seen_u:
            seen_u.add(r["url"]); uniq.append(r)
    if a.format in ("csv", "both"):
        write_csv(uniq, a.out + ".csv")
    if a.format in ("xlsx", "both"):
        write_xlsx(uniq, a.out + ".xlsx")
    print(f"done: {len(uniq)} unique items, {skipped} skipped, written to {a.out}.{{{a.format if a.format != 'both' else 'csv,xlsx'}}}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
