import http.server, threading, sys, os, csv, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import export_catalog as ec
from openpyxl import load_workbook

def card(i, price="£10.%02d", cls="Three", stock="In stock"):
    return f'<article class="product_pod"><p class="star-rating {cls}"></p><h3><a href="item{i}.html" title="Book {i} &amp; more">Book...</a></h3><div><p class="price_color">{price % i if "%" in price else price}</p><p class="availability">{stock}</p></div></article>'
PAGES = {
 "/robots.txt": "User-agent: *\nDisallow: /private/",
 "/p1.html": "<html>" + card(1) + card(2, cls="Five") + card(3, stock="Out of stock") + '<article class="product_pod"><h3>broken card</h3></article><li class="next"><a href="p2.html">next</a></li></html>',
 "/p2.html": "<html>" + card(2) + card(4, price="€7,50") + "</html>",
}
class H(http.server.BaseHTTPRequestHandler):
    def do_GET(s):
        b = PAGES.get(s.path)
        if b is None: s.send_response(404); s.end_headers(); return
        s.send_response(200); s.end_headers(); s.wfile.write(b.encode())
    def log_message(*a): pass
srv = http.server.HTTPServer(("127.0.0.1", 0), H); threading.Thread(target=srv.serve_forever, daemon=True).start()
B = f"http://127.0.0.1:{srv.server_port}"
ok = []
def t(n, f):
    try: f(); ok.append(1); print("PASS", n)
    except Exception as e: ok.append(0); print("FAIL", n, repr(e))
d = tempfile.mkdtemp(); out = os.path.join(d, "o")
ec.main(["--url", B + "/p1.html", "--pages", "5", "--out", out, "--delay", "0"])
rows = list(csv.DictReader(open(out + ".csv", encoding="utf-8-sig")))
t("follows pagination and de-duplicates (4 unique)", lambda: (len(rows) == 4) or (_ for _ in ()).throw(AssertionError(len(rows))))
t("skips broken card without crashing", lambda: all(r["title"].startswith("Book") for r in rows))
t("title entity decoded", lambda: rows[0]["title"] == "Book 1 & more" or (_ for _ in ()).throw(AssertionError(rows[0]["title"])))
t("price parsed, euro comma handled", lambda: [r["price"] for r in rows] == ["10.01", "10.02", "10.03", "7.5"] and rows[3]["currency"] == "€" or (_ for _ in ()).throw(AssertionError([r["price"] for r in rows])))
t("pound sign decoded when server sends no charset", lambda: rows[0]["currency"] == "£" or (_ for _ in ()).throw(AssertionError(rows[0]["currency"])))
t("rating and stock", lambda: rows[1]["rating"] == "5" and rows[2]["in_stock"] == "False" and rows[0]["in_stock"] == "True")
t("absolute urls", lambda: rows[0]["url"] == B + "/item1.html" or (_ for _ in ()).throw(AssertionError(rows[0]["url"])))
wb = load_workbook(out + ".xlsx"); ws = wb["Catalog"]
t("xlsx header, 4 rows, numeric price", lambda: ws.max_row == 5 and ws["A1"].value == "title" and isinstance(ws["B2"].value, float))
t("--pages 1 reads one page", lambda: (ec.main(["--url", B + "/p1.html", "--pages", "1", "--out", out + "1", "--format", "csv", "--delay", "0"]), len(list(csv.DictReader(open(out + "1.csv", encoding="utf-8-sig")))) == 3 or (_ for _ in ()).throw(AssertionError("count"))))
def exits(args, msg):
    try: ec.main(args)
    except SystemExit as e: assert msg in str(e) or e.code == 2, str(e); return
    raise AssertionError("no exit")
t("bad --pages rejected", lambda: exits(["--url", B, "--pages", "0"], ""))
t("non-http url rejected", lambda: exits(["--url", "file:///etc/passwd"], ""))
t("page with no products exits with message", lambda: exits(["--url", B + "/robots.txt", "--delay", "0"], "No products"))
PAGES["/robots.txt"] = "User-agent: *\nDisallow: /"
t("robots.txt disallow respected", lambda: exits(["--url", B + "/p1.html"], "robots.txt"))
# second site with different markup, driven by a config file
import json
PAGES["/s1.html"] = '<html><div class="product-card"><h2 class="name"><a href="/x/a.html">Lamp</a></h2><span class="price">$24.90</span><span class="rating">4.5</span><span class="stock">Available</span></div><div class="product-card"><h2 class="name"><a href="/x/b.html">Desk</a></h2><span class="price">$120.00</span><span class="rating">3</span><span class="stock">Sold out</span></div><a class="next-page" href="s2.html">more</a></html>'
PAGES["/s2.html"] = '<html><div class="product-card"><h2 class="name"><a href="/x/c.html">Chair</a></h2><span class="price">$60.00</span><span class="stock">Available</span></div></html>'
PAGES["/robots.txt"] = "User-agent: *\nDisallow: /private/"
cfg = os.path.join(d, "sel.json"); json.dump({"item": "div.product-card", "title": "h2.name a", "link": "h2.name a", "price": "span.price", "rating": "span.rating", "stock": "span.stock", "next": "a.next-page", "in_stock_text": "available"}, open(cfg, "w"))
o2 = os.path.join(d, "second"); ec.main(["--url", B + "/s1.html", "--config", cfg, "--out", o2, "--format", "csv", "--delay", "0"])
r2 = list(csv.DictReader(open(o2 + ".csv", encoding="utf-8-sig")))
t("config: second site parsed across 2 pages", lambda: [r["title"] for r in r2] == ["Lamp", "Desk", "Chair"] or (_ for _ in ()).throw(AssertionError(r2)))
t("config: numeric rating, price, currency, stock", lambda: r2[0]["rating"] == "4.5" and r2[1]["rating"] == "3.0" and r2[0]["price"] == "24.9" and r2[0]["currency"] == "$" and r2[0]["in_stock"] == "True" and r2[1]["in_stock"] == "False" and r2[2]["rating"] == "")
t("config: absolute url from root-relative link", lambda: r2[0]["url"] == B + "/x/a.html")
bad = os.path.join(d, "bad.json"); json.dump({"itme": "x"}, open(bad, "w"))
t("config: unknown key rejected", lambda: exits(["--url", B + "/s1.html", "--config", bad], "unknown config keys"))
t("default selectors still work after change", lambda: len(rows) == 4)
print({"passed": sum(ok), "total": len(ok)})
