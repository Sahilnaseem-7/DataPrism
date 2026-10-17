"""DataPrism local analytics MVP. Binds to loopback only; uploads stay on this machine."""
from datetime import datetime, date
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO, StringIO
from zipfile import ZipFile
from pathlib import Path
import csv
import json
import mimetypes
import re
import sqlite3
import sys
import urllib.parse
import uuid

ROOT = Path(__file__).resolve().parent
DB = ROOT / "data" / "margimap.sqlite3"
MAX_UPLOAD = 50 * 1024 * 1024

ALIASES = {
    "date": {"date", "order_date", "created_at", "invoicedate", "invoice_date", "timestamp", "day"},
    "order": {"order_id", "order", "invoice", "invoiceno", "invoice_no", "order_number", "transaction_id", "name"},
    "product": {"product", "product_name", "description", "item", "title", "sku", "stockcode", "stock_code"},
    "channel": {"channel", "source", "sales_channel", "platform", "store"},
    "country": {"country", "region", "shipping_country"},
    "qty": {"quantity", "qty", "units", "item_quantity"},
    "amount": {"revenue", "sales", "gross_sales", "net_sales", "amount", "total", "line_total", "subtotal", "total_sales"},
    "price": {"unit_price", "unitprice", "price", "item_price"},
    "refund": {"refund", "refunds", "return_amount", "refund_amount"},
    "cogs": {"cogs", "cost", "cost_of_goods", "product_cost"},
    "shipping": {"shipping", "shipping_cost", "shipping_fees"},
    "ad_spend": {"ad_spend", "spend", "advertising", "marketing_spend"},
    "customer": {"customer_id", "customer", "customerid", "buyer_id", "email_hash"},
    "campaign": {"campaign", "campaign_name", "ad_campaign"},
}

def connect():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("CREATE TABLE IF NOT EXISTS datasets (id INTEGER PRIMARY KEY, name TEXT NOT NULL, currency TEXT NOT NULL, created_at TEXT NOT NULL, records INTEGER NOT NULL, has_cogs INTEGER NOT NULL, fields TEXT NOT NULL, duplicates_removed INTEGER NOT NULL DEFAULT 0)")
    con.execute("CREATE TABLE IF NOT EXISTS sales (id INTEGER PRIMARY KEY, dataset_id INTEGER NOT NULL, day TEXT NOT NULL, order_id TEXT NOT NULL, product TEXT NOT NULL, channel TEXT NOT NULL, country TEXT NOT NULL, quantity REAL NOT NULL, sales REAL NOT NULL, refund REAL NOT NULL, cogs REAL NOT NULL, shipping REAL NOT NULL, ad_spend REAL NOT NULL, customer_id TEXT NOT NULL, campaign TEXT NOT NULL, FOREIGN KEY(dataset_id) REFERENCES datasets(id))")
    con.execute("CREATE INDEX IF NOT EXISTS ix_sales_dataset_day ON sales(dataset_id, day)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_sales_product ON sales(dataset_id, product)")
    con.execute("CREATE INDEX IF NOT EXISTS ix_sales_channel ON sales(dataset_id, channel)")
    con.commit()
    return con

def norm_header(x):
    return re.sub(r"[^a-z0-9]+", "_", str(x or "").strip().lower()).strip("_")

def parse_date(value):
    if value is None or str(value).strip() == "": return None
    if isinstance(value, (datetime, date)): return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    try: return datetime.fromisoformat(s.replace("Z", "+00:00")).strftime("%Y-%m-%d")
    except ValueError: pass
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%m/%d/%Y %H:%M", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try: return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError: continue
    return None

def number(value):
    if value is None: return 0.0
    s = str(value).strip().replace(",", "").replace("$", "").replace("£", "").replace("€", "")
    if s.startswith("(") and s.endswith(")"): s = "-" + s[1:-1]
    try: return float(s or 0)
    except ValueError: return 0.0

def decode_file(filename, blob):
    ext = Path(filename).suffix.lower()
    if ext == ".csv":
        text = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try: text = blob.decode(encoding); break
            except UnicodeDecodeError: continue
        reader = csv.reader(StringIO(text or ""))
        return list(reader)
    if ext in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError as exc:
            raise ValueError("XLSX support needs openpyxl. Install the app requirements and restart.") from exc
        wb = load_workbook(BytesIO(blob), read_only=True, data_only=True)
        ws = wb.active
        return [list(row) for row in ws.iter_rows(values_only=True)]
    raise ValueError("Choose a CSV or XLSX file.")

def map_headers(headers):
    cleaned = [norm_header(x) for x in headers]
    mapped = {}
    for field, names in ALIASES.items():
        for i, name in enumerate(cleaned):
            if name in names:
                mapped[field] = i
                break
    if "date" not in mapped: raise ValueError("Could not find a date column. Try date, order_date, InvoiceDate, or created_at.")
    if "amount" not in mapped and "price" not in mapped: raise ValueError("Could not find a sales amount or unit price column.")
    return mapped, cleaned

def transform(rows):
    if not rows or len(rows) < 2: raise ValueError("The file needs a header and at least one data row.")
    mapped, fields = map_headers(rows[0])
    out = []
    seen = set()
    duplicates_removed = 0
    has_cogs = "cogs" in mapped
    def get(row, field, fallback=None):
        i = mapped.get(field)
        return row[i] if i is not None and i < len(row) else fallback
    for n, row in enumerate(rows[1:], start=2):
        if not row or not any(str(v or "").strip() for v in row): continue
        signature = tuple(str(v or "").strip() for v in row)
        if signature in seen:
            duplicates_removed += 1
            continue
        seen.add(signature)
        day = parse_date(get(row, "date"))
        if not day: continue
        qty = number(get(row, "qty", 1))
        amount_raw = get(row, "amount")
        amount = number(amount_raw) if amount_raw not in (None, "") else qty * number(get(row, "price"))
        explicit_refund = max(0.0, number(get(row, "refund")))
        refund = max(explicit_refund, abs(amount) if amount < 0 else 0.0)
        if qty < 0 and amount > 0 and explicit_refund == 0: refund += amount; amount = 0
        sales = max(0.0, amount)
        if sales == 0 and refund == 0: continue
        order = str(get(row, "order", "") or "").strip()
        out.append((day, order or f"row-{n}-{uuid.uuid4().hex[:8]}", str(get(row, "product", "Unknown product") or "Unknown product").strip(), str(get(row, "channel", "Unknown") or "Unknown").strip(), str(get(row, "country", "Unknown") or "Unknown").strip(), qty, sales, refund, max(0, number(get(row, "cogs"))), max(0, number(get(row, "shipping"))), max(0, number(get(row, "ad_spend"))), str(get(row, "customer", "") or "").strip(), str(get(row, "campaign", "Unattributed") or "Unattributed").strip()))
    if not out: raise ValueError("No usable rows found. Include a valid date and a sales amount or unit price.")
    return out, has_cogs, fields, duplicates_removed

def filters_from(query, dataset_id):
    args = [dataset_id]
    where = ["dataset_id=?"]
    start, end = query.get("from", [""])[0], query.get("to", [""])[0]
    channel, product = query.get("channel", [""])[0], query.get("product", [""])[0]
    if start: where.append("day>=?"); args.append(start)
    if end: where.append("day<=?"); args.append(end)
    if channel: where.append("channel=?"); args.append(channel)
    if product: where.append("product=?"); args.append(product)
    return " AND ".join(where), args

def dataset_list(con):
    return [dict(r) for r in con.execute("SELECT id,name,currency,created_at,records,has_cogs,fields,duplicates_removed FROM datasets ORDER BY id DESC")]

def dashboard(con, dataset_id, query):
    meta = con.execute("SELECT * FROM datasets WHERE id=?", (dataset_id,)).fetchone()
    if not meta: return {"datasets": dataset_list(con), "dataset": None, "error": "Upload a CSV or XLSX file to get started."}
    where, args = filters_from(query, dataset_id)
    agg = con.execute(f"SELECT SUM(sales) gross, SUM(refund) refunds, SUM(quantity) units, SUM(cogs) cogs, SUM(shipping) shipping, SUM(ad_spend) ad, COUNT(DISTINCT CASE WHEN sales>0 THEN order_id END) orders, COUNT(DISTINCT customer_id) customers FROM sales WHERE {where}", args).fetchone()
    gross = agg["gross"] or 0; refunds = agg["refunds"] or 0; net = gross-refunds; order_count = agg["orders"] or 0
    products = [dict(r) for r in con.execute(f"SELECT product,SUM(sales)-SUM(refund) net_sales,SUM(quantity) units,COUNT(DISTINCT CASE WHEN sales>0 THEN order_id END) orders,SUM(refund) refunds,SUM(cogs) cogs,SUM(shipping) shipping,SUM(ad_spend) ad FROM sales WHERE {where} GROUP BY product ORDER BY net_sales DESC", args)]
    channels = [dict(r) for r in con.execute(f"SELECT channel,SUM(sales)-SUM(refund) net_sales,COUNT(DISTINCT CASE WHEN sales>0 THEN order_id END) orders,SUM(ad_spend) ad FROM sales WHERE {where} GROUP BY channel ORDER BY net_sales DESC", args)]
    countries = [dict(r) for r in con.execute(f"SELECT country,SUM(sales)-SUM(refund) net_sales FROM sales WHERE {where} GROUP BY country ORDER BY net_sales DESC LIMIT 10", args)]
    monthly = [dict(r) for r in con.execute(f"SELECT substr(day,1,7) month,SUM(sales)-SUM(refund) net_sales,COUNT(DISTINCT CASE WHEN sales>0 THEN order_id END) orders FROM sales WHERE {where} GROUP BY month ORDER BY month", args)]
    weekdays = [dict(r) for r in con.execute(f"SELECT strftime('%w',day) weekday,COUNT(DISTINCT CASE WHEN sales>0 THEN order_id END) orders,SUM(sales)-SUM(refund) net_sales FROM sales WHERE {where} GROUP BY weekday ORDER BY weekday", args)]
    customers = con.execute(f"SELECT COUNT(*) total,SUM(CASE WHEN orders>1 THEN 1 ELSE 0 END) repeat FROM (SELECT customer_id,COUNT(DISTINCT order_id) orders FROM sales WHERE {where} AND sales>0 AND customer_id!='' GROUP BY customer_id)", args).fetchone()
    customer_total = customers["total"] or 0; repeat = customers["repeat"] or 0
    rows = con.execute(f"SELECT * FROM sales WHERE {where} ORDER BY day DESC LIMIT 6000", args).fetchall()
    fields = sorted(set(meta["fields"].split(",")))
    return {
        "datasets": dataset_list(con), "dataset": dict(meta), "filters": {"from": query.get("from", [""])[0], "to": query.get("to", [""])[0], "channel": query.get("channel", [""])[0], "product": query.get("product", [""])[0]},
        "options": {"channels": [r[0] for r in con.execute("SELECT DISTINCT channel FROM sales WHERE dataset_id=? ORDER BY channel", (dataset_id,))], "products": [r[0] for r in con.execute("SELECT DISTINCT product FROM sales WHERE dataset_id=? ORDER BY product", (dataset_id,))]},
        "metrics": {"gross_sales":gross,"refunds":refunds,"net_sales":net,"orders":order_count,"aov":net/order_count if order_count else 0,"units":agg["units"] or 0,"identified_customers":customer_total,"repeat_customers":repeat,"repeat_pct":repeat/customer_total*100 if customer_total else 0,"estimated_margin":net-(agg["cogs"] or 0)-(agg["shipping"] or 0)-(agg["ad"] or 0),"margin_available":bool(meta["has_cogs"]),"ad_spend":agg["ad"] or 0,"roas":net/(agg["ad"] or 1) if (agg["ad"] or 0)>0 else None},
        "products":products,"channels":channels,"countries":countries,"monthly":monthly,"weekdays":weekdays,"fields":fields,
        "rows":[dict(r) for r in rows],"rows_truncated":con.execute(f"SELECT COUNT(*) FROM sales WHERE {where}",args).fetchone()[0]>6000
    }

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args): print(f"[{self.log_date_time_string()}] {fmt % args}")
    def send_json(self, data, status=200):
        body=json.dumps(data, ensure_ascii=False).encode(); self.send_response(status); self.send_header("Content-Type","application/json; charset=utf-8"); self.send_header("Content-Length",str(len(body))); self.send_header("Cache-Control","no-store"); self.end_headers(); self.wfile.write(body)
    def do_GET(self):
        parsed=urllib.parse.urlparse(self.path)
        if parsed.path=="/api/dashboard":
            query=urllib.parse.parse_qs(parsed.query); con=connect()
            try:
                did=int(query.get("dataset",[0])[0])
                if did==0:
                    latest=con.execute("SELECT id FROM datasets ORDER BY id DESC LIMIT 1").fetchone(); did=latest[0] if latest else 0
                self.send_json(dashboard(con,did,query))
            finally: con.close()
            return
        if parsed.path=="/api/export":
            query=urllib.parse.parse_qs(parsed.query); did=int(query.get("dataset",[0])[0]); con=connect()
            try:
                where,args=filters_from(query,did); data=con.execute(f"SELECT day,order_id,product,channel,country,quantity,sales,refund,cogs,shipping,ad_spend,customer_id,campaign FROM sales WHERE {where} ORDER BY day",args).fetchall()
                sio=StringIO(); writer=csv.writer(sio); writer.writerow(["date","order_id","product","channel","country","quantity","sales","refund","cogs","shipping","ad_spend","customer_id","campaign"]); writer.writerows([tuple(r) for r in data]); body=sio.getvalue().encode("utf-8-sig")
                self.send_response(200); self.send_header("Content-Type","text/csv; charset=utf-8"); self.send_header("Content-Disposition","attachment; filename=dataprism-export.csv"); self.send_header("Content-Length",str(len(body))); self.end_headers(); self.wfile.write(body)
            finally: con.close()
            return
        path="index.html" if parsed.path=="/" else parsed.path.lstrip("/")
        target=(ROOT/path).resolve()
        if ROOT not in target.parents and target!=ROOT: self.send_error(403); return
        if not target.is_file(): self.send_error(404); return
        body=target.read_bytes(); typ=mimetypes.guess_type(target.name)[0] or "application/octet-stream"; self.send_response(200); self.send_header("Content-Type",typ+ ("; charset=utf-8" if typ.startswith("text/") else "")); self.send_header("Content-Length",str(len(body))); self.end_headers(); self.wfile.write(body)
    def do_POST(self):
        request_path=urllib.parse.urlparse(self.path).path
        if request_path=="/api/load-sample":
            try:
                archive=ROOT.parent/"online-retail-uci.zip"
                if not archive.is_file(): raise ValueError("The included UCI sample archive is missing beside the app folder.")
                with ZipFile(archive) as zf: blob=zf.read("Online Retail.xlsx")
                rows=decode_file("Online Retail.xlsx",blob); records,has_cogs,fields,duplicates_removed=transform(rows)
                con=connect(); cur=con.execute("INSERT INTO datasets(name,currency,created_at,records,has_cogs,fields,duplicates_removed) VALUES(?,?,?,?,?,?,?)",("UCI Online Retail · 2010–2011","GBP",datetime.now().isoformat(timespec="seconds"),len(records),int(has_cogs),",".join(fields),duplicates_removed)); did=cur.lastrowid
                con.executemany("INSERT INTO sales(dataset_id,day,order_id,product,channel,country,quantity,sales,refund,cogs,shipping,ad_spend,customer_id,campaign) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",[(did,*r) for r in records]); con.commit(); con.close()
                self.send_json({"ok":True,"dataset_id":did,"records":len(records),"duplicates_removed":duplicates_removed,"name":"UCI Online Retail · 2010–2011"})
            except Exception as exc: self.send_json({"error":str(exc)},400)
            return
        if request_path!="/api/upload": self.send_error(404); return
        length=int(self.headers.get("Content-Length",0))
        if length<=0 or length>MAX_UPLOAD: self.send_json({"error":"File is empty or exceeds the 50 MB upload limit."},413); return
        body=self.rfile.read(length)
        try:
            envelope=(f"Content-Type: {self.headers.get('Content-Type')}\r\nMIME-Version: 1.0\r\n\r\n").encode()+body
            msg=BytesParser(policy=default).parsebytes(envelope)
            part=next(p for p in msg.iter_parts() if p.get_filename())
            filename=Path(part.get_filename()).name; blob=part.get_payload(decode=True)
            currency=urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query).get("currency",["USD"])[0]
            rows=decode_file(filename,blob); records,has_cogs,fields,duplicates_removed=transform(rows)
            con=connect(); cur=con.execute("INSERT INTO datasets(name,currency,created_at,records,has_cogs,fields,duplicates_removed) VALUES(?,?,?,?,?,?,?)",(filename,currency,datetime.now().isoformat(timespec="seconds"),len(records),int(has_cogs),",".join(fields),duplicates_removed)); did=cur.lastrowid
            con.executemany("INSERT INTO sales(dataset_id,day,order_id,product,channel,country,quantity,sales,refund,cogs,shipping,ad_spend,customer_id,campaign) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",[(did,*r) for r in records]); con.commit(); con.close()
            self.send_json({"ok":True,"dataset_id":did,"records":len(records),"duplicates_removed":duplicates_removed,"name":filename})
        except Exception as exc:
            self.send_json({"error":str(exc)},400)

if __name__=="__main__":
    port=int(sys.argv[1]) if len(sys.argv)>1 else 8766
    connect().close()
    print(f"DataPrism running at http://127.0.0.1:{port} — uploads are stored locally in {DB}")
    try: ThreadingHTTPServer(("127.0.0.1",port),Handler).serve_forever()
    except KeyboardInterrupt: print("\nDataPrism stopped.")
