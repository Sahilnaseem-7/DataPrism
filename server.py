"""DataPrism Local Web App Server (Single-Workspace E-Commerce Server)."""
from __future__ import annotations

import csv
import email.parser
import email.policy
import io
import json
import os
import re
import sqlite3
import sys
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

try:
    import openpyxl
except ImportError:
    openpyxl = None

ROOT = Path(__file__).resolve().parent
DB_PATH = Path(os.environ.get("DATAPRISM_LOCAL_DB", str(ROOT / "data" / "margimap.sqlite3")))

ALIASES = {
    "date": ["date", "day", "order_date", "transaction_date", "created_at", "timestamp"],
    "order_id": ["order_id", "order", "id", "transaction_id", "trans_id", "invoice_id"],
    "product": ["product", "item", "product_name", "sku", "title"],
    "channel": ["channel", "source", "platform", "store", "medium"],
    "country": ["country", "region", "market", "location"],
    "qty": ["qty", "quantity", "units", "count"],
    "amount": ["amount", "sales", "revenue", "net_sales", "total", "gross_sales"],
    "price": ["price", "unit_price", "rate"],
    "refund": ["refund", "refunds", "returns", "return_amount"],
    "cogs": ["cogs", "cost", "unit_cost", "cost_of_goods"],
    "shipping": ["shipping", "freight", "delivery_fee"],
    "ad_spend": ["ad_spend", "marketing", "ads", "spend"],
    "customer": ["customer", "customer_name", "client", "buyer", "email", "customer_id"],
    "campaign": ["campaign", "utm_campaign", "promo"],
}


def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_PATH), timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA journal_mode=WAL")
    con.executescript(
        """
        CREATE TABLE IF NOT EXISTS datasets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            filename TEXT NOT NULL,
            row_count INTEGER NOT NULL DEFAULT 0,
            is_active INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dataset_id INTEGER NOT NULL REFERENCES datasets(id) ON DELETE CASCADE,
            day TEXT,
            order_id TEXT,
            product TEXT,
            channel TEXT,
            country TEXT,
            qty REAL DEFAULT 1,
            amount REAL DEFAULT 0,
            price REAL DEFAULT 0,
            refund REAL DEFAULT 0,
            cogs REAL DEFAULT 0,
            shipping REAL DEFAULT 0,
            ad_spend REAL DEFAULT 0,
            customer TEXT,
            campaign TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_sales_dataset ON sales(dataset_id);
        """
    )
    return con


def to_float(val) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "").replace("$", "").replace("€", "").replace("£", "")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except ValueError:
        return 0.0


def norm_date(val) -> str:
    if not val:
        return ""
    if hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s[:10], fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s[:10]


def parse_rows(filename: str, raw: bytes) -> list[dict]:
    ext = Path(filename).suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        if openpyxl is None:
            raise ValueError("openpyxl is required to import Excel files.")
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        ws = wb.active
        matrix = [list(row) for row in ws.iter_rows(values_only=True)]
        if len(matrix) < 2:
            return []
        headers = [str(h or f"col_{i}").strip() for i, h in enumerate(matrix[0])]
        return [
            {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}
            for row in matrix[1:]
            if any(c is not None and str(c).strip() != "" for c in row)
        ]
    text = raw.decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(text))
    return [dict(r) for r in reader if any(v and str(v).strip() for v in r.values())]


def map_column(headers: list[str], target: str) -> str | None:
    cand = ALIASES.get(target, [target])
    norm_map = {re.sub(r"[^a-z0-9]+", "_", h.lower()).strip("_"): h for h in headers}
    for c in cand:
        if c in norm_map:
            return norm_map[c]
    return None


def parse_multipart_form(content_type: str, body: bytes) -> tuple[dict[str, str], tuple[str, bytes] | None]:
    fields: dict[str, str] = {}
    file_item: tuple[str, bytes] | None = None
    header_bytes = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
    msg = email.parser.BytesParser(policy=email.policy.default).parsebytes(header_bytes + body)
    if not msg.is_multipart():
        return fields, file_item
    for part in msg.iter_parts():
        disp = part.get("Content-Disposition", "")
        name_m = re.search(r'name="([^"]+)"', disp)
        if not name_m:
            continue
        field_name = name_m.group(1)
        fn_m = re.search(r'filename="([^"]*)"', disp)
        payload = part.get_payload(decode=True) or b""
        if fn_m:
            file_item = (fn_m.group(1) or "upload.csv", payload)
        else:
            fields[field_name] = payload.decode("utf-8", errors="replace").strip()
    return fields, file_item


class LocalHandler(BaseHTTPRequestHandler):
    def _json(self, payload: dict, status: int = 200) -> None:
        raw = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/health":
            self._json({"ok": True, "service": "dataprism-local"})
            return
        if parsed.path == "/api/dashboard":
            qs = parse_qs(parsed.query)
            rng = qs.get("range", ["all"])[0]
            with connect() as con:
                datasets = [dict(r) for r in con.execute("SELECT * FROM datasets ORDER BY id DESC").fetchall()]
                active = next((d for d in datasets if d["is_active"]), datasets[0] if datasets else None)
                if not active:
                    self._json({"datasets": [], "dataset": None, "metrics": {}, "monthly": [], "products": [], "channels": [], "customers": [], "rows": []})
                    return
                ds_id = active["id"]
                where = "WHERE dataset_id = ?"
                params: list = [ds_id]
                if rng in ("30", "90"):
                    max_day_row = con.execute("SELECT MAX(day) AS max_day FROM sales WHERE dataset_id = ?", (ds_id,)).fetchone()
                    if max_day_row and max_day_row["max_day"]:
                        try:
                            cutoff = (datetime.strptime(max_day_row["max_day"], "%Y-%m-%d") - timedelta(days=int(rng))).strftime("%Y-%m-%d")
                            where += " AND day >= ?"
                            params.append(cutoff)
                        except ValueError:
                            pass
                m_row = con.execute(
                    f"SELECT COALESCE(SUM(amount - refund), 0) AS total_sales, COUNT(*) AS order_count, COUNT(DISTINCT customer) AS customers FROM sales {where}",
                    params,
                ).fetchone()
                total_sales = float(m_row["total_sales"] or 0)
                order_count = int(m_row["order_count"] or 0)
                aov = total_sales / order_count if order_count else 0.0
                monthly = [
                    dict(r)
                    for r in con.execute(
                        f"SELECT substr(day, 1, 7) AS month, SUM(amount - refund) AS sales FROM sales {where} AND day != '' GROUP BY substr(day, 1, 7) ORDER BY month",
                        params,
                    ).fetchall()
                ]
                products = [
                    dict(r)
                    for r in con.execute(
                        f"SELECT COALESCE(NULLIF(product, ''), 'Unknown') AS Product, COUNT(*) AS orders, SUM(amount - refund) AS net_sales FROM sales {where} GROUP BY Product ORDER BY net_sales DESC LIMIT 10",
                        params,
                    ).fetchall()
                ]
                channels = [
                    dict(r)
                    for r in con.execute(
                        f"SELECT COALESCE(NULLIF(channel, ''), 'Direct') AS Channel, COUNT(*) AS orders, SUM(amount - refund) AS net_sales FROM sales {where} GROUP BY Channel ORDER BY net_sales DESC LIMIT 10",
                        params,
                    ).fetchall()
                ]
                customers = [
                    dict(r)
                    for r in con.execute(
                        f"SELECT COALESCE(NULLIF(customer, ''), 'Guest') AS Customer, COUNT(*) AS orders, SUM(amount - refund) AS total_spent FROM sales {where} GROUP BY Customer ORDER BY total_spent DESC LIMIT 10",
                        params,
                    ).fetchall()
                ]
                rows = [
                    {
                        "Date": r["day"],
                        "Order_ID": r["order_id"],
                        "Channel": r["channel"],
                        "Product": r["product"],
                        "Customer": r["customer"],
                        "Amount": r["amount"] - r["refund"],
                    }
                    for r in con.execute(f"SELECT * FROM sales {where} ORDER BY day DESC, id DESC LIMIT 100", params).fetchall()
                ]
                self._json(
                    {
                        "datasets": datasets,
                        "dataset": active,
                        "metrics": {
                            "total_sales": round(total_sales, 2),
                            "order_count": order_count,
                            "aov": round(aov, 2),
                            "customers": int(m_row["customers"] or 0),
                        },
                        "monthly": monthly,
                        "products": products,
                        "channels": channels,
                        "customers": customers,
                        "rows": rows,
                    }
                )
            return
        if parsed.path in ("/", "/index.html"):
            content = (ROOT / "index.html").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return
        if parsed.path == "/favicon.svg":
            content = (ROOT / "favicon.svg").read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "image/svg+xml")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return
        self.send_error(404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length) if length > 0 else b""
        if parsed.path == "/api/active-dataset":
            data = json.loads(body.decode("utf-8") or "{}")
            ds_id = int(data.get("id") or 0)
            with connect() as con:
                con.execute("UPDATE datasets SET is_active = CASE WHEN id = ? THEN 1 ELSE 0 END", (ds_id,))
            self._json({"ok": True})
            return
        if parsed.path == "/api/upload":
            ctype = self.headers.get("Content-Type", "")
            try:
                fields, file_item = parse_multipart_form(ctype, body)
                if not file_item:
                    self._json({"ok": False, "error": "No file uploaded."}, 400)
                    return
                filename, raw = file_item
                records = parse_rows(filename, raw)
                if not records:
                    self._json({"ok": False, "error": "File has no valid data rows."}, 400)
                    return
                headers = list(records[0].keys())
                col_map = {k: map_column(headers, k) for k in ALIASES}
                name = fields.get("name") or Path(filename).stem
                now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                with connect() as con:
                    con.execute("UPDATE datasets SET is_active = 0")
                    cur = con.execute(
                        "INSERT INTO datasets (name, filename, row_count, is_active, created_at) VALUES (?, ?, ?, 1, ?)",
                        (name, filename, len(records), now),
                    )
                    ds_id = cur.lastrowid
                    batch = []
                    for idx, r in enumerate(records, start=1):
                        day = norm_date(r.get(col_map["date"], "")) if col_map["date"] else ""
                        oid = str(r.get(col_map["order_id"], "") or f"ORD-{idx}") if col_map["order_id"] else f"ORD-{idx}"
                        prod = str(r.get(col_map["product"], "") or "General") if col_map["product"] else "General"
                        chan = str(r.get(col_map["channel"], "") or "Direct") if col_map["channel"] else "Direct"
                        country = str(r.get(col_map["country"], "") or "") if col_map["country"] else ""
                        qty = to_float(r.get(col_map["qty"], 1)) if col_map["qty"] else 1.0
                        amt = to_float(r.get(col_map["amount"], 0)) if col_map["amount"] else 0.0
                        price = to_float(r.get(col_map["price"], amt)) if col_map["price"] else amt
                        refund = to_float(r.get(col_map["refund"], 0)) if col_map["refund"] else 0.0
                        cogs = to_float(r.get(col_map["cogs"], 0)) if col_map["cogs"] else 0.0
                        ship = to_float(r.get(col_map["shipping"], 0)) if col_map["shipping"] else 0.0
                        ads = to_float(r.get(col_map["ad_spend"], 0)) if col_map["ad_spend"] else 0.0
                        cust = str(r.get(col_map["customer"], "") or "Guest") if col_map["customer"] else "Guest"
                        camp = str(r.get(col_map["campaign"], "") or "") if col_map["campaign"] else ""
                        batch.append((ds_id, day, oid, prod, chan, country, qty, amt, price, refund, cogs, ship, ads, cust, camp))
                    con.executemany(
                        """
                        INSERT INTO sales (dataset_id, day, order_id, product, channel, country, qty, amount, price, refund, cogs, shipping, ad_spend, customer, campaign)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        batch,
                    )
                self._json({"ok": True, "dataset_id": ds_id})
            except Exception as exc:
                self._json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_error(404)


def main() -> None:
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", "8780"))
    host = os.environ.get("DATAPRISM_HOST", "127.0.0.1")
    srv = HTTPServer((host, port), LocalHandler)
    print(f"DataPrism local server listening on http://{host}:{port}")
    srv.serve_forever()


if __name__ == "__main__":
    main()
