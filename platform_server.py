"""DataPrism Multi-User Analytics Platform Server.

Provides authenticated workspaces, SQLite persistence, transparent dataset health
diagnostics, beginner-friendly statistical explanations, guided time-trend and
relationship explorers, editable category-specific analyses, and a validated
chart builder.
"""
from __future__ import annotations

import csv
import email.parser
import email.policy
import hashlib
import io
import json
import math
import os
import re
import secrets
import sqlite3
import statistics
import sys
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

try:
    import openpyxl
except ImportError:
    openpyxl = None

ROOT = Path(__file__).resolve().parent
DEFAULT_DB = ROOT / "data" / "platform.sqlite3"
DB_PATH = Path(os.environ.get("DATAPRISM_DB", str(DEFAULT_DB)))

MAX_UPLOAD = 50 * 1024 * 1024  # 50 MB
MAX_ROWS = 100_000
SESSION_DAYS = 30
PBKDF2_ROUNDS = 310_000
COOKIE_NAME = "mm_session"

CATEGORIES: dict[str, dict[str, Any]] = {
    "ecommerce": {
        "label": "Sales & E-commerce",
        "icon": "📦",
        "tagline": "Revenue trends, product/category totals, order volume, and profit margins (when cost/profit fields exist).",
        "roles": {
            "date": {"label": "Order / Transaction Date", "kind": "date", "aliases": ["order_date", "date", "day", "transaction_date", "created_at", "invoice_date"]},
            "revenue": {"label": "Revenue / Sales Amount", "kind": "number", "aliases": ["revenue", "sales", "net_sales", "amount", "total", "gross_sales", "order_value", "price"]},
            "cost": {"label": "Cost / COGS (Optional for Profit)", "kind": "number", "aliases": ["cost", "cogs", "unit_cost", "total_cost", "cost_of_goods"]},
            "profit": {"label": "Profit (Optional direct field)", "kind": "number", "aliases": ["profit", "net_profit", "gross_profit", "margin_amount"]},
            "quantity": {"label": "Quantity / Units", "kind": "number", "aliases": ["qty", "quantity", "units", "items_sold", "volume"]},
            "product": {"label": "Product / Item", "kind": "category", "aliases": ["product", "product_name", "item", "item_name", "title", "model"]},
            "category": {"label": "Category / Channel", "kind": "category", "aliases": ["category", "product_category", "channel", "sales_channel", "department", "region"]},
            "customer": {"label": "Customer Identifier", "kind": "any", "aliases": ["customer_id", "customer", "customer_name", "client", "buyer", "email"]},
        },
    },
    "retail": {
        "label": "Retail & Inventory",
        "icon": "🏪",
        "tagline": "Inventory stock levels, store/category breakdowns, reorder risks, and product movement summaries.",
        "roles": {
            "product": {"label": "Product / SKU Name", "kind": "any", "aliases": ["product", "product_name", "item", "item_name", "sku_name", "description"]},
            "category": {"label": "Store / Department / Category", "kind": "category", "aliases": ["store", "location", "branch", "category", "department", "aisle", "warehouse"]},
            "stock": {"label": "Stock / On-Hand Inventory", "kind": "number", "aliases": ["stock", "inventory", "on_hand", "qty_on_hand", "stock_level", "available_qty", "units_in_stock"]},
            "reorder": {"label": "Reorder Point / Min Stock (Optional)", "kind": "number", "aliases": ["reorder_point", "min_stock", "safety_stock", "reorder_level", "threshold"]},
            "sales_qty": {"label": "Units Sold / Movement (Optional)", "kind": "number", "aliases": ["units_sold", "qty_sold", "sales_qty", "qty", "quantity", "demand"]},
            "price": {"label": "Unit Price / Retail Value (Optional)", "kind": "number", "aliases": ["unit_price", "price", "retail_price", "value", "revenue", "sales"]},
            "date": {"label": "Count / Transaction Date (Optional)", "kind": "date", "aliases": ["date", "last_restocked", "count_date", "transaction_date", "day"]},
        },
    },
    "education": {
        "label": "Education & Academics",
        "icon": "🎓",
        "tagline": "Enrollment counts, attendance rates, course/program breakdowns, and student grade distributions.",
        "roles": {
            "group": {"label": "Course / Program / Grade Level", "kind": "category", "aliases": ["course", "program", "department", "subject", "class", "grade_level", "major", "cohort", "section"]},
            "score": {"label": "Grade / Exam Score", "kind": "number", "aliases": ["grade", "score", "exam_score", "final_score", "gpa", "marks", "percentage", "test_score"]},
            "attendance": {"label": "Attendance Rate / Days", "kind": "number", "aliases": ["attendance", "attendance_rate", "attendance_pct", "days_present", "present_pct"]},
            "enrollment": {"label": "Enrollment Count / Credits (Optional)", "kind": "number", "aliases": ["enrollment", "students", "headcount", "credits", "enrolled_count"]},
            "student": {"label": "Student Identifier (Optional)", "kind": "any", "aliases": ["student_id", "student", "student_name", "learner_id", "roll_no"]},
            "date": {"label": "Term / Assessment Date (Optional)", "kind": "date", "aliases": ["date", "term_date", "exam_date", "enrollment_date", "semester_date"]},
        },
    },
    "marketing": {
        "label": "Marketing & Campaigns",
        "icon": "📣",
        "tagline": "Campaign performance, spend vs. outcome efficiency (CPA/ROAS), and channel conversion breakdowns.",
        "roles": {
            "campaign": {"label": "Campaign / Channel", "kind": "category", "aliases": ["campaign", "campaign_name", "channel", "source", "medium", "ad_group", "platform"]},
            "spend": {"label": "Ad Spend / Cost", "kind": "number", "aliases": ["spend", "ad_spend", "cost", "budget", "media_cost", "amount_spent"]},
            "conversions": {"label": "Conversions / Leads / Signups", "kind": "number", "aliases": ["conversions", "leads", "signups", "purchases", "orders", "acquisitions", "outcomes"]},
            "revenue": {"label": "Attributed Revenue (Optional for ROAS)", "kind": "number", "aliases": ["revenue", "attributed_revenue", "conversion_value", "sales", "return"]},
            "clicks": {"label": "Clicks / Visits (Optional)", "kind": "number", "aliases": ["clicks", "visits", "sessions", "traffic"]},
            "impressions": {"label": "Impressions / Reach (Optional)", "kind": "number", "aliases": ["impressions", "reach", "views"]},
            "date": {"label": "Campaign Date (Optional)", "kind": "date", "aliases": ["date", "day", "start_date", "report_date"]},
        },
    },
    "customers": {
        "label": "Customer Analytics",
        "icon": "👥",
        "tagline": "Customer counts, segment comparisons, repeat activity, and revenue concentration.",
        "roles": {
            "customer": {"label": "Customer ID / Name", "kind": "any", "aliases": ["customer_id", "customer", "user_id", "client_id", "account_id", "email", "customer_name", "member_id"]},
            "segment": {"label": "Customer Segment / Tier / Region", "kind": "category", "aliases": ["segment", "tier", "plan", "cohort", "region", "country", "customer_type", "status", "channel"]},
            "value": {"label": "Spend / Revenue / LTV (Optional)", "kind": "number", "aliases": ["spend", "total_spent", "revenue", "amount", "ltv", "mrr", "order_value", "sales"]},
            "frequency": {"label": "Orders / Sessions / Activity Count (Optional)", "kind": "number", "aliases": ["orders", "order_count", "sessions", "visits", "logins", "transactions", "frequency"]},
            "date": {"label": "Activity / Signup Date (Optional)", "kind": "date", "aliases": ["date", "order_date", "last_seen", "signup_date", "created_at", "activity_date"]},
        },
    },
    "finance": {
        "label": "Finance & Accounting",
        "icon": "💼",
        "tagline": "Account/department totals, budget vs. actual variance, and cashflow trends when corresponding fields exist.",
        "roles": {
            "category": {"label": "Account / Department / Cost Center", "kind": "category", "aliases": ["account", "department", "category", "cost_center", "type", "ledger"]},
            "actual": {"label": "Actual Amount", "kind": "number", "aliases": ["actual", "amount", "total", "expense", "revenue", "balance", "net_amount"]},
            "budget": {"label": "Budget / Planned Amount (Optional)", "kind": "number", "aliases": ["budget", "planned", "forecast", "target", "allocated"]},
            "date": {"label": "Posting / Transaction Date (Optional)", "kind": "date", "aliases": ["date", "posting_date", "transaction_date", "month", "period"]},
        },
    },
    "product": {
        "label": "Product & SaaS",
        "icon": "🚀",
        "tagline": "Plan/tier breakdowns, MRR/subscription metrics, and feature usage or churn indicators.",
        "roles": {
            "segment": {"label": "Plan / Tier / Feature", "kind": "category", "aliases": ["plan", "tier", "subscription_tier", "feature", "cohort", "status", "segment"]},
            "revenue": {"label": "MRR / ARR / Subscription Revenue", "kind": "number", "aliases": ["mrr", "arr", "revenue", "amount", "price", "subscription_value"]},
            "usage": {"label": "Active Users / Usage / Seats (Optional)", "kind": "number", "aliases": ["usage", "active_users", "seats", "sessions", "events", "logins"]},
            "date": {"label": "Activity / Billing Date (Optional)", "kind": "date", "aliases": ["date", "billing_date", "signup_date", "month", "created_at"]},
        },
    },
    "general": {
        "label": "General / Custom Dataset",
        "icon": "🔍",
        "tagline": "Universal exploratory data analysis for any tabular dataset without domain assumptions.",
        "roles": {
            "category": {"label": "Primary Grouping Category", "kind": "category", "aliases": ["category", "group", "type", "status", "region", "department"]},
            "measure": {"label": "Primary Numeric Measure", "kind": "number", "aliases": ["value", "amount", "total", "score", "count", "metric", "revenue"]},
            "date": {"label": "Date / Time Column (Optional)", "kind": "date", "aliases": ["date", "day", "created_at", "timestamp", "period"]},
        },
    },
}

# Alias legacy keys so existing saved datasets never break
for _legacy_key, _fallback in [("operations", "general"), ("logistics", "retail"), ("research", "general")]:
    if _legacy_key not in CATEGORIES:
        CATEGORIES[_legacy_key] = {
            **CATEGORIES[_fallback],
            "label": _legacy_key.capitalize() + " Analytics",
        }

ID_NAME_PATTERNS = re.compile(
    r"(?:^|_)(?:id|uuid|guid|sku|ssn|isbn|asin|upc|ean|zip|zipcode|postal|postal_code|phone|mobile|tel|order_no|invoice_no|tx_id|trans_id|transaction_id|record_id|row_id|emp_id|student_id|customer_id|user_id|account_id|ticket_id)(?:$|_)",
    re.IGNORECASE,
)


def is_ephemeral_storage() -> bool:
    """Detect if the SQLite database path is on ephemeral storage (such as /tmp on Render free plan)."""
    db_str = str(DB_PATH.resolve())
    if db_str.startswith("/tmp") or db_str.startswith("/var/tmp"):
        return True
    if os.environ.get("RENDER") and not os.environ.get("RENDER_DISK_MOUNT_PATH"):
        return True
    return False


def storage_info() -> dict[str, Any]:
    ephemeral = is_ephemeral_storage()
    return {
        "db_path": str(DB_PATH),
        "ephemeral": ephemeral,
        "mode": "ephemeral" if ephemeral else "persistent",
        "notice": (
            "Demo Storage Notice: This instance stores SQLite data in temporary container storage (/tmp). "
            "Saved accounts and uploaded datasets will reset when the free Render service sleeps, restarts, or redeploys. "
            "Export any charts or cleaned CSVs you want to keep."
            if ephemeral
            else "Local SQLite storage is active. Uploaded datasets and accounts persist in your local data directory."
        ),
    }


def connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_PATH), timeout=30)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys=ON")
    con.execute("PRAGMA journal_mode=WAL")
    con.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE,
            salt TEXT NOT NULL,
            pass_hash TEXT NOT NULL,
            password_hash TEXT,
            category TEXT NOT NULL DEFAULT 'ecommerce',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS datasets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            category TEXT NOT NULL DEFAULT 'ecommerce',
            created_at TEXT NOT NULL,
            row_count INTEGER NOT NULL,
            columns_json TEXT NOT NULL,
            parent_id INTEGER REFERENCES datasets(id) ON DELETE SET NULL,
            transform_note TEXT DEFAULT '',
            quality_notes_json TEXT DEFAULT '{}',
            custom_units_json TEXT DEFAULT '{}',
            role_overrides_json TEXT DEFAULT '{}'
        );

        CREATE TABLE IF NOT EXISTS dataset_rows (
            dataset_id INTEGER NOT NULL REFERENCES datasets(id) ON DELETE CASCADE,
            row_num INTEGER NOT NULL,
            payload TEXT NOT NULL,
            PRIMARY KEY (dataset_id, row_num)
        );

        CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id);
        CREATE INDEX IF NOT EXISTS idx_datasets_user ON datasets(user_id);
        CREATE INDEX IF NOT EXISTS idx_rows_dataset ON dataset_rows(dataset_id);
        """
    )
    # Safe migrations for pre-existing databases.
    # Earlier DataPrism builds stored the same PBKDF2 digest in `password_hash`;
    # this build names that field `pass_hash`. Copy it so existing accounts keep
    # working without deleting or recreating the user's database.
    user_cols = {r["name"] for r in con.execute("PRAGMA table_info(users)").fetchall()}
    if "category" not in user_cols:
        con.execute("ALTER TABLE users ADD COLUMN category TEXT NOT NULL DEFAULT 'ecommerce'")
    if "pass_hash" not in user_cols:
        con.execute("ALTER TABLE users ADD COLUMN pass_hash TEXT")
        if "password_hash" in user_cols:
            con.execute("UPDATE users SET pass_hash = password_hash WHERE password_hash IS NOT NULL")
    if "password_hash" not in user_cols:
        con.execute("ALTER TABLE users ADD COLUMN password_hash TEXT")

    # Previous builds stored only a hash of each session token. Those cookies
    # cannot be resumed by this build, so replace only the session table; user
    # accounts and datasets remain intact and users can simply sign in again.
    session_cols = {r["name"] for r in con.execute("PRAGMA table_info(sessions)").fetchall()}
    if "token" not in session_cols:
        con.execute("DROP INDEX IF EXISTS idx_sessions_user")
        con.execute("ALTER TABLE sessions RENAME TO sessions_legacy")
        con.execute("""CREATE TABLE sessions (
            token TEXT PRIMARY KEY,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL
        )""")
        con.execute("DROP TABLE sessions_legacy")
        con.execute("CREATE INDEX idx_sessions_user ON sessions(user_id)")

    ds_cols = {r["name"] for r in con.execute("PRAGMA table_info(datasets)").fetchall()}
    for col_name, col_def in [
        ("parent_id", "INTEGER REFERENCES datasets(id) ON DELETE SET NULL"),
        ("transform_note", "TEXT DEFAULT ''"),
        ("quality_notes_json", "TEXT DEFAULT '{}'"),
        ("custom_units_json", "TEXT DEFAULT '{}'"),
        ("role_overrides_json", "TEXT DEFAULT '{}'"),
    ]:
        if col_name not in ds_cols:
            con.execute(f"ALTER TABLE datasets ADD COLUMN {col_name} {col_def}")
    return con


def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def hash_password(password: str, salt_hex: str | None = None) -> tuple[str, str]:
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ROUNDS)
    return salt.hex(), digest.hex()


def norm_name(raw: Any, used: set[str]) -> str:
    text = str(raw or "").strip()
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "_", text.lower()).strip("_") or "column"
    if cleaned[0].isdigit():
        cleaned = "col_" + cleaned
    candidate = cleaned
    idx = 2
    while candidate in used:
        candidate = f"{cleaned}_{idx}"
        idx += 1
    used.add(candidate)
    return candidate


def is_missing(val: Any) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and math.isnan(val):
        return True
    s = str(val).strip()
    return s == "" or s.lower() in {"null", "none", "na", "n/a", "nan", "-", "--", "nil", "#n/a", "undefined"}


def to_number(val: Any) -> float | None:
    if is_missing(val):
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        f = float(val)
        return f if math.isfinite(f) else None
    s = str(val).strip()
    # Reject strings that look like dates or alphanumeric codes with internal hyphens
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", s) or re.match(r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$", s):
        return None
    # Strip currency symbols, whitespace, commas, and trailing %
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    s = re.sub(r"^[\$€£¥₹\s]+", "", s)
    s = re.sub(r"[\$€£¥₹%\s]+$", "", s)
    s = s.replace(",", "")
    if not s:
        return None
    try:
        num = float(s)
        if not math.isfinite(num):
            return None
        return -num if neg else num
    except ValueError:
        return None


def parse_date_value(val: Any, prefer_day_first: bool = False) -> tuple[str | None, bool]:
    """Return (YYYY-MM-DD or None, was_ambiguous_slash_date)."""
    if is_missing(val):
        return None, False
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d"), False
    if hasattr(val, "strftime"):
        try:
            return val.strftime("%Y-%m-%d"), False
        except Exception:
            return None, False
    if isinstance(val, (int, float)):
        return None, False
    s = str(val).strip()
    if len(s) < 6:
        return None, False
    # Split off time component if ISO or space-separated
    date_part = re.split(r"[T\s]", s)[0].strip()
    # Unambiguous ISO format YYYY-MM-DD or YYYY/MM/DD
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            dt = datetime.strptime(date_part, fmt)
            if 1900 <= dt.year <= 2100:
                return dt.strftime("%Y-%m-%d"), False
        except ValueError:
            pass
    # Named month formats
    for fmt in ("%b %d, %Y", "%B %d, %Y", "%d %b %Y", "%d %B %Y", "%d-%b-%Y", "%Y-%m"):
        try:
            dt = datetime.strptime(s if "," in fmt or " " in fmt else date_part, fmt)
            if 1900 <= dt.year <= 2100:
                return dt.strftime("%Y-%m-%d"), False
        except ValueError:
            pass
    # Numeric slash or dash formats: check for MM/DD/YYYY vs DD/MM/YYYY ambiguity
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$", date_part)
    if m:
        p1, p2, yr_raw = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = yr_raw + 2000 if yr_raw < 100 else yr_raw
        if not (1900 <= year <= 2100):
            return None, False
        ambiguous = (1 <= p1 <= 12) and (1 <= p2 <= 12) and (p1 != p2)
        if prefer_day_first:
            day_first_order = [(p2, p1), (p1, p2)]
        else:
            day_first_order = [(p1, p2), (p2, p1)]
        for month_cand, day_cand in day_first_order:
            try:
                dt = datetime(year, month_cand, day_cand)
                return dt.strftime("%Y-%m-%d"), ambiguous
            except ValueError:
                continue
    return None, False


def detect_date_preference(values: list[Any]) -> tuple[bool, int]:
    """Inspect slash/dash dates across the column to see if DD/MM or MM/DD is indicated."""
    first_gt_12 = 0
    second_gt_12 = 0
    ambiguous_count = 0
    for v in values:
        if is_missing(v) or not isinstance(v, str):
            continue
        date_part = re.split(r"[T\s]", v.strip())[0]
        m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$", date_part)
        if m:
            p1, p2 = int(m.group(1)), int(m.group(2))
            if p1 > 12 and p2 <= 12:
                first_gt_12 += 1
            elif p2 > 12 and p1 <= 12:
                second_gt_12 += 1
            elif 1 <= p1 <= 12 and 1 <= p2 <= 12 and p1 != p2:
                ambiguous_count += 1
    prefer_day_first = first_gt_12 > second_gt_12
    # Only flag ambiguity if no disambiguating dates (>12) were found in the column
    unresolved_ambiguous = ambiguous_count if (first_gt_12 == 0 and second_gt_12 == 0) else 0
    return prefer_day_first, unresolved_ambiguous


def infer_column_meta(name: str, label: str, values: list[Any]) -> dict[str, Any]:
    non_missing = [v for v in values if not is_missing(v)]
    total = len(values)
    valid_count = len(non_missing)
    missing_count = total - valid_count

    if valid_count == 0:
        return {
            "name": name,
            "label": label,
            "kind": "text",
            "is_identifier": False,
            "id_reason": "",
            "non_null_count": 0,
            "missing_count": missing_count,
            "distinct_count": 0,
            "unparsed_count": 0,
            "ambiguous_date_count": 0,
            "date_format_note": "",
        }

    name_is_id = bool(ID_NAME_PATTERNS.search(name) or ID_NAME_PATTERNS.search(label))
    prefer_day_first, unresolved_ambiguous = detect_date_preference(non_missing)

    num_Count = 0
    date_count = 0
    for v in non_missing:
        if to_number(v) is not None:
            num_Count += 1
        d_val, _ = parse_date_value(v, prefer_day_first=prefer_day_first)
        if d_val is not None:
            date_count += 1

    num_ratio = num_Count / valid_count
    date_ratio = date_count / valid_count

    str_vals = [str(v).strip() for v in non_missing]
    distinct_count = len(set(str_vals))
    distinct_ratio = distinct_count / valid_count

    # Determine if column is an identifier
    is_identifier = False
    id_reason = ""
    if name_is_id and date_ratio < 0.75:
        is_identifier = True
        id_reason = f"Column name '{label}' matches identifier/code naming patterns."
    elif date_ratio < 0.75 and num_ratio < 0.75 and valid_count >= 10 and distinct_ratio >= 0.92:
        avg_len = sum(len(s) for s in str_vals) / len(str_vals)
        if avg_len <= 36 and all(" " not in s for s in str_vals[:25]):
            is_identifier = True
            id_reason = f"{distinct_count} of {valid_count} non-empty values ({round(distinct_ratio*100)}%) are unique codes."
    elif num_ratio >= 0.95 and valid_count >= 15 and distinct_count == valid_count:
        # Check if strictly sequential integers (1, 2, 3... or 1001, 1002...)
        nums = [to_number(v) for v in non_missing]
        nums_clean = [n for n in nums if n is not None]
        if all(float(n).is_integer() for n in nums_clean):
            sorted_nums = sorted(int(n) for n in nums_clean)
            diffs = [sorted_nums[i + 1] - sorted_nums[i] for i in range(len(sorted_nums) - 1)]
            if diffs and all(d == 1 for d in diffs):
                is_identifier = True
                id_reason = "Values are strictly sequential unique integers (1-by-1 row index/ID)."

    unparsed_count = 0
    date_format_note = ""

    if date_ratio >= 0.75 and not name_is_id:
        kind = "date"
        unparsed_count = valid_count - date_count
        if unresolved_ambiguous > 0:
            date_format_note = (
                f"{unresolved_ambiguous} slash/dash date(s) had both day and month <= 12 (parsed as MM/DD/YYYY by default; "
                "use YYYY-MM-DD to avoid ambiguity)."
            )
        elif prefer_day_first:
            date_format_note = "Detected DD/MM/YYYY date format and normalized to YYYY-MM-DD."
    elif is_identifier:
        kind = "id"
        unparsed_count = 0
    elif num_ratio >= 0.75:
        kind = "number"
        unparsed_count = valid_count - num_Count
    else:
        max_cats = max(25, min(100, int(valid_count * 0.5)))
        if distinct_count <= max_cats or (valid_count <= 20 and distinct_count < valid_count):
            kind = "category"
        else:
            kind = "text"

    return {
        "name": name,
        "label": label,
        "kind": kind,
        "is_identifier": is_identifier,
        "id_reason": id_reason,
        "non_null_count": valid_count,
        "missing_count": missing_count,
        "distinct_count": distinct_count,
        "unparsed_count": unparsed_count,
        "ambiguous_date_count": unresolved_ambiguous if kind == "date" else 0,
        "date_format_note": date_format_note,
        "prefer_day_first": prefer_day_first,
    }


def decode_upload(filename: str, blob: bytes) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    """Parse CSV, TSV, JSON, JSONL, or XLSX upload into (columns_meta, normalized_records, parse_diagnostics)."""
    if not blob:
        raise ValueError("The uploaded file is empty (0 bytes). Choose a file with a header row and data rows.")
    if len(blob) > MAX_UPLOAD:
        raise ValueError(f"File exceeds the {MAX_UPLOAD // (1024 * 1024)} MB limit.")

    ext = Path(filename or "upload.csv").suffix.lower()
    raw_headers: list[str] = []
    raw_rows: list[list[Any]] = []
    encoding_used = "utf-8"
    delimiter_used = ","
    empty_rows_skipped = 0

    if ext in {".xlsx", ".xlsm"}:
        if openpyxl is None:
            raise ValueError("Excel (.xlsx) support requires openpyxl, which is not installed in this environment.")
        try:
            book = openpyxl.load_workbook(io.BytesIO(blob), read_only=True, data_only=True)
        except Exception as exc:
            raise ValueError(f"Could not read Excel workbook '{filename}'. Ensure it is a valid, unencrypted .xlsx file. ({exc})") from exc
        sheet = book.active
        if sheet is None:
            raise ValueError("The Excel workbook does not contain an active worksheet.")
        # Fixed bug: iterate rows into a single 2D list of rows (not double-bracketed)
        all_rows = [list(r) for r in sheet.iter_rows(values_only=True)]
        non_empty_rows = []
        for r in all_rows:
            if any(not is_missing(cell) for cell in r):
                non_empty_rows.append(r)
            else:
                empty_rows_skipped += 1
        if len(non_empty_rows) < 2:
            raise ValueError("Excel sheet must include a header row and at least one row of data.")
        raw_headers = [str(h if h is not None else f"Column_{i+1}").strip() for i, h in enumerate(non_empty_rows[0])]
        raw_rows = non_empty_rows[1:]
        encoding_used = "xlsx (binary)"
        delimiter_used = "worksheet cells"

    elif ext in {".json", ".jsonl", ".ndjson"}:
        text = blob.decode("utf-8-sig", errors="replace").strip()
        if not text:
            raise ValueError("JSON file is empty.")
        try:
            if ext in {".jsonl", ".ndjson"} or ("\n" in text and not text.startswith("[")):
                objs = [json.loads(line) for line in text.splitlines() if line.strip()]
            else:
                parsed_json = json.loads(text)
                if isinstance(parsed_json, dict):
                    # Check if a top-level list key exists (e.g. {"data": [...]})
                    list_val = next((v for v in parsed_json.values() if isinstance(v, list)), None)
                    objs = list_val if list_val is not None else [parsed_json]
                elif isinstance(parsed_json, list):
                    objs = parsed_json
                else:
                    raise ValueError("JSON root must be an array of objects or a record object.")
        except json.JSONDecodeError as exc:
            raise ValueError(f"Invalid JSON syntax on line {exc.lineno}: {exc.msg}") from exc
        if not objs or not isinstance(objs[0], dict):
            raise ValueError("JSON file must contain at least one record object with key-value fields.")
        header_set: list[str] = []
        for obj in objs:
            if isinstance(obj, dict):
                for k in obj.keys():
                    if str(k) not in header_set:
                        header_set.append(str(k))
        raw_headers = header_set
        for obj in objs:
            if isinstance(obj, dict) and any(not is_missing(v) for v in obj.values()):
                raw_rows.append([obj.get(h) for h in raw_headers])
            else:
                empty_rows_skipped += 1
        encoding_used = "utf-8"
        delimiter_used = "json keys"

    elif ext in {".csv", ".tsv", ".txt", ""}:
        text = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                text = blob.decode(enc)
                encoding_used = "utf-8" if enc == "utf-8-sig" else enc
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            text = blob.decode("utf-8", errors="replace")
            encoding_used = "utf-8 (with replacement)"

        sample_lines = "\n".join(text.splitlines()[:25])
        if ext == ".tsv":
            delimiter_used = "\t"
        else:
            try:
                dialect = csv.Sniffer().sniff(sample_lines, delimiters=",;\t|")
                delimiter_used = dialect.delimiter
            except Exception:
                counts = {d: sample_lines.count(d) for d in (",", ";", "\t", "|")}
                delimiter_used = max(counts, key=counts.get) if max(counts.values()) > 0 else ","

        reader = csv.reader(io.StringIO(text), delimiter=delimiter_used)
        all_rows = list(reader)
        non_empty = []
        for r in all_rows:
            if any(str(c).strip() != "" for c in r):
                non_empty.append(r)
            else:
                empty_rows_skipped += 1
        if len(non_empty) < 2:
            raise ValueError("CSV file must include a header row and at least one row of data.")
        raw_headers = [str(h or f"Column_{i+1}").strip() for i, h in enumerate(non_empty[0])]
        raw_rows = non_empty[1:]
    else:
        raise ValueError(f"Unsupported file type '{ext}'. Upload a .csv, .tsv, .xlsx, .json, or .jsonl file.")

    # Trim completely empty trailing columns
    num_cols = len(raw_headers)
    keep_indices = [
        i
        for i in range(num_cols)
        if raw_headers[i] != f"Column_{i+1}" or any(i < len(r) and not is_missing(r[i]) for r in raw_rows)
    ]
    if not keep_indices:
        raise ValueError("No valid columns with data were found in the uploaded file.")

    raw_headers = [raw_headers[i] for i in keep_indices]
    raw_rows = [[r[i] if i < len(r) else None for i in keep_indices] for r in raw_rows]

    truncated_rows = 0
    if len(raw_rows) > MAX_ROWS:
        truncated_rows = len(raw_rows) - MAX_ROWS
        raw_rows = raw_rows[:MAX_ROWS]

    used_names: set[str] = set()
    col_names: list[str] = []
    col_labels: list[str] = []
    renamed_columns: list[str] = []
    for idx, h in enumerate(raw_headers):
        label = h.strip() or f"Column {idx + 1}"
        norm = norm_name(label, used_names)
        if label.lower() != norm and label in raw_headers[:idx]:
            renamed_columns.append(f"Duplicate column header '{label}' renamed to '{norm}'")
        col_names.append(norm)
        col_labels.append(label)

    col_values: dict[str, list[Any]] = {c: [] for c in col_names}
    for r in raw_rows:
        for idx, c in enumerate(col_names):
            col_values[c].append(r[idx] if idx < len(r) else None)

    cols_meta: list[dict[str, Any]] = []
    for idx, c in enumerate(col_names):
        meta = infer_column_meta(c, col_labels[idx], col_values[c])
        cols_meta.append(meta)

    records: list[dict[str, Any]] = []
    for r in raw_rows:
        rec: dict[str, Any] = {}
        for idx, meta in enumerate(cols_meta):
            c = meta["name"]
            raw_val = r[idx] if idx < len(r) else None
            if is_missing(raw_val):
                rec[c] = None
            elif meta["kind"] == "number":
                num_v = to_number(raw_val)
                rec[c] = num_v if num_v is not None else str(raw_val).strip()
            elif meta["kind"] == "date":
                dt_v, _ = parse_date_value(raw_val, prefer_day_first=meta.get("prefer_day_first", False))
                rec[c] = dt_v if dt_v is not None else str(raw_val).strip()
            else:
                rec[c] = str(raw_val).strip()
        records.append(rec)

    diagnostics = {
        "filename": filename,
        "encoding": encoding_used,
        "delimiter": "tab (\\t)" if delimiter_used == "\t" else delimiter_used,
        "empty_rows_skipped": empty_rows_skipped,
        "truncated_rows": truncated_rows,
        "renamed_columns": renamed_columns,
        "unparsed_by_column": {m["name"]: m["unparsed_count"] for m in cols_meta if m.get("unparsed_count", 0) > 0},
        "date_notes_by_column": {m["name"]: m["date_format_note"] for m in cols_meta if m.get("date_format_note")},
    }
    return cols_meta, records, diagnostics


def save_dataset(
    con: sqlite3.Connection,
    user_id: int,
    name: str,
    category: str,
    cols: list[dict[str, Any]],
    records: list[dict[str, Any]],
    parent_id: int | None = None,
    transform_note: str = "",
    diagnostics: dict[str, Any] | None = None,
    custom_units: dict[str, str] | None = None,
    role_overrides: dict[str, str] | None = None,
) -> int:
    clean_cols = [
        {
            "name": c["name"],
            "label": c.get("label", c["name"]),
            "kind": c["kind"],
            "is_identifier": bool(c.get("is_identifier", c["kind"] == "id")),
            "id_reason": c.get("id_reason", ""),
            "unparsed_count": int(c.get("unparsed_count", 0)),
            "date_format_note": c.get("date_format_note", ""),
        }
        for c in cols
    ]
    cur = con.execute(
        """
        INSERT INTO datasets (
            user_id, name, category, created_at, row_count, columns_json,
            parent_id, transform_note, quality_notes_json, custom_units_json, role_overrides_json
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            user_id,
            name.strip() or "Untitled Dataset",
            category if category in CATEGORIES else "general",
            now_iso(),
            len(records),
            json.dumps(clean_cols),
            parent_id,
            transform_note,
            json.dumps(diagnostics or {}),
            json.dumps(custom_units or {}),
            json.dumps(role_overrides or {}),
        ),
    )
    ds_id = int(cur.lastrowid)
    batch = [(ds_id, i + 1, json.dumps(rec)) for i, rec in enumerate(records)]
    con.executemany("INSERT INTO dataset_rows (dataset_id, row_num, payload) VALUES (?, ?, ?)", batch)
    return ds_id


def load_dataset_and_rows(
    con: sqlite3.Connection, user_id: int, dataset_id: int | None = None
) -> tuple[dict[str, Any] | None, list[dict[str, Any]]]:
    if dataset_id:
        row = con.execute("SELECT * FROM datasets WHERE id = ? AND user_id = ?", (dataset_id, user_id)).fetchone()
    else:
        row = con.execute("SELECT * FROM datasets WHERE user_id = ? ORDER BY id DESC LIMIT 1", (user_id,)).fetchone()
    if not row:
        return None, []
    ds = dict(row)
    ds["columns"] = json.loads(ds.get("columns_json") or "[]")
    ds["diagnostics"] = json.loads(ds.get("quality_notes_json") or "{}")
    ds["custom_units"] = json.loads(ds.get("custom_units_json") or "{}")
    ds["role_overrides"] = json.loads(ds.get("role_overrides_json") or "{}")
    raw_rows = con.execute(
        "SELECT payload FROM dataset_rows WHERE dataset_id = ? ORDER BY row_num ASC", (ds["id"],)
    ).fetchall()
    records = [json.loads(r["payload"]) for r in raw_rows]
    return ds, records


def apply_filter(records: list[dict[str, Any]], field: str, op: str, value: str) -> list[dict[str, Any]]:
    """Filter rows with type-aware comparison so numeric equality works seamlessly."""
    if not field or not op:
        return records
    op = op.lower().strip()
    if op == "is_empty":
        return [r for r in records if is_missing(r.get(field))]
    if op == "not_empty":
        return [r for r in records if not is_missing(r.get(field))]
    if value == "" or value is None:
        return records

    target_num = to_number(value)
    target_str = str(value).strip().lower()
    out: list[dict[str, Any]] = []

    for r in records:
        cell = r.get(field)
        if is_missing(cell):
            continue
        cell_num = to_number(cell)
        cell_str = str(cell).strip().lower()

        if op in {"eq", "is", "="}:
            if target_num is not None and cell_num is not None:
                if math.isclose(cell_num, target_num, rel_tol=1e-9, abs_tol=1e-9):
                    out.append(r)
            elif cell_str == target_str:
                out.append(r)
        elif op in {"neq", "not", "!="}:
            if target_num is not None and cell_num is not None:
                if not math.isclose(cell_num, target_num, rel_tol=1e-9, abs_tol=1e-9):
                    out.append(r)
            elif cell_str != target_str:
                out.append(r)
        elif op == "contains":
            if target_str in cell_str:
                out.append(r)
        elif op in {"gt", "gte", "lt", "lte"}:
            if target_num is not None and cell_num is not None:
                if (
                    (op == "gt" and cell_num > target_num)
                    or (op == "gte" and cell_num >= target_num)
                    or (op == "lt" and cell_num < target_num)
                    or (op == "lte" and cell_num <= target_num)
                ):
                    out.append(r)
            else:
                if (
                    (op == "gt" and cell_str > target_str)
                    or (op == "gte" and cell_str >= target_str)
                    or (op == "lt" and cell_str < target_str)
                    or (op == "lte" and cell_str <= target_str)
                ):
                    out.append(r)
    return out


def compute_quantiles(sorted_vals: list[float]) -> tuple[float, float, float]:
    """Compute (Q1, Median, Q3) on a non-empty sorted list of floats."""
    n = len(sorted_vals)
    if n == 1:
        return sorted_vals[0], sorted_vals[0], sorted_vals[0]
    med = float(statistics.median(sorted_vals))
    if n < 4:
        return sorted_vals[0], med, sorted_vals[-1]
    qs = statistics.quantiles(sorted_vals, n=4, method="inclusive")
    return float(qs[0]), float(qs[1]), float(qs[2])


def build_dataset_health(
    rows: list[dict[str, Any]], cols: list[dict[str, Any]], diagnostics: dict[str, Any] | None = None
) -> dict[str, Any]:
    """Section 1: Dataset Health diagnostics, detected types, warnings, and non-destructive vs mutating explanations."""
    diagnostics = diagnostics or {}
    row_count = len(rows)
    col_count = len(cols)
    total_cells = row_count * col_count
    missing_cells = 0
    complete_rows = 0

    seen_rows: set[str] = set()
    duplicate_rows = 0
    for r in rows:
        key = json.dumps(r, sort_keys=True, default=str)
        if key in seen_rows:
            duplicate_rows += 1
        else:
            seen_rows.add(key)
        row_missing = sum(1 for c in cols if is_missing(r.get(c["name"])))
        missing_cells += row_missing
        if row_missing == 0:
            complete_rows += 1

    col_reports: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    if duplicate_rows > 0:
        dup_pct = round((duplicate_rows / max(1, row_count)) * 100, 1)
        warnings.append(
            {
                "severity": "warning",
                "title": f"{duplicate_rows} exact duplicate row(s) detected ({dup_pct}% of dataset)",
                "columns": ["All columns"],
                "what_it_means": "Identical rows appear more than once. If these are accidental double-exports rather than separate transactions, totals and counts will be inflated.",
                "action_effect": "Reporting only — rows are still included until you click 'Remove Duplicate Rows' (which creates a new cleaned dataset copy).",
            }
        )

    if diagnostics.get("empty_rows_skipped", 0) > 0:
        warnings.append(
            {
                "severity": "info",
                "title": f"{diagnostics['empty_rows_skipped']} completely blank row(s) skipped during import",
                "columns": ["All columns"],
                "what_it_means": "Blank lines in your source file were ignored automatically so they do not count as missing data records.",
                "action_effect": "Applied automatically during import.",
            }
        )

    for c in cols:
        cname = c["name"]
        clabel = c.get("label", cname)
        ckind = c["kind"]
        is_id = bool(c.get("is_identifier", ckind == "id"))

        raw_vals = [r.get(cname) for r in rows]
        non_null = [v for v in raw_vals if not is_missing(v)]
        miss_cnt = row_count - len(non_null)
        miss_pct = round((miss_cnt / max(1, row_count)) * 100, 1)
        distinct_cnt = len({str(v) for v in non_null})

        unparsed_cnt = 0
        if ckind == "number":
            unparsed_cnt = sum(1 for v in non_null if to_number(v) is None)
        elif ckind == "date":
            unparsed_cnt = sum(1 for v in non_null if parse_date_value(v)[0] is None)

        samples = []
        for v in non_null:
            sv = str(v)
            if sv not in samples:
                samples.append(sv)
            if len(samples) >= 3:
                break

        type_explanation = {
            "number": "Quantitative numeric measure (used in Numeric Summary, Time Trends, and Relationships).",
            "date": "Calendar date normalized to YYYY-MM-DD (used in Trends Over Time).",
            "category": "Grouping/categorical field with repeated values (used in Category Breakdown).",
            "id": "Unique identifier or reference code (excluded from numeric math and category rankings).",
            "text": "Free-form or high-cardinality text field (searchable in Data Preview).",
        }.get(ckind, "Text column.")

        col_reports.append(
            {
                "name": cname,
                "label": clabel,
                "kind": ckind,
                "is_identifier": is_id,
                "id_reason": c.get("id_reason", ""),
                "type_explanation": type_explanation,
                "valid_count": len(non_null),
                "missing_count": miss_cnt,
                "missing_pct": miss_pct,
                "distinct_count": distinct_cnt,
                "unparsed_count": unparsed_cnt,
                "sample_values": samples,
                "date_format_note": c.get("date_format_note", ""),
            }
        )

        if miss_cnt > 0:
            warnings.append(
                {
                    "severity": "warning" if miss_pct >= 10 else "info",
                    "title": f"Column '{clabel}' has {miss_cnt} missing value(s) ({miss_pct}%)",
                    "columns": [clabel],
                    "what_it_means": (
                        f"{miss_cnt} of {row_count} rows are blank or null in '{clabel}'. "
                        "Averages, correlations, and charts using this column automatically skip missing rows rather than treating blanks as zero."
                    ),
                    "action_effect": "Reporting only — you can keep blanks as-is or use a cleaning action to create a new dataset with blanks removed or filled.",
                }
            )

        if unparsed_cnt > 0:
            warnings.append(
                {
                    "severity": "warning",
                    "title": f"Column '{clabel}' has {unparsed_cnt} value(s) that could not be parsed as {ckind}",
                    "columns": [clabel],
                    "what_it_means": (
                        f"Although '{clabel}' was detected as a {ckind} column, {unparsed_cnt} non-empty cell(s) contained text or invalid formatting "
                        "and are excluded from mathematical calculations."
                    ),
                    "action_effect": "Reporting only — unparseable values remain visible in the Data Preview table but are skipped during math aggregation.",
                }
            )

        if c.get("date_format_note"):
            warnings.append(
                {
                    "severity": "info",
                    "title": f"Date parsing note on '{clabel}'",
                    "columns": [clabel],
                    "what_it_means": c["date_format_note"],
                    "action_effect": "Dates were normalized to ISO YYYY-MM-DD during import.",
                }
            )

        if is_id:
            warnings.append(
                {
                    "severity": "info",
                    "title": f"'{clabel}' detected as an Identifier (ID) column",
                    "columns": [clabel],
                    "what_it_means": (
                        f"{c.get('id_reason') or 'Values act as unique record keys.'} "
                        "DataPrism excludes identifier columns from numeric averages/sums and category frequency charts so IDs do not clutter your analysis."
                    ),
                    "action_effect": "Reporting only — column remains searchable in Data Preview and usable for counting unique entities.",
                }
            )

        if len(non_null) > 1 and distinct_cnt == 1:
            warnings.append(
                {
                    "severity": "info",
                    "title": f"Column '{clabel}' has only 1 unique value ('{samples[0] if samples else ''}')",
                    "columns": [clabel],
                    "what_it_means": "Every non-empty row has the exact same value, so this column cannot explain differences or correlations across records.",
                    "action_effect": "Reporting only.",
                }
            )

    missing_pct_total = round((missing_cells / max(1, total_cells)) * 100, 1)
    complete_rows_pct = round((complete_rows / max(1, row_count)) * 100, 1)

    return {
        "row_count": row_count,
        "column_count": col_count,
        "total_cells": total_cells,
        "missing_cells": missing_cells,
        "missing_pct": missing_pct_total,
        "complete_rows": complete_rows,
        "complete_rows_pct": complete_rows_pct,
        "duplicate_rows": duplicate_rows,
        "duplicate_pct": round((duplicate_rows / max(1, row_count)) * 100, 1),
        "encoding": diagnostics.get("encoding", "utf-8"),
        "delimiter": diagnostics.get("delimiter", ","),
        "columns": col_reports,
        "warnings": warnings,
        "explanation": {
            "what_calculated": "Row and column counts, missing cell counts, exact full-row duplicates, inferred column types, and parsing diagnostics.",
            "columns_used": [c.get("label", c["name"]) for c in cols],
            "how_calculated": (
                "Each cell is checked against empty/null tokens ('', 'null', 'N/A', 'NaN'). Column types are inferred when >=75% of non-empty values parse as numbers or dates, "
                "while identifier columns are flagged via naming patterns (e.g., _id, sku) or near-100% uniqueness."
            ),
            "what_it_means": (
                f"Your dataset has {row_count:,} rows across {col_count} columns with {complete_rows_pct}% complete rows "
                f"and {duplicate_rows:,} duplicate row(s)."
            ),
            "assumptions_limitations": (
                "Health diagnostics are non-destructive: viewing warnings never changes your data. "
                "If you run a Data Cleaning action (such as removing duplicates or filling missing values), DataPrism saves a new derived dataset copy so your original upload stays untouched."
            ),
        },
    }


def build_numeric_summary(
    rows: list[dict[str, Any]], cols: list[dict[str, Any]], custom_units: dict[str, str] | None = None
) -> dict[str, Any]:
    """Section 2: Numeric Summary with count, missing, mean, median, std dev, variance, min, Q1, Q3, max, and unit labels."""
    custom_units = custom_units or {}
    row_count = len(rows)
    numeric_cols = [c for c in cols if c["kind"] == "number" and not c.get("is_identifier")]
    excluded_id_cols = [c.get("label", c["name"]) for c in cols if c.get("is_identifier") or c["kind"] == "id"]

    summaries: list[dict[str, Any]] = []
    for c in numeric_cols:
        cname = c["name"]
        clabel = c.get("label", cname)
        unit = (custom_units.get(cname) or "").strip() or "Unknown (not specified)"

        vals: list[float] = []
        for r in rows:
            num = to_number(r.get(cname))
            if num is not None:
                vals.append(num)

        valid_cnt = len(vals)
        missing_cnt = row_count - valid_cnt
        if valid_cnt == 0:
            continue

        sorted_vals = sorted(vals)
        min_v = sorted_vals[0]
        max_v = sorted_vals[-1]
        mean_v = float(statistics.fmean(sorted_vals))
        q1_v, med_v, q3_v = compute_quantiles(sorted_vals)
        iqr_v = q3_v - q1_v
        var_v = float(statistics.variance(sorted_vals)) if valid_cnt >= 2 else 0.0
        std_v = float(statistics.stdev(sorted_vals)) if valid_cnt >= 2 else 0.0
        sum_v = float(sum(sorted_vals))

        low_fence = q1_v - 1.5 * iqr_v
        high_fence = q3_v + 1.5 * iqr_v
        outliers = [v for v in sorted_vals if v < low_fence or v > high_fence] if iqr_v > 0 else []

        diff_ratio = abs(mean_v - med_v) / max(abs(med_v), std_v, 1e-9)
        if len(outliers) > 0 and diff_ratio > 0.15:
            direction = "higher" if mean_v > med_v else "lower"
            interpretation = (
                f"The mean ({mean_v:,.2f}) is noticeably {direction} than the median ({med_v:,.2f}), "
                f"and {len(outliers)} extreme value(s) fall outside the 1.5x IQR range. "
                "The median and quartiles (Q1-Q3) give a steadier picture of a typical row than the mean or variance."
            )
        else:
            interpretation = (
                f"The mean ({mean_v:,.2f}) and median ({med_v:,.2f}) are close together, suggesting a balanced distribution without heavy skew."
            )

        summaries.append(
            {
                "name": cname,
                "label": clabel,
                "unit": unit,
                "count": valid_cnt,
                "missing_count": missing_cnt,
                "mean": round(mean_v, 4),
                "median": round(med_v, 4),
                "std_dev": round(std_v, 4),
                "variance": round(var_v, 4),
                "min": round(min_v, 4),
                "q1": round(q1_v, 4),
                "q3": round(q3_v, 4),
                "iqr": round(iqr_v, 4),
                "max": round(max_v, 4),
                "sum": round(sum_v, 4),
                "outlier_count": len(outliers),
                "interpretation": interpretation,
            }
        )

    return {
        "columns": summaries,
        "excluded_identifiers": excluded_id_cols,
        "explanation": {
            "what_calculated": (
                "For each quantitative numeric column: non-missing count, missing count, arithmetic mean, median (50th percentile), "
                "sample standard deviation (s), sample variance (s²), minimum, 25th percentile (Q1), 75th percentile (Q3), and maximum."
            ),
            "columns_used": [s["label"] for s in summaries],
            "how_calculated": (
                "Mean = Σx / n; Sample Variance (s²) = Σ(x - mean)² / (n - 1); Standard Deviation (s) = √s²; "
                "Median and Quartiles (Q1, Q3) sort valid numbers from lowest to highest and split them into four equal quarters."
            ),
            "what_it_means": (
                "Compare Mean vs. Median to check for skew, and compare Standard Deviation / Variance vs. the Q1–Q3 range to see how spread out values are."
            ),
            "assumptions_limitations": (
                "Units are labeled 'Unknown (not specified)' unless you set a unit label above. "
                "Mean, standard deviation, and especially variance (which squares differences from the mean) can be strongly pulled by extreme high or low outliers; "
                "when outliers are present, rely on the Median and Q1–Q3 quartiles for a typical value."
            ),
        },
    }


def build_category_breakdown(rows: list[dict[str, Any]], cols: list[dict[str, Any]]) -> dict[str, Any]:
    """Section 3: Category Breakdown showing counts, percentages, distinct count, top values, excluding IDs."""
    row_count = len(rows)
    cat_cols = [c for c in cols if c["kind"] == "category" and not c.get("is_identifier")]
    excluded_ids = [
        {"name": c["name"], "label": c.get("label", c["name"]), "reason": c.get("id_reason") or "Detected as unique identifier."}
        for c in cols
        if c.get("is_identifier") or c["kind"] == "id"
    ]

    breakdowns: list[dict[str, Any]] = []
    for c in cat_cols:
        cname = c["name"]
        clabel = c.get("label", cname)
        counter: Counter[str] = Counter()
        missing_cnt = 0
        for r in rows:
            val = r.get(cname)
            if is_missing(val):
                missing_cnt += 1
            else:
                counter[str(val).strip()] += 1

        valid_cnt = row_count - missing_cnt
        distinct_cnt = len(counter)
        if valid_cnt == 0:
            continue

        most_common = counter.most_common(8)
        top_items: list[dict[str, Any]] = []
        shown_count = 0
        for val_str, cnt in most_common:
            shown_count += cnt
            top_items.append(
                {
                    "value": val_str,
                    "count": cnt,
                    "pct_of_total": round((cnt / max(1, row_count)) * 100, 1),
                    "pct_of_valid": round((cnt / max(1, valid_cnt)) * 100, 1),
                }
            )

        remaining_cnt = valid_cnt - shown_count
        remaining_cats = distinct_cnt - len(most_common)
        if remaining_cnt > 0:
            top_items.append(
                {
                    "value": f"Other ({remaining_cats} more categories)",
                    "count": remaining_cnt,
                    "pct_of_total": round((remaining_cnt / max(1, row_count)) * 100, 1),
                    "pct_of_valid": round((remaining_cnt / max(1, valid_cnt)) * 100, 1),
                    "is_other": True,
                }
            )

        top_val, top_cnt = most_common[0]
        top_share = round((top_cnt / max(1, valid_cnt)) * 100, 1)
        if top_share >= 50:
            meaning = f"'{top_val}' dominates this column, accounting for {top_share}% ({top_cnt:,} of {valid_cnt:,}) of non-empty records."
        elif distinct_cnt > 15:
            meaning = f"Values are spread across {distinct_cnt} distinct categories; the most frequent is '{top_val}' at {top_share}%."
        else:
            meaning = f"Most common category is '{top_val}' ({top_cnt:,} rows, {top_share}% of valid rows) across {distinct_cnt} distinct categories."

        breakdowns.append(
            {
                "name": cname,
                "label": clabel,
                "valid_count": valid_cnt,
                "missing_count": missing_cnt,
                "missing_pct": round((missing_cnt / max(1, row_count)) * 100, 1),
                "distinct_count": distinct_cnt,
                "top_value": top_val,
                "top_count": top_cnt,
                "top_pct_of_valid": top_share,
                "frequencies": top_items,
                "interpretation": meaning,
            }
        )

    return {
        "columns": breakdowns,
        "excluded_identifiers": excluded_ids,
        "explanation": {
            "what_calculated": "Category row counts, percentage share of valid rows, percentage share of all dataset rows, number of distinct categories, and most frequent values.",
            "columns_used": [b["label"] for b in breakdowns],
            "how_calculated": (
                "Rows are grouped by exact trimmed text value. Share of valid rows = (category count / non-empty rows) × 100; "
                "Share of total rows = (category count / total dataset rows) × 100."
            ),
            "what_it_means": "Shows which groups or segments appear most frequently in your dataset and whether records are concentrated in one category or balanced across many.",
            "assumptions_limitations": (
                "Identifier columns (such as Order ID, Customer ID, or SKU) are intentionally excluded because almost every row has a unique ID. "
                "Text matching is case- and spelling-sensitive after trimming whitespace (e.g., 'Online' vs. 'online Store' are counted as separate categories)."
            ),
        },
    }


def bucket_date(iso_date: str, grain: str) -> str:
    """Group YYYY-MM-DD into day, week (ISO Monday date), month (YYYY-MM), quarter (YYYY-Q#), or year (YYYY)."""
    try:
        dt = datetime.strptime(iso_date[:10], "%Y-%m-%d")
    except ValueError:
        return iso_date[:10]
    if grain == "day":
        return dt.strftime("%Y-%m-%d")
    if grain == "week":
        monday = dt - timedelta(days=dt.weekday())
        return monday.strftime("%Y-%m-%d")
    if grain == "quarter":
        q = (dt.month - 1) // 3 + 1
        return f"{dt.year}-Q{q}"
    if grain == "year":
        return f"{dt.year}"
    return dt.strftime("%Y-%m")


def aggregate_values(vals: list[float], agg: str) -> float:
    if not vals:
        return 0.0
    if agg == "count":
        return float(len(vals))
    if agg == "mean":
        return float(statistics.fmean(vals))
    if agg == "median":
        return float(statistics.median(vals))
    if agg == "min":
        return float(min(vals))
    if agg == "max":
        return float(max(vals))
    return float(sum(vals))


def build_time_trends(
    rows: list[dict[str, Any]],
    cols: list[dict[str, Any]],
    date_col: str | None = None,
    measure_col: str | None = None,
    grain: str | None = None,
    agg: str | None = None,
) -> dict[str, Any]:
    """Section 4: Trends Over Time with user-selectable date column, numeric measure, time grouping, and aggregation."""
    date_cols = [{"name": c["name"], "label": c.get("label", c["name"])} for c in cols if c["kind"] == "date"]
    measure_cols = [
        {"name": c["name"], "label": c.get("label", c["name"])}
        for c in cols
        if c["kind"] == "number" and not c.get("is_identifier")
    ]

    if not date_cols:
        return {
            "available": False,
            "reason": "No valid date column was detected in this dataset. Trends Over Time requires at least one date column (e.g., Order Date, Signup Date).",
            "date_columns": [],
            "measure_columns": measure_cols,
            "selected": None,
            "series": [],
            "explanation": {
                "what_calculated": "No time trend calculated because the dataset has no date column.",
                "columns_used": [],
                "how_calculated": "Requires a column where >=75% of non-empty values parse as calendar dates.",
                "what_it_means": "DataPrism does not silently treat arbitrary numbers or text as dates.",
                "assumptions_limitations": "Upload a dataset with a date column (formatted as YYYY-MM-DD, MM/DD/YYYY, or DD/MM/YYYY) to analyze trends over time.",
            },
        }

    sel_date = date_col if any(d["name"] == date_col for d in date_cols) else date_cols[0]["name"]
    sel_date_label = next((d["label"] for d in date_cols if d["name"] == sel_date), sel_date)

    if measure_col == "__count__" or not measure_cols:
        sel_measure = "__count__"
        sel_measure_label = "Record Count (rows)"
    else:
        sel_measure = measure_col if any(m["name"] == measure_col for m in measure_cols) else measure_cols[0]["name"]
        sel_measure_label = next((m["label"] for m in measure_cols if m["name"] == sel_measure), sel_measure)

    sel_agg = agg if agg in {"sum", "mean", "median", "count", "min", "max"} else ("count" if sel_measure == "__count__" else "sum")
    if sel_measure == "__count__":
        sel_agg = "count"

    # Collect valid dates first to auto-select a sensible grain if not specified
    parsed_pairs: list[tuple[str, float]] = []
    skipped_missing_date = 0
    skipped_missing_measure = 0

    for r in rows:
        raw_d = r.get(sel_date)
        iso_d, _ = parse_date_value(raw_d)
        if not iso_d:
            skipped_missing_date += 1
            continue
        if sel_measure == "__count__":
            parsed_pairs.append((iso_d, 1.0))
        else:
            num_v = to_number(r.get(sel_measure))
            if num_v is None:
                skipped_missing_measure += 1
                continue
            parsed_pairs.append((iso_d, num_v))

    if not parsed_pairs:
        return {
            "available": False,
            "reason": f"No rows had both a valid date in '{sel_date_label}' and a valid numeric value in '{sel_measure_label}'.",
            "date_columns": date_cols,
            "measure_columns": measure_cols,
            "selected": {"date_col": sel_date, "measure_col": sel_measure, "grain": grain or "month", "agg": sel_agg},
            "series": [],
            "explanation": {
                "what_calculated": "No valid date-measure pairs found.",
                "columns_used": [sel_date_label, sel_measure_label],
                "how_calculated": "N/A",
                "what_it_means": "Check Dataset Health for missing or unparseable values in these columns.",
                "assumptions_limitations": "Both date and measure must be non-empty on the same row.",
            },
        }

    all_dates = sorted(p[0] for p in parsed_pairs)
    min_date, max_date = all_dates[0], all_dates[-1]
    try:
        span_days = (datetime.strptime(max_date, "%Y-%m-%d") - datetime.strptime(min_date, "%Y-%m-%d")).days
    except ValueError:
        span_days = 90

    if grain in {"day", "week", "month", "quarter", "year"}:
        sel_grain = grain
    else:
        if span_days <= 45:
            sel_grain = "day"
        elif span_days <= 180:
            sel_grain = "week"
        else:
            sel_grain = "month"

    buckets: dict[str, list[float]] = defaultdict(list)
    for iso_d, val in parsed_pairs:
        b_key = bucket_date(iso_d, sel_grain)
        buckets[b_key].append(val)

    series: list[dict[str, Any]] = []
    for period_key in sorted(buckets.keys()):
        b_vals = buckets[period_key]
        agg_val = aggregate_values(b_vals, sel_agg)
        series.append(
            {
                "period": period_key,
                "value": round(agg_val, 4),
                "row_count": len(b_vals),
            }
        )

    first_val = series[0]["value"]
    last_val = series[-1]["value"]
    peak_item = max(series, key=lambda x: x["value"])
    low_item = min(series, key=lambda x: x["value"])
    pct_change = round(((last_val - first_val) / abs(first_val)) * 100, 1) if first_val != 0 else None

    grain_labels = {
        "day": "Daily (YYYY-MM-DD)",
        "week": "Weekly (Monday start date)",
        "month": "Monthly (YYYY-MM)",
        "quarter": "Quarterly (YYYY-Q#)",
        "year": "Yearly (YYYY)",
    }
    agg_labels = {
        "sum": "Sum (total across rows in period)",
        "mean": "Average (arithmetic mean per row in period)",
        "median": "Median (middle value in period)",
        "count": "Count of matching records in period",
        "min": "Minimum value in period",
        "max": "Maximum value in period",
    }

    if len(series) == 1:
        trend_meaning = f"All records fall into a single {sel_grain} period ({series[0]['period']}) with {sel_agg} = {first_val:,.2f}. Try switching Time Grouping to a finer grain (like Daily or Weekly)."
    else:
        dir_word = "increased" if (pct_change or 0) > 0 else ("decreased" if (pct_change or 0) < 0 else "remained flat")
        pct_str = f" ({pct_change:+.1f}%)" if pct_change is not None else ""
        trend_meaning = (
            f"Across {len(series)} {sel_grain} periods from {min_date} to {max_date}, {sel_measure_label} ({sel_agg}) "
            f"{dir_word} from {first_val:,.2f} in {series[0]['period']} to {last_val:,.2f} in {series[-1]['period']}{pct_str}. "
            f"Highest period: {peak_item['period']} ({peak_item['value']:,.2f}); lowest: {low_item['period']} ({low_item['value']:,.2f})."
        )

    return {
        "available": True,
        "date_columns": date_cols,
        "measure_columns": measure_cols,
        "selected": {
            "date_col": sel_date,
            "date_label": sel_date_label,
            "measure_col": sel_measure,
            "measure_label": sel_measure_label,
            "grain": sel_grain,
            "grain_label": grain_labels.get(sel_grain, sel_grain),
            "agg": sel_agg,
            "agg_label": agg_labels.get(sel_agg, sel_agg),
        },
        "stats": {
            "start_date": min_date,
            "end_date": max_date,
            "span_days": span_days,
            "periods_count": len(series),
            "rows_used": len(parsed_pairs),
            "skipped_missing_date": skipped_missing_date,
            "skipped_missing_measure": skipped_missing_measure,
            "first_period": series[0]["period"],
            "first_value": first_val,
            "last_period": series[-1]["period"],
            "last_value": last_val,
            "change_pct": pct_change,
            "peak_period": peak_item["period"],
            "peak_value": peak_item["value"],
            "low_period": low_item["period"],
            "low_value": low_item["value"],
        },
        "series": series,
        "explanation": {
            "what_calculated": f"{agg_labels.get(sel_agg, sel_agg)} of '{sel_measure_label}' grouped by {grain_labels.get(sel_grain, sel_grain)} using '{sel_date_label}'.",
            "columns_used": [sel_date_label] if sel_measure == "__count__" else [sel_date_label, sel_measure_label],
            "how_calculated": (
                f"Each row's '{sel_date_label}' is normalized to ISO YYYY-MM-DD and assigned to its {sel_grain} bucket. "
                f"Within each bucket, '{sel_measure_label}' values are combined using {sel_agg.upper()} across {len(parsed_pairs):,} valid rows."
            ),
            "what_it_means": trend_meaning,
            "assumptions_limitations": (
                f"Skipped {skipped_missing_date} row(s) with missing/invalid dates and {skipped_missing_measure} row(s) with missing measure values. "
                "Time buckets with zero records in the dataset are not interpolated with synthetic zeros, and the first or last period may represent a partial week/month."
            ),
        },
    }


def pearson_r(xs: list[float], ys: list[float]) -> float | None:
    n = len(xs)
    if n < 3:
        return None
    mean_x = statistics.fmean(xs)
    mean_y = statistics.fmean(ys)
    num = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
    den_x = sum((x - mean_x) ** 2 for x in xs)
    den_y = sum((y - mean_y) ** 2 for y in ys)
    if den_x <= 1e-12 or den_y <= 1e-12:
        return None
    r = num / math.sqrt(den_x * den_y)
    return max(-1.0, min(1.0, float(r)))


def interpret_correlation(r: float | None, n: int) -> tuple[str, str]:
    if r is None:
        return "Unsuitable / Constant", "Cannot compute correlation because one column has zero variance or fewer than 3 paired observations."
    abs_r = abs(r)
    direction = "positive" if r > 0 else ("negative" if r < 0 else "flat")
    if abs_r < 0.15:
        strength = "No meaningful linear association"
        desc = f"r = {r:+.3f} across n = {n:,} paired rows indicates virtually no linear relationship between these two columns."
    elif abs_r < 0.35:
        strength = f"Weak {direction} linear association"
        desc = f"r = {r:+.3f} (n = {n:,}) suggests a weak {direction} tendency, with substantial scatter around the trend."
    elif abs_r < 0.65:
        strength = f"Moderate {direction} linear association"
        desc = f"r = {r:+.3f} (n = {n:,}) shows a moderate {direction} linear pattern where higher X values tend to align with {'higher' if r > 0 else 'lower'} Y values."
    else:
        strength = f"Strong {direction} linear association"
        desc = f"r = {r:+.3f} (n = {n:,}) shows a strong {direction} linear relationship in this sample."
    return strength, desc


def build_relationships(
    rows: list[dict[str, Any]],
    cols: list[dict[str, Any]],
    x_col: str | None = None,
    y_col: str | None = None,
    show_corr: bool = True,
) -> dict[str, Any]:
    """Section 5: Relationships scatter plot and optional Pearson correlation with suitability checks."""
    num_cols = [
        {"name": c["name"], "label": c.get("label", c["name"])}
        for c in cols
        if c["kind"] == "number" and not c.get("is_identifier")
    ]

    if len(num_cols) < 2:
        return {
            "available": False,
            "reason": (
                f"Relationships analysis requires at least 2 quantitative numeric columns (excluding IDs), "
                f"but this dataset has {len(num_cols)}."
            ),
            "numeric_columns": num_cols,
            "selected": None,
            "points": [],
            "correlation": {"suitable": False, "r": None, "reason": "Fewer than 2 numeric measure columns exist."},
            "top_pairs": [],
            "explanation": {
                "what_calculated": "No scatter plot or correlation calculated because fewer than two numeric measure columns exist.",
                "columns_used": [c["label"] for c in num_cols],
                "how_calculated": "Requires two quantitative numeric columns with at least 3 paired numeric rows.",
                "what_it_means": "DataPrism does not calculate correlation on categorical text or identifier codes.",
                "assumptions_limitations": "Correlation measures linear association between two numeric variables and never proves causation.",
            },
        }

    # Compute top pairs across numeric columns (excluding IDs)
    top_pairs: list[dict[str, Any]] = []
    for i in range(len(num_cols)):
        for j in range(i + 1, len(num_cols)):
            c1, c2 = num_cols[i], num_cols[j]
            xs_pair: list[float] = []
            ys_pair: list[float] = []
            for r in rows:
                v1 = to_number(r.get(c1["name"]))
                v2 = to_number(r.get(c2["name"]))
                if v1 is not None and v2 is not None:
                    xs_pair.append(v1)
                    ys_pair.append(v2)
            r_val = pearson_r(xs_pair, ys_pair)
            if r_val is not None:
                strength, _ = interpret_correlation(r_val, len(xs_pair))
                top_pairs.append(
                    {
                        "x_col": c1["name"],
                        "x_label": c1["label"],
                        "y_col": c2["name"],
                        "y_label": c2["label"],
                        "r": round(r_val, 3),
                        "abs_r": round(abs(r_val), 3),
                        "paired_count": len(xs_pair),
                        "strength": strength,
                    }
                )
    top_pairs.sort(key=lambda item: item["abs_r"], reverse=True)

    sel_x = x_col if any(c["name"] == x_col for c in num_cols) else num_cols[0]["name"]
    if y_col and any(c["name"] == y_col for c in num_cols):
        sel_y = y_col
    else:
        sel_y = next((c["name"] for c in num_cols if c["name"] != sel_x), num_cols[1]["name"])

    sel_x_label = next((c["label"] for c in num_cols if c["name"] == sel_x), sel_x)
    sel_y_label = next((c["label"] for c in num_cols if c["name"] == sel_y), sel_y)

    xs: list[float] = []
    ys: list[float] = []
    points: list[dict[str, Any]] = []
    skipped_rows = 0

    for idx, r in enumerate(rows, start=1):
        vx = to_number(r.get(sel_x))
        vy = to_number(r.get(sel_y))
        if vx is None or vy is None:
            skipped_rows += 1
            continue
        xs.append(vx)
        ys.append(vy)
        if len(points) < 400:
            points.append({"x": round(vx, 4), "y": round(vy, 4), "row": idx})

    n_paired = len(xs)
    if sel_x == sel_y:
        corr_info = {
            "suitable": False,
            "r": None,
            "r_squared": None,
            "strength": "Same Column Selected",
            "interpretation": "Select two different numeric columns to evaluate their relationship.",
            "reason": "X and Y are set to the exact same column (a column always has a trivial r = 1.00 with itself).",
        }
    elif n_paired < 3:
        corr_info = {
            "suitable": False,
            "r": None,
            "r_squared": None,
            "strength": "Insufficient Paired Data",
            "interpretation": f"Only {n_paired} row(s) have valid numbers in both '{sel_x_label}' and '{sel_y_label}'. At least 3 paired rows are required.",
            "reason": "Fewer than 3 paired numeric rows available.",
        }
    else:
        r_val = pearson_r(xs, ys)
        if r_val is None:
            corr_info = {
                "suitable": False,
                "r": None,
                "r_squared": None,
                "strength": "Zero Variance (Constant Column)",
                "interpretation": "One of the selected columns has the exact same value on every paired row, so correlation is undefined (division by zero standard deviation).",
                "reason": "Zero variance in X or Y.",
            }
        else:
            strength, desc = interpret_correlation(r_val, n_paired)
            sample_warn = " (Note: sample size n < 10 is very small, so treat this value cautiously.)" if n_paired < 10 else ""
            corr_info = {
                "suitable": True,
                "r": round(r_val, 4),
                "r_squared": round(r_val**2, 4),
                "strength": strength,
                "interpretation": desc + sample_warn,
                "reason": None,
            }

    return {
        "available": True,
        "numeric_columns": num_cols,
        "selected": {
            "x_col": sel_x,
            "x_label": sel_x_label,
            "y_col": sel_y,
            "y_label": sel_y_label,
            "show_correlation": show_corr,
        },
        "paired_count": n_paired,
        "skipped_rows": skipped_rows,
        "points": points,
        "correlation": corr_info,
        "top_pairs": top_pairs[:8],
        "explanation": {
            "what_calculated": (
                f"Paired scatter coordinates ({sel_x_label} vs. {sel_y_label}) across {n_paired:,} complete numeric rows"
                + (", plus Pearson linear correlation coefficient (r)." if show_corr else ".")
            ),
            "columns_used": [sel_x_label, sel_y_label],
            "how_calculated": (
                "Each point represents one row with valid numeric values in both columns. "
                "Pearson r = Σ[(x - x̄)(y - ȳ)] / √[Σ(x - x̄)² · Σ(y - ȳ)²], ranging from -1.0 (perfect inverse linear trend) to +1.0 (perfect direct linear trend)."
            ),
            "what_it_means": corr_info["interpretation"],
            "assumptions_limitations": (
                "Correlation describes linear association only and DOES NOT prove causation. "
                "A high correlation can be caused by a third confounding variable, shared time trends, or a few extreme outliers; "
                "a near-zero correlation can still hide a curved (non-linear) relationship, which is why inspecting the scatter plot is essential."
            ),
        },
    }


def resolve_category_roles(
    cols: list[dict[str, Any]], category: str, role_overrides: dict[str, str] | None = None
) -> dict[str, dict[str, Any]]:
    """Map domain roles to dataset columns using user overrides first, then conservative name & type matching."""
    role_overrides = role_overrides or {}
    cat_spec = CATEGORIES.get(category, CATEGORIES["general"])
    role_defs = cat_spec.get("roles", {})
    col_by_name = {c["name"]: c for c in cols}

    resolved: dict[str, dict[str, Any]] = {}
    used_cols: set[str] = set()

    for role_key, rdef in role_defs.items():
        override_col = role_overrides.get(role_key)
        if override_col == "__none__":
            resolved[role_key] = {
                "role": role_key,
                "role_label": rdef["label"],
                "expected_kind": rdef["kind"],
                "column": None,
                "column_label": None,
                "status": "cleared_by_user",
                "note": "Manually set to 'None / Not Present' by user.",
            }
            continue
        if override_col and override_col in col_by_name:
            cmeta = col_by_name[override_col]
            resolved[role_key] = {
                "role": role_key,
                "role_label": rdef["label"],
                "expected_kind": rdef["kind"],
                "column": cmeta["name"],
                "column_label": cmeta.get("label", cmeta["name"]),
                "status": "user_selected",
                "note": f"Mapped to '{cmeta.get('label', cmeta['name'])}' (user selection).",
            }
            used_cols.add(cmeta["name"])
            continue

        # Automatic conservative matching: require both alias match AND compatible type
        expected_kind = rdef["kind"]
        aliases = [a.lower() for a in rdef.get("aliases", [])]
        matched_col = None

        for c in cols:
            if c["name"] in used_cols:
                continue
            cname_low = c["name"].lower()
            clabel_low = re.sub(r"[^a-z0-9]+", "_", c.get("label", "").lower()).strip("_")
            kind_ok = (
                expected_kind == "any"
                or c["kind"] == expected_kind
                or (expected_kind == "category" and c["kind"] in {"category", "text"} and not c.get("is_identifier"))
            )
            if not kind_ok:
                continue
            if expected_kind == "number" and c.get("is_identifier"):
                continue
            if any(a == cname_low or a == clabel_low for a in aliases) or any(
                len(a) >= 4 and (a in cname_low or a in clabel_low) for a in aliases
            ):
                matched_col = c
                break

        if matched_col:
            used_cols.add(matched_col["name"])
            resolved[role_key] = {
                "role": role_key,
                "role_label": rdef["label"],
                "expected_kind": expected_kind,
                "column": matched_col["name"],
                "column_label": matched_col.get("label", matched_col["name"]),
                "status": "auto_matched",
                "note": f"Auto-matched to '{matched_col.get('label', matched_col['name'])}' by column name and type.",
            }
        else:
            resolved[role_key] = {
                "role": role_key,
                "role_label": rdef["label"],
                "expected_kind": expected_kind,
                "column": None,
                "column_label": None,
                "status": "missing",
                "note": f"No matching {expected_kind} column found for {rdef['label']}.",
            }

    return resolved


def group_sum_and_count(
    rows: list[dict[str, Any]], group_col: str, measure_col: str | None = None, limit: int = 8
) -> list[dict[str, Any]]:
    totals: dict[str, float] = defaultdict(float)
    counts: dict[str, int] = defaultdict(int)
    for r in rows:
        g_val = r.get(group_col)
        if is_missing(g_val):
            continue
        g_key = str(g_val).strip()
        if measure_col:
            m_val = to_number(r.get(measure_col))
            if m_val is None:
                continue
            totals[g_key] += m_val
            counts[g_key] += 1
        else:
            totals[g_key] += 1.0
            counts[g_key] += 1

    grand_total = sum(totals.values()) or 1.0
    ranked = sorted(totals.items(), key=lambda kv: kv[1], reverse=True)[:limit]
    return [
        {
            "group": k,
            "total": round(v, 2),
            "count": counts[k],
            "mean": round(v / max(1, counts[k]), 2),
            "share_pct": round((v / grand_total) * 100, 1),
        }
        for k, v in ranked
    ]


def build_category_analysis(
    rows: list[dict[str, Any]],
    cols: list[dict[str, Any]],
    category: str,
    role_overrides: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Section 6: Category-Specific Analysis with visible/editable assumptions and no invented metrics."""
    cat_key = category if category in CATEGORIES else "general"
    cat_spec = CATEGORIES[cat_key]
    roles = resolve_category_roles(cols, cat_key, role_overrides)
    analyses: list[dict[str, Any]] = []
    skipped_analyses: list[dict[str, Any]] = []

    def col_id(rkey: str) -> str | None:
        return roles.get(rkey, {}).get("column")

    def col_lbl(rkey: str) -> str:
        return roles.get(rkey, {}).get("column_label") or rkey

    if cat_key == "ecommerce":
        rev_c, prod_c, cat_c, cost_c, prof_c, date_c = (
            col_id("revenue"),
            col_id("product"),
            col_id("category"),
            col_id("cost"),
            col_id("profit"),
            col_id("date"),
        )
        if rev_c and (prod_c or cat_c):
            grp_c = prod_c or cat_c
            grp_lbl = col_lbl("product") if prod_c else col_lbl("category")
            rev_lbl = col_lbl("revenue")
            items = group_sum_and_count(rows, grp_c, rev_c)
            total_rev = sum(to_number(r.get(rev_c)) or 0.0 for r in rows)
            analyses.append(
                {
                    "title": f"Revenue & Order Totals by {grp_lbl}",
                    "status": "ready",
                    "what_calculated": f"Total '{rev_lbl}', record count, average per record, and share of revenue grouped by '{grp_lbl}'.",
                    "columns_used": [grp_lbl, rev_lbl],
                    "how_calculated": f"Grouped rows by '{grp_lbl}' and summed valid numbers in '{rev_lbl}' (total dataset revenue = {total_rev:,.2f}).",
                    "what_it_means": (
                        f"Top contributor is '{items[0]['group']}' with {items[0]['total']:,.2f} ({items[0]['share_pct']}% of total) across {items[0]['count']} records."
                        if items
                        else "No valid grouped rows found."
                    ),
                    "assumptions_limitations": "Assumes each row represents a distinct line item or order and that revenue values share the same currency.",
                    "table_headers": [grp_lbl, "Records", f"Total {rev_lbl}", f"Avg {rev_lbl}", "Share %"],
                    "table_rows": [[x["group"], x["count"], f"{x['total']:,.2f}", f"{x['mean']:,.2f}", f"{x['share_pct']}%"] for x in items],
                }
            )
        else:
            skipped_analyses.append(
                {
                    "title": "Product / Category Revenue Totals",
                    "missing_fields": [k for k, v in [("Revenue / Sales column", rev_c), ("Product or Category column", prod_c or cat_c)] if not v],
                    "explanation": "Skipped because calculating revenue totals by product or channel requires both a numeric revenue column and a product/category grouping column.",
                }
            )

        # Profit analysis ONLY if profit or cost exists
        if prof_c or (rev_c and cost_c):
            if prof_c:
                prof_lbl = col_lbl("profit")
                prof_vals = [to_number(r.get(prof_c)) for r in rows]
                prof_clean = [v for v in prof_vals if v is not None]
                total_prof = sum(prof_clean)
                neg_cnt = sum(1 for v in prof_clean if v < 0)
                analyses.append(
                    {
                        "title": f"Profit Summary (Direct '{prof_lbl}' Field)",
                        "status": "ready",
                        "what_calculated": f"Total profit, average profit per record, and loss-making record count using '{prof_lbl}'.",
                        "columns_used": [prof_lbl] + ([col_lbl("product")] if prod_c else []),
                        "how_calculated": f"Summed valid numeric values directly from '{prof_lbl}' across {len(prof_clean):,} rows.",
                        "what_it_means": f"Total recorded profit is {total_prof:,.2f} (average {total_prof/max(1,len(prof_clean)):,.2f} per record), with {neg_cnt} loss-making (< 0) record(s).",
                        "assumptions_limitations": f"Uses the existing '{prof_lbl}' column directly without subtracting additional overhead or tax.",
                        "table_headers": ["Metric", "Value"],
                        "table_rows": [
                            ["Valid Profit Rows", f"{len(prof_clean):,}"],
                            ["Total Profit", f"{total_prof:,.2f}"],
                            ["Average Profit / Row", f"{total_prof / max(1, len(prof_clean)):,.2f}"],
                            ["Loss-Making Rows (< 0)", f"{neg_cnt:,}"],
                        ],
                    }
                )
            elif rev_c and cost_c:
                rev_lbl, cost_lbl = col_lbl("revenue"), col_lbl("cost")
                paired_margins: list[tuple[float, float, float]] = []
                for r in rows:
                    rv = to_number(r.get(rev_c))
                    cv = to_number(r.get(cost_c))
                    if rv is not None and cv is not None:
                        paired_margins.append((rv, cv, rv - cv))
                tot_r = sum(p[0] for p in paired_margins)
                tot_c = sum(p[1] for p in paired_margins)
                tot_p = sum(p[2] for p in paired_margins)
                margin_pct = (tot_p / tot_r * 100) if tot_r != 0 else 0.0
                analyses.append(
                    {
                        "title": "Gross Profit & Margin Analysis (Revenue minus Cost)",
                        "status": "ready",
                        "what_calculated": f"Gross Profit ({rev_lbl} - {cost_lbl}) and Gross Margin % across {len(paired_margins):,} paired rows.",
                        "columns_used": [rev_lbl, cost_lbl],
                        "how_calculated": f"Gross Profit = Σ({rev_lbl}) - Σ({cost_lbl}); Gross Margin % = (Gross Profit / Total {rev_lbl}) × 100.",
                        "what_it_means": f"Total revenue of {tot_r:,.2f} minus cost of {tot_c:,.2f} yields {tot_p:,.2f} gross profit ({margin_pct:.1f}% gross margin).",
                        "assumptions_limitations": f"Assumes '{cost_lbl}' is recorded on the same total-row basis as '{rev_lbl}' (not unit cost requiring multiplication by quantity).",
                        "table_headers": ["Component", "Amount"],
                        "table_rows": [
                            [f"Total {rev_lbl}", f"{tot_r:,.2f}"],
                            [f"Total {cost_lbl}", f"{tot_c:,.2f}"],
                            ["Calculated Gross Profit", f"{tot_p:,.2f}"],
                            ["Gross Margin %", f"{margin_pct:.1f}%"],
                        ],
                    }
                )
        else:
            skipped_analyses.append(
                {
                    "title": "Profit & Margin Analysis",
                    "missing_fields": ["Profit column or Cost/COGS column"],
                    "explanation": (
                        "Skipped because neither a Profit column nor a Cost/COGS column exists in this dataset. "
                        "DataPrism never invents profit margins from revenue alone. If your dataset has a cost or profit column under a custom name, select it in the Column Mapping controls above."
                    ),
                }
            )

        if date_c and rev_c:
            trend_res = build_time_trends(rows, cols, date_col=date_c, measure_col=rev_c, agg="sum")
            if trend_res.get("available") and trend_res.get("series"):
                s_list = trend_res["series"]
                analyses.append(
                    {
                        "title": f"Sales Revenue Trend Over Time ({trend_res['selected']['grain_label']})",
                        "status": "ready",
                        "what_calculated": trend_res["explanation"]["what_calculated"],
                        "columns_used": trend_res["explanation"]["columns_used"],
                        "how_calculated": trend_res["explanation"]["how_calculated"],
                        "what_it_means": trend_res["explanation"]["what_it_means"],
                        "assumptions_limitations": trend_res["explanation"]["assumptions_limitations"],
                        "table_headers": ["Period", "Orders / Rows", f"Total {col_lbl('revenue')}"],
                        "table_rows": [[p["period"], p["row_count"], f"{p['value']:,.2f}"] for p in s_list[-8:]],
                    }
                )

    elif cat_key == "retail":
        prod_c, cat_c, stock_c, reorder_c, sales_c = (
            col_id("product"),
            col_id("category"),
            col_id("stock"),
            col_id("reorder"),
            col_id("sales_qty"),
        )
        if stock_c and (prod_c or cat_c):
            grp_c = cat_c or prod_c
            grp_lbl = col_lbl("category") if cat_c else col_lbl("product")
            stk_lbl = col_lbl("stock")
            items = group_sum_and_count(rows, grp_c, stock_c)
            zero_stock = sum(1 for r in rows if (to_number(r.get(stock_c)) or 0) <= 0 and not is_missing(r.get(stock_c)))
            analyses.append(
                {
                    "title": f"Inventory Stock Summary by {grp_lbl}",
                    "status": "ready",
                    "what_calculated": f"Total on-hand units ('{stk_lbl}'), average units per record, and zero/out-of-stock count grouped by '{grp_lbl}'.",
                    "columns_used": [grp_lbl, stk_lbl],
                    "how_calculated": f"Summed valid numeric values in '{stk_lbl}' for each distinct '{grp_lbl}'.",
                    "what_it_means": f"'{items[0]['group']}' holds the largest stock volume ({items[0]['total']:,.0f} units, {items[0]['share_pct']}%). {zero_stock} row(s) have stock <= 0.",
                    "assumptions_limitations": "Assumes stock quantities represent current snapshot balances rather than cumulative historical counts.",
                    "table_headers": [grp_lbl, "Items / Rows", f"Total {stk_lbl}", f"Avg {stk_lbl}", "Share %"],
                    "table_rows": [[x["group"], x["count"], f"{x['total']:,.0f}", f"{x['mean']:,.1f}", f"{x['share_pct']}%"] for x in items],
                }
            )
        else:
            skipped_analyses.append(
                {
                    "title": "Inventory Stock Breakdown",
                    "missing_fields": [k for k, v in [("Stock / On-Hand column", stock_c), ("Product or Store/Category column", prod_c or cat_c)] if not v],
                    "explanation": "Skipped because inventory summaries require a numeric Stock/On-Hand column and a Product or Store/Category column.",
                }
            )

        if stock_c and reorder_c:
            stk_lbl, reo_lbl = col_lbl("stock"), col_lbl("reorder")
            below_cnt = 0
            checked_cnt = 0
            for r in rows:
                sv = to_number(r.get(stock_c))
                rv = to_number(r.get(reorder_c))
                if sv is not None and rv is not None:
                    checked_cnt += 1
                    if sv <= rv:
                        below_cnt += 1
            analyses.append(
                {
                    "title": "Reorder Threshold Alert Check",
                    "status": "ready",
                    "what_calculated": f"Records where '{stk_lbl}' is at or below '{reo_lbl}'.",
                    "columns_used": [stk_lbl, reo_lbl],
                    "how_calculated": f"Compared {stk_lbl} <= {reo_lbl} row-by-row across {checked_cnt:,} valid records.",
                    "what_it_means": f"{below_cnt:,} of {checked_cnt:,} records ({round(below_cnt/max(1,checked_cnt)*100,1)}%) are at or below their reorder point.",
                    "assumptions_limitations": "Does not account for in-transit purchase orders unless already included in on-hand stock.",
                    "table_headers": ["Check", "Count"],
                    "table_rows": [
                        ["Records Checked", f"{checked_cnt:,}"],
                        ["At or Below Reorder Point", f"{below_cnt:,}"],
                        ["Above Reorder Point", f"{checked_cnt - below_cnt:,}"],
                    ],
                }
            )
        else:
            skipped_analyses.append(
                {
                    "title": "Reorder Threshold Check",
                    "missing_fields": [k for k, v in [("Stock column", stock_c), ("Reorder Point column", reorder_c)] if not v],
                    "explanation": "Skipped because checking reorder risk requires both an On-Hand Stock column and a Reorder Point / Minimum Stock column.",
                }
            )

    elif cat_key == "education":
        grp_c, score_c, att_c, enr_c = col_id("group"), col_id("score"), col_id("attendance"), col_id("enrollment")
        if score_c:
            sc_lbl = col_lbl("score")
            sc_vals = sorted(v for v in (to_number(r.get(score_c)) for r in rows) if v is not None)
            if sc_vals:
                q1, med, q3 = compute_quantiles(sc_vals)
                mean_sc = statistics.fmean(sc_vals)
                table_rows = [
                    ["Students / Records Evaluated", f"{len(sc_vals):,}"],
                    [f"Mean {sc_lbl}", f"{mean_sc:.2f}"],
                    [f"Median {sc_lbl}", f"{med:.2f}"],
                    ["25th – 75th Percentile (Q1 – Q3)", f"{q1:.2f} – {q3:.2f}"],
                    ["Min – Max", f"{sc_vals[0]:.2f} – {sc_vals[-1]:.2f}"],
                ]
                analyses.append(
                    {
                        "title": f"Grade / Score Summary ('{sc_lbl}')",
                        "status": "ready",
                        "what_calculated": f"Mean, median, quartiles, and range of '{sc_lbl}'" + (f" across '{col_lbl('group')}'." if grp_c else "."),
                        "columns_used": [sc_lbl] + ([col_lbl("group")] if grp_c else []),
                        "how_calculated": f"Aggregated {len(sc_vals):,} valid numeric scores in '{sc_lbl}'.",
                        "what_it_means": f"Typical (median) score is {med:.2f} (mean {mean_sc:.2f}), with the middle 50% of records scoring between {q1:.2f} and {q3:.2f}.",
                        "assumptions_limitations": "Does not assume a fixed passing cutoff (such as 60% or 4.0 scale) because grading scales vary by institution.",
                        "table_headers": ["Statistic", "Value"],
                        "table_rows": table_rows,
                    }
                )
        else:
            skipped_analyses.append(
                {
                    "title": "Grade / Score Distribution",
                    "missing_fields": ["Grade / Exam Score column"],
                    "explanation": "Skipped because no numeric Grade, Score, or GPA column was found in this dataset.",
                }
            )

        if att_c:
            att_lbl = col_lbl("attendance")
            att_vals = sorted(v for v in (to_number(r.get(att_c)) for r in rows) if v is not None)
            if att_vals:
                mean_att = statistics.fmean(att_vals)
                med_att = statistics.median(att_vals)
                analyses.append(
                    {
                        "title": f"Attendance Summary ('{att_lbl}')",
                        "status": "ready",
                        "what_calculated": f"Mean, median, and minimum/maximum attendance across {len(att_vals):,} records.",
                        "columns_used": [att_lbl],
                        "how_calculated": f"Computed arithmetic mean and median over non-empty numeric values in '{att_lbl}'.",
                        "what_it_means": f"Mean attendance is {mean_att:.2f} (median {med_att:.2f}, range {att_vals[0]:.2f} to {att_vals[-1]:.2f}).",
                        "assumptions_limitations": f"Reports '{att_lbl}' in its recorded unit (percentage rate or days present).",
                        "table_headers": ["Attendance Metric", "Value"],
                        "table_rows": [
                            ["Records with Attendance", f"{len(att_vals):,}"],
                            ["Mean Attendance", f"{mean_att:.2f}"],
                            ["Median Attendance", f"{med_att:.2f}"],
                            ["Lowest – Highest", f"{att_vals[0]:.2f} – {att_vals[-1]:.2f}"],
                        ],
                    }
                )
        else:
            skipped_analyses.append(
                {
                    "title": "Attendance Summary",
                    "missing_fields": ["Attendance Rate / Days Present column"],
                    "explanation": "Skipped because no Attendance column exists in this dataset.",
                }
            )

        if grp_c:
            grp_lbl = col_lbl("group")
            items = group_sum_and_count(rows, grp_c, enr_c)
            analyses.append(
                {
                    "title": f"Enrollment / Record Breakdown by {grp_lbl}",
                    "status": "ready",
                    "what_calculated": f"{'Sum of ' + col_lbl('enrollment') if enr_c else 'Record count'} grouped by '{grp_lbl}'.",
                    "columns_used": [grp_lbl] + ([col_lbl("enrollment")] if enr_c else []),
                    "how_calculated": f"Grouped rows by '{grp_lbl}' and {'summed ' + col_lbl('enrollment') if enr_c else 'counted records'}.",
                    "what_it_means": f"Largest group is '{items[0]['group']}' with {items[0]['total']:,.0f} ({items[0]['share_pct']}% share)." if items else "No groups found.",
                    "assumptions_limitations": "If each row represents one student enrollment, record count equals enrolled headcount.",
                    "table_headers": [grp_lbl, "Rows", "Total Enrollment / Count", "Share %"],
                    "table_rows": [[x["group"], x["count"], f"{x['total']:,.0f}", f"{x['share_pct']}%"] for x in items],
                }
            )

    elif cat_key == "marketing":
        camp_c, spend_c, conv_c, rev_c = col_id("campaign"), col_id("spend"), col_id("conversions"), col_id("revenue")
        if spend_c and (conv_c or rev_c):
            sp_lbl = col_lbl("spend")
            out_c = conv_c or rev_c
            out_lbl = col_lbl("conversions") if conv_c else col_lbl("revenue")
            tot_spend = sum(to_number(r.get(spend_c)) or 0.0 for r in rows)
            tot_out = sum(to_number(r.get(out_c)) or 0.0 for r in rows)
            cpa = (tot_spend / tot_out) if tot_out > 0 else None
            roas = (tot_out / tot_spend) if (rev_c and tot_spend > 0) else None

            rows_tbl = [
                [f"Total {sp_lbl}", f"{tot_spend:,.2f}"],
                [f"Total {out_lbl}", f"{tot_out:,.2f}"],
            ]
            if conv_c and cpa is not None:
                rows_tbl.append([f"Cost per {out_lbl} (Spend / Outcome)", f"{cpa:,.2f}"])
            if rev_c and roas is not None:
                rows_tbl.append(["Return on Ad Spend (Revenue / Spend)", f"{roas:.2f}x"])

            analyses.append(
                {
                    "title": "Campaign Spend vs. Outcome Efficiency",
                    "status": "ready",
                    "what_calculated": f"Total '{sp_lbl}', total '{out_lbl}', and efficiency ratio ({'CPA' if conv_c else 'ROAS'}).",
                    "columns_used": [sp_lbl, out_lbl] + ([col_lbl("campaign")] if camp_c else []),
                    "how_calculated": f"Summed '{sp_lbl}' ({tot_spend:,.2f}) and '{out_lbl}' ({tot_out:,.2f}) across all valid rows.",
                    "what_it_means": (
                        f"Total spend of {tot_spend:,.2f} generated {tot_out:,.2f} in '{out_lbl}'"
                        + (f" ({cpa:,.2f} spend per unit outcome)." if conv_c and cpa is not None else "")
                    ),
                    "assumptions_limitations": "Aggregate ratio weights each dollar of spend equally across campaigns; inspect per-campaign rows to spot underperforming outliers.",
                    "table_headers": ["Marketing Metric", "Value"],
                    "table_rows": rows_tbl,
                }
            )
        else:
            skipped_analyses.append(
                {
                    "title": "Campaign Efficiency (CPA / ROAS)",
                    "missing_fields": [k for k, v in [("Ad Spend / Cost column", spend_c), ("Conversions or Revenue Outcome column", conv_c or rev_c)] if not v],
                    "explanation": "Skipped because evaluating campaign efficiency requires both a Spend/Cost column and an Outcome column (Conversions, Leads, or Revenue).",
                }
            )

        if camp_c and (spend_c or conv_c):
            m_c = conv_c or spend_c
            m_lbl = col_lbl("conversions") if conv_c else col_lbl("spend")
            c_lbl = col_lbl("campaign")
            items = group_sum_and_count(rows, camp_c, m_c)
            analyses.append(
                {
                    "title": f"{m_lbl} Breakdown by {c_lbl}",
                    "status": "ready",
                    "what_calculated": f"Total '{m_lbl}' and share of total across '{c_lbl}'.",
                    "columns_used": [c_lbl, m_lbl],
                    "how_calculated": f"Grouped records by '{c_lbl}' and summed '{m_lbl}'.",
                    "what_it_means": f"Top campaign/channel is '{items[0]['group']}' with {items[0]['total']:,.2f} ({items[0]['share_pct']}%)." if items else "No campaign groups found.",
                    "assumptions_limitations": "Based on attributed records present in the uploaded file.",
                    "table_headers": [c_lbl, "Rows", f"Total {m_lbl}", f"Avg {m_lbl}", "Share %"],
                    "table_rows": [[x["group"], x["count"], f"{x['total']:,.2f}", f"{x['mean']:,.2f}", f"{x['share_pct']}%"] for x in items],
                }
            )

    elif cat_key == "customers":
        cust_c, seg_c, val_c = col_id("customer"), col_id("segment"), col_id("value")
        if cust_c:
            cust_lbl = col_lbl("customer")
            counts_by_cust: Counter[str] = Counter()
            spend_by_cust: dict[str, float] = defaultdict(float)
            for r in rows:
                cid = r.get(cust_c)
                if is_missing(cid):
                    continue
                ckey = str(cid).strip()
                counts_by_cust[ckey] += 1
                if val_c:
                    v = to_number(r.get(val_c))
                    if v is not None:
                        spend_by_cust[ckey] += v
            unique_custs = len(counts_by_cust)
            repeat_custs = sum(1 for c in counts_by_cust.values() if c > 1)
            repeat_pct = round((repeat_custs / max(1, unique_custs)) * 100, 1)
            tbl = [
                ["Unique Customers", f"{unique_custs:,}"],
                ["Single-Record Customers (1 row)", f"{unique_custs - repeat_custs:,}"],
                ["Repeat Customers (2+ rows)", f"{repeat_custs:,} ({repeat_pct}%)"],
            ]
            if val_c and spend_by_cust:
                tot_v = sum(spend_by_cust.values())
                tbl.append([f"Average {col_lbl('value')} per Unique Customer", f"{tot_v / max(1, unique_custs):,.2f}"])
            analyses.append(
                {
                    "title": f"Customer Count & Repeat Activity ('{cust_lbl}')",
                    "status": "ready",
                    "what_calculated": f"Distinct customer count, repeat vs. single-record customer split, and average value per customer.",
                    "columns_used": [cust_lbl] + ([col_lbl("value")] if val_c else []),
                    "how_calculated": f"Counted distinct non-empty values in '{cust_lbl}' and tallied how many appear in multiple rows.",
                    "what_it_means": f"Found {unique_custs:,} unique customers across {len(rows):,} rows; {repeat_custs:,} ({repeat_pct}%) appear more than once.",
                    "assumptions_limitations": f"Assumes '{cust_lbl}' uniquely identifies each customer (customers using multiple emails or guest checkouts will appear as separate IDs).",
                    "table_headers": ["Customer Metric", "Value"],
                    "table_rows": tbl,
                }
            )
        else:
            skipped_analyses.append(
                {
                    "title": "Unique Customer & Repeat Count",
                    "missing_fields": ["Customer ID / Name column"],
                    "explanation": "Skipped because no Customer Identifier or Name column was mapped.",
                }
            )

        if seg_c:
            seg_lbl = col_lbl("segment")
            val_lbl = col_lbl("value") if val_c else "Records"
            items = group_sum_and_count(rows, seg_c, val_c)
            analyses.append(
                {
                    "title": f"Customer Segment Comparison by {seg_lbl}",
                    "status": "ready",
                    "what_calculated": f"Record count and {'total/average ' + val_lbl if val_c else 'share of records'} across '{seg_lbl}'.",
                    "columns_used": [seg_lbl] + ([val_lbl] if val_c else []),
                    "how_calculated": f"Grouped rows by '{seg_lbl}' and {'summed ' + val_lbl if val_c else 'counted records'}.",
                    "what_it_means": f"Largest segment is '{items[0]['group']}' ({items[0]['share_pct']}% share, {items[0]['count']} rows)." if items else "No segments found.",
                    "assumptions_limitations": "Segment totals reflect all rows belonging to that segment.",
                    "table_headers": [seg_lbl, "Rows", f"Total {val_lbl}" if val_c else "Count", f"Avg {val_lbl}" if val_c else "Avg", "Share %"],
                    "table_rows": [[x["group"], x["count"], f"{x['total']:,.2f}" if val_c else str(x["count"]), f"{x['mean']:,.2f}" if val_c else "-", f"{x['share_pct']}%"] for x in items],
                }
            )
        else:
            skipped_analyses.append(
                {
                    "title": "Customer Segment Comparison",
                    "missing_fields": ["Customer Segment / Tier / Region column"],
                    "explanation": "Skipped because no categorical Segment, Tier, Plan, or Region column was found.",
                }
            )

    # Fallback / General grouped analysis for other categories or when specific fields are absent
    if not analyses:
        cat_col_fallback = next((c for c in cols if c["kind"] == "category" and not c.get("is_identifier")), None)
        num_col_fallback = next((c for c in cols if c["kind"] == "number" and not c.get("is_identifier")), None)
        if cat_col_fallback and num_col_fallback:
            g_name, g_lbl = cat_col_fallback["name"], cat_col_fallback.get("label", cat_col_fallback["name"])
            m_name, m_lbl = num_col_fallback["name"], num_col_fallback.get("label", num_col_fallback["name"])
            items = group_sum_and_count(rows, g_name, m_name)
            analyses.append(
                {
                    "title": f"General Grouped Breakdown: {m_lbl} by {g_lbl}",
                    "status": "fallback",
                    "what_calculated": f"Sum, average, and record count of '{m_lbl}' grouped by '{g_lbl}'.",
                    "columns_used": [g_lbl, m_lbl],
                    "how_calculated": f"Because domain-specific columns for {cat_spec['label']} were absent or unmapped, DataPrism grouped the primary category '{g_lbl}' by the numeric measure '{m_lbl}'.",
                    "what_it_means": f"'{items[0]['group']}' has the highest total '{m_lbl}' at {items[0]['total']:,.2f} ({items[0]['share_pct']}% of total)." if items else "No valid grouped records.",
                    "assumptions_limitations": "You can map specific columns above if your dataset uses custom column names.",
                    "table_headers": [g_lbl, "Records", f"Total {m_lbl}", f"Average {m_lbl}", "Share %"],
                    "table_rows": [[x["group"], x["count"], f"{x['total']:,.2f}", f"{x['mean']:,.2f}", f"{x['share_pct']}%"] for x in items],
                }
            )

    return {
        "category": cat_key,
        "category_label": cat_spec["label"],
        "category_icon": cat_spec["icon"],
        "category_tagline": cat_spec["tagline"],
        "roles": roles,
        "analyses": analyses,
        "skipped_analyses": skipped_analyses,
        "explanation": {
            "what_calculated": f"Domain-tailored checks for {cat_spec['label']} based on verified column mappings.",
            "columns_used": [r["column_label"] for r in roles.values() if r.get("column_label")],
            "how_calculated": "Columns are matched conservatively by name and detected data type, or mapped explicitly by your selection above.",
            "what_it_means": (
                f"{len(analyses)} domain analysis block(s) computed and {len(skipped_analyses)} skipped due to absent columns."
            ),
            "assumptions_limitations": (
                "DataPrism never invents domain metrics (like Profit, ROI, or Grades) from unrelated columns. "
                "When required fields are absent, those specific cards are skipped with a clear explanation of what is missing."
            ),
        },
    }


def build_chart_data(
    rows: list[dict[str, Any]],
    cols: list[dict[str, Any]],
    chart_type: str = "bar",
    x_col: str | None = None,
    y_col: str | None = None,
    agg: str = "sum",
    group_col: str | None = None,
    date_grain: str = "month",
) -> dict[str, Any]:
    """Validated Chart Builder supporting Bar, Line, Scatter, Donut, and Histogram with optional grouping."""
    col_map = {c["name"]: c for c in cols}
    cat_cols = [c for c in cols if c["kind"] == "category" and not c.get("is_identifier")]
    date_cols = [c for c in cols if c["kind"] == "date"]
    num_cols = [c for c in cols if c["kind"] == "number" and not c.get("is_identifier")]

    ctype = chart_type if chart_type in {"bar", "line", "scatter", "donut", "histogram"} else "bar"

    # Choose smart defaults when x_col or y_col is not provided
    if not x_col or x_col not in col_map:
        if ctype == "line" and date_cols:
            x_col = date_cols[0]["name"]
        elif ctype in {"scatter", "histogram"} and num_cols:
            x_col = num_cols[0]["name"]
        elif cat_cols:
            x_col = cat_cols[0]["name"]
        elif date_cols:
            x_col = date_cols[0]["name"]
        elif cols:
            x_col = cols[0]["name"]

    if not y_col or (y_col != "__count__" and y_col not in col_map):
        if ctype == "histogram":
            y_col = "__count__"
        elif ctype == "scatter" and len(num_cols) >= 2:
            y_col = next((c["name"] for c in num_cols if c["name"] != x_col), num_cols[0]["name"])
        elif num_cols:
            y_col = num_cols[0]["name"]
        else:
            y_col = "__count__"

    if not x_col or x_col not in col_map:
        return {"valid": False, "error": "No columns available to build a chart.", "points": [], "series": []}

    x_meta = col_map[x_col]
    x_label = x_meta.get("label", x_col)
    y_meta = col_map.get(y_col) if y_col != "__count__" else None
    y_label = "Row Count" if y_col == "__count__" else (y_meta.get("label", y_col) if y_meta else str(y_col))
    grp_meta = col_map.get(group_col) if (group_col and group_col in col_map and group_col != x_col) else None
    grp_label = grp_meta.get("label", group_col) if grp_meta else None

    warnings: list[str] = []

    # Validate compatibility between chart type, X column, Y column, and aggregation
    if ctype == "scatter":
        if x_meta["kind"] != "number":
            return {
                "valid": False,
                "error": f"Scatter plots require two numeric columns, but X-axis '{x_label}' is a {x_meta['kind']} column. Choose a numeric X column or switch Chart Type to Bar.",
                "points": [],
                "series": [],
            }
        if not y_meta or y_meta["kind"] != "number":
            return {
                "valid": False,
                "error": f"Scatter plots require a numeric Y-axis column, but '{y_label}' is not a numeric measure. Select a numeric column for Y-axis.",
                "points": [],
                "series": [],
            }
        if x_meta.get("is_identifier") or y_meta.get("is_identifier"):
            warnings.append("One of the selected columns was flagged as an identifier (ID); plotting IDs rarely shows meaningful analytical patterns.")
        agg = "raw"

    elif ctype == "histogram":
        if x_meta["kind"] != "number":
            return {
                "valid": False,
                "error": f"Histograms show the frequency distribution of a numeric column, but '{x_label}' is a {x_meta['kind']} column. Choose a numeric X column or use a Bar chart for categories.",
                "points": [],
                "series": [],
            }
        y_col = "__count__"
        y_label = "Row Count (Frequency)"
        agg = "count"

    else:
        # Bar, Line, Donut
        if y_col == "__count__":
            agg = "count"
        else:
            if y_meta and y_meta["kind"] != "number" and agg != "count":
                return {
                    "valid": False,
                    "error": f"Cannot calculate {agg.upper()} on '{y_label}' because it is a {y_meta['kind']} column. Choose a numeric Y-axis column or set Y-axis to 'Row Count'.",
                    "points": [],
                    "series": [],
                }
            if y_meta and y_meta.get("is_identifier") and agg in {"sum", "mean", "median"}:
                return {
                    "valid": False,
                    "error": f"'{y_label}' is flagged as an Identifier (ID) column. Summing or averaging ID numbers is not meaningful—use 'Row Count' or pick a quantitative measure.",
                    "points": [],
                    "series": [],
                }
        if ctype == "line" and x_meta["kind"] not in {"date", "number"}:
            warnings.append(
                f"X-axis '{x_label}' is an unordered category column. Line charts imply a continuous sequence (like dates); a Bar chart is usually clearer for comparing categories."
            )
        if x_meta.get("is_identifier") and ctype in {"bar", "donut"}:
            warnings.append(
                f"X-axis '{x_label}' is an identifier column where nearly every row is unique. Grouping by an ID simply lists individual records."
            )

    # Execute chart calculation
    rows_used = 0
    rows_skipped = 0

    if ctype == "scatter":
        pts: list[dict[str, Any]] = []
        for idx, r in enumerate(rows, start=1):
            vx = to_number(r.get(x_col))
            vy = to_number(r.get(y_col)) if y_col else None
            if vx is None or vy is None:
                rows_skipped += 1
                continue
            rows_used += 1
            g_val = str(r.get(grp_meta["name"]) or "Unassigned").strip() if grp_meta else "All Records"
            if len(pts) < 400:
                pts.append({"x": round(vx, 4), "y": round(vy, 4), "group": g_val, "row": idx})
        title = f"{y_label} vs. {x_label}" + (f" (colored by {grp_label})" if grp_label else "")
        how_text = f"Plots raw unaggregated rows (each dot is 1 row) with X = '{x_label}' and Y = '{y_label}' across {rows_used:,} complete rows."
        return {
            "valid": True,
            "type": "scatter",
            "x": x_col,
            "x_label": x_label,
            "y": y_col,
            "y_label": y_label,
            "agg": "raw",
            "group": grp_meta["name"] if grp_meta else None,
            "group_label": grp_label,
            "title": title,
            "how_calculated": how_text,
            "rows_used": rows_used,
            "rows_skipped": rows_skipped,
            "warnings": warnings,
            "points": pts,
            "series": [],
        }

    if ctype == "histogram":
        vals: list[float] = []
        for r in rows:
            v = to_number(r.get(x_col))
            if v is None:
                rows_skipped += 1
            else:
                rows_used += 1
                vals.append(v)
        if not vals:
            return {"valid": False, "error": f"No valid numeric values in '{x_label}'.", "points": [], "series": []}
        min_v, max_v = min(vals), max(vals)
        bin_count = min(12, max(5, int(math.sqrt(len(vals)))))
        if math.isclose(min_v, max_v):
            pts = [{"x": f"{min_v:,.2f}", "y": len(vals), "row_count": len(vals)}]
        else:
            step = (max_v - min_v) / bin_count
            bins = [0] * bin_count
            for v in vals:
                idx = min(bin_count - 1, int((v - min_v) / step))
                bins[idx] += 1
            pts = []
            for i, count in enumerate(bins):
                b_low = min_v + i * step
                b_high = min_v + (i + 1) * step
                pts.append({"x": f"{b_low:,.1f}–{b_high:,.1f}", "y": count, "row_count": count})
        return {
            "valid": True,
            "type": "histogram",
            "x": x_col,
            "x_label": f"{x_label} (Binned Ranges)",
            "y": "__count__",
            "y_label": "Number of Rows",
            "agg": "count",
            "group": None,
            "group_label": None,
            "title": f"Distribution of {x_label} ({len(pts)} equal-width bins)",
            "how_calculated": f"Divides '{x_label}' range ({min_v:,.2f} to {max_v:,.2f}) into {len(pts)} equal-width numeric intervals and counts how many rows fall into each bin ({rows_used:,} rows used).",
            "rows_used": rows_used,
            "rows_skipped": rows_skipped,
            "warnings": warnings,
            "points": pts,
            "series": [],
        }

    # Grouped / Aggregated Bar, Line, Donut
    agg = agg if agg in {"sum", "mean", "median", "count", "min", "max"} else "sum"
    buckets: dict[str, dict[str, list[float]]] = defaultdict(lambda: defaultdict(list))
    group_totals: Counter[str] = Counter()

    for r in rows:
        raw_x = r.get(x_col)
        if is_missing(raw_x):
            rows_skipped += 1
            continue
        if x_meta["kind"] == "date":
            iso_d, _ = parse_date_value(raw_x)
            if not iso_d:
                rows_skipped += 1
                continue
            x_key = bucket_date(iso_d, date_grain or "month")
        else:
            x_key = str(raw_x).strip()

        if y_col == "__count__" or agg == "count":
            y_val = 1.0
        else:
            y_val = to_number(r.get(y_col)) if y_col else None
            if y_val is None:
                rows_skipped += 1
                continue

        g_key = str(r.get(grp_meta["name"]) or "Missing").strip() if grp_meta else "All"
        buckets[x_key][g_key].append(y_val)
        group_totals[g_key] += 1
        rows_used += 1

    if not buckets:
        return {
            "valid": False,
            "error": f"No valid rows had non-empty values for '{x_label}' and '{y_label}'.",
            "points": [],
            "series": [],
        }

    # Sort X keys chronologically for dates/numbers, or by descending total for categories
    x_Totals: list[tuple[str, float, int]] = []
    for xk, gmap in buckets.items():
        all_vals = [v for glist in gmap.values() for v in glist]
        x_Totals.append((xk, aggregate_values(all_vals, agg), len(all_vals)))

    if x_meta["kind"] == "date" or ctype == "line":
        x_Totals.sort(key=lambda item: item[0])
    else:
        x_Totals.sort(key=lambda item: item[1], reverse=True)

    max_bars = 10 if ctype == "donut" else (36 if ctype == "line" else 15)
    if len(x_Totals) > max_bars and ctype in {"bar", "donut"}:
        shown = x_Totals[: max_bars - 1]
        tail = x_Totals[max_bars - 1 :]
        if agg in {"sum", "count"}:
            tail_val = sum(t[1] for t in tail)
            tail_rows = sum(t[2] for t in tail)
            shown.append((f"Other ({len(tail)} more)", tail_val, tail_rows))
            warnings.append(f"Grouped the smallest {len(tail)} categories into 'Other' so the {ctype} chart remains readable.")
        else:
            warnings.append(f"Showing top {len(shown)} of {len(x_Totals)} categories ranked by {agg.upper()} {y_label}.")
        x_Totals = shown

    if ctype == "donut" and any(t[1] < 0 for t in x_Totals):
        return {
            "valid": False,
            "error": "Donut charts cannot display negative values because slice angles must be positive proportions of a whole. Switch Chart Type to Bar.",
            "points": [],
            "series": [],
        }

    top_groups = [g for g, _ in group_totals.most_common(5)] if grp_meta else ["All"]
    points = []
    for xk, total_val, r_cnt in x_Totals:
        pt: dict[str, Any] = {"x": xk, "y": round(total_val, 4), "row_count": r_cnt}
        if grp_meta and xk in buckets:
            pt["groups"] = {
                g: round(aggregate_values(buckets[xk].get(g, []), agg), 4) for g in top_groups if buckets[xk].get(g)
            }
        points.append(pt)

    agg_verb = {
        "sum": f"Sum of {y_label}",
        "mean": f"Average {y_label}",
        "median": f"Median {y_label}",
        "count": "Row Count",
        "min": f"Minimum {y_label}",
        "max": f"Maximum {y_label}",
    }.get(agg, y_label)

    grain_suffix = f" (by {date_grain})" if x_meta["kind"] == "date" else ""
    title = f"{agg_verb} by {x_label}{grain_suffix}" + (f" • Grouped by {grp_label}" if grp_label else "")
    how_text = (
        f"Groups {rows_used:,} valid rows by '{x_label}'{grain_suffix}"
        + (f" and '{grp_label}'" if grp_label else "")
        + f", then calculates {agg.upper()} of '{y_label}' for each group."
    )

    return {
        "valid": True,
        "type": ctype,
        "x": x_col,
        "x_label": x_label + grain_suffix,
        "y": y_col,
        "y_label": agg_verb,
        "agg": agg,
        "date_grain": date_grain if x_meta["kind"] == "date" else None,
        "group": grp_meta["name"] if grp_meta else None,
        "group_label": grp_label,
        "groups_list": top_groups if grp_meta else [],
        "title": title,
        "how_calculated": how_text,
        "rows_used": rows_used,
        "rows_skipped": rows_skipped,
        "warnings": warnings,
        "points": points,
    }


def make_sample(category: str) -> tuple[str, list[dict[str, Any]], list[dict[str, Any]], dict[str, Any]]:
    """Create a realistic, transparent starter dataset for the selected category."""
    cat = category if category in CATEGORIES else "ecommerce"
    if cat == "retail":
        csv_text = (
            "Store_Location,Product_Name,Department,Stock_On_Hand,Reorder_Point,Units_Sold_30d,Unit_Price,Count_Date\n"
            "Downtown Flagship,Organic Oat Milk 1L,Grocery,142,50,310,4.50,2025-01-15\n"
            "Downtown Flagship,Artisan Sourdough,Bakery,18,25,195,6.80,2025-01-15\n"
            "Westside Market,Organic Oat Milk 1L,Grocery,34,40,220,4.50,2025-01-15\n"
            "Westside Market,Cold Brew Bottle,Beverages,95,30,175,3.95,2025-01-15\n"
            "North Hills Store,Ceramic Pour-Over,Homeware,12,15,28,24.00,2025-02-01\n"
            "North Hills Store,Whole Bean Espresso,Grocery,64,35,140,14.50,2025-02-01\n"
            "Harbor View Store,Artisan Sourdough,Bakery,8,20,160,6.80,2025-02-01\n"
            "Harbor View Store,Cold Brew Bottle,Beverages,110,30,190,3.95,2025-02-15\n"
            "Downtown Flagship,Whole Bean Espresso,Grocery,88,35,205,14.50,2025-02-15\n"
            "Westside Market,Ceramic Pour-Over,Homeware,6,10,19,24.00,2025-03-01\n"
            "North Hills Store,Organic Oat Milk 1L,Grocery,120,45,260,4.50,2025-03-01\n"
            "Harbor View Store,Whole Bean Espresso,Grocery,29,35,115,14.50,2025-03-15\n"
        )
        name = "Retail Inventory & Store Sample"
    elif cat == "education":
        csv_text = (
            "Student_ID,Program,Term_Date,Attendance_Pct,Exam_Score,Credits_Enrolled\n"
            "STU-1001,Data Science,2025-01-20,96.5,91.0,15\n"
            "STU-1002,Data Science,2025-01-20,88.0,84.5,15\n"
            "STU-1003,Computer Science,2025-01-20,92.0,88.0,16\n"
            "STU-1004,Business Analytics,2025-01-20,79.5,73.0,12\n"
            "STU-1005,Computer Science,2025-02-15,98.0,95.5,16\n"
            "STU-1006,Business Analytics,2025-02-15,85.0,78.0,15\n"
            "STU-1007,Data Science,2025-02-15,68.0,64.0,12\n"
            "STU-1008,Economics,2025-03-10,91.0,86.0,15\n"
            "STU-1009,Economics,2025-03-10,83.5,79.5,15\n"
            "STU-1010,Computer Science,2025-03-10,94.0,90.0,16\n"
            "STU-1011,Business Analytics,2025-04-05,90.5,85.0,15\n"
            "STU-1012,Data Science,2025-04-05,95.0,92.5,15\n"
        )
        name = "Academic Course & Assessment Sample"
    elif cat == "marketing":
        csv_text = (
            "Campaign_ID,Channel,Campaign_Name,Report_Date,Ad_Spend,Impressions,Clicks,Conversions,Attributed_Revenue\n"
            "CMP-01,Paid Search,Spring Brand Search,2025-01-10,1250.00,48000,1920,96,7680.00\n"
            "CMP-02,Social Ads,Creator Video Reel,2025-01-24,2100.00,115000,2300,64,4480.00\n"
            "CMP-03,Email,Newsletter Reactivation,2025-02-07,320.00,22000,1540,85,5950.00\n"
            "CMP-04,Paid Search,Competitor Keywords,2025-02-21,1850.00,39000,1170,41,3280.00\n"
            "CMP-05,Display,Retargeting Banner,2025-03-07,940.00,68000,1020,52,3900.00\n"
            "CMP-06,Social Ads,Spring Lookbook Carousel,2025-03-21,1680.00,92000,1840,73,5475.00\n"
            "CMP-07,Affiliate,Partner Review Blog,2025-04-04,780.00,19500,980,49,3920.00\n"
            "CMP-08,Paid Search,Spring Brand Search,2025-04-18,1420.00,52000,2180,112,8960.00\n"
        )
        name = "Marketing Campaign Performance Sample"
    elif cat == "customers":
        csv_text = (
            "Customer_ID,Customer_Segment,Region,Activity_Date,Orders_Count,Total_Spent\n"
            "CUST-501,Enterprise,North America,2025-01-12,5,4250.00\n"
            "CUST-502,Mid-Market,Europe,2025-01-19,3,1890.00\n"
            "CUST-503,Self-Serve,Asia-Pacific,2025-01-28,1,240.00\n"
            "CUST-501,Enterprise,North America,2025-02-14,4,3600.00\n"
            "CUST-504,Mid-Market,North America,2025-02-22,2,1320.00\n"
            "CUST-505,Self-Serve,Europe,2025-03-03,1,195.00\n"
            "CUST-502,Mid-Market,Europe,2025-03-18,2,1450.00\n"
            "CUST-506,Enterprise,Asia-Pacific,2025-03-29,6,5400.00\n"
            "CUST-507,Self-Serve,North America,2025-04-08,1,310.00\n"
            "CUST-506,Enterprise,Asia-Pacific,2025-04-19,3,2950.00\n"
        )
        name = "Customer Segment & Spend Sample"
    else:
        csv_text = (
            "Order_ID,Order_Date,Product_Name,Category,Sales_Channel,Customer_ID,Quantity,Revenue,Cost\n"
            "ORD-1001,2025-01-08,Ergonomic Desk Chair,Furniture,Direct Web,CUST-01,2,580.00,310.00\n"
            "ORD-1002,2025-01-19,Mechanical Keyboard,Accessories,Marketplace,CUST-02,3,360.00,180.00\n"
            "ORD-1003,2025-01-27,4K Studio Monitor,Electronics,Direct Web,CUST-03,1,650.00,410.00\n"
            "ORD-1004,2025-02-05,USB-C Docking Station,Accessories,B2B Wholesale,CUST-01,5,750.00,400.00\n"
            "ORD-1005,2025-02-14,Standing Desk Frame,Furniture,Direct Web,CUST-04,1,490.00,265.00\n"
            "ORD-1006,2025-02-23,Mechanical Keyboard,Accessories,Direct Web,CUST-05,2,240.00,120.00\n"
            "ORD-1007,2025-03-06,4K Studio Monitor,Electronics,B2B Wholesale,CUST-06,4,2480.00,1640.00\n"
            "ORD-1008,2025-03-17,Ergonomic Desk Chair,Furniture,Marketplace,CUST-02,1,290.00,155.00\n"
            "ORD-1009,2025-03-28,Acoustic Desk Panel,Furniture,Direct Web,CUST-07,3,420.00,195.00\n"
            "ORD-1010,2025-04-09,Standing Desk Frame,Furniture,B2B Wholesale,CUST-01,3,1470.00,795.00\n"
            "ORD-1011,2025-04-16,USB-C Docking Station,Accessories,Marketplace,CUST-08,2,300.00,160.00\n"
            "ORD-1012,2025-04-24,4K Studio Monitor,Electronics,Direct Web,CUST-04,2,1300.00,820.00\n"
        )
        name = "Sales & E-commerce Orders Sample"

    cols, records, diag = decode_upload(f"{cat}_sample.csv", csv_text.encode("utf-8"))
    return name, cols, records, diag


def transform_records(
    cols: list[dict[str, Any]],
    rows: list[dict[str, Any]],
    action: str,
    target_col: str | None = None,
) -> tuple[list[dict[str, Any]], str]:
    """Execute a data cleaning transformation and return (new_rows, human_readable_note)."""
    before_count = len(rows)
    if action == "deduplicate":
        seen: set[str] = set()
        deduped: list[dict[str, Any]] = []
        for r in rows:
            k = json.dumps(r, sort_keys=True, default=str)
            if k not in seen:
                seen.add(k)
                deduped.append(dict(r))
        removed = before_count - len(deduped)
        return deduped, f"Removed {removed} exact duplicate row(s) ({len(deduped)} rows kept)."

    if action == "drop_missing":
        if target_col:
            kept = [dict(r) for r in rows if not is_missing(r.get(target_col))]
            removed = before_count - len(kept)
            return kept, f"Dropped {removed} row(s) with missing values in '{target_col}' ({len(kept)} rows kept)."
        kept = [dict(r) for r in rows if all(not is_missing(r.get(c["name"])) for c in cols)]
        removed = before_count - len(kept)
        return kept, f"Dropped {removed} row(s) containing any missing value ({len(kept)} complete rows kept)."

    if action == "fill_missing_median" and target_col:
        vals = sorted(v for v in (to_number(r.get(target_col)) for r in rows) if v is not None)
        if not vals:
            raise ValueError(f"Column '{target_col}' has no valid numeric values to compute a median.")
        med = round(float(statistics.median(vals)), 4)
        filled = 0
        out = []
        for r in rows:
            nr = dict(r)
            if to_number(nr.get(target_col)) is None:
                nr[target_col] = med
                filled += 1
            out.append(nr)
        return out, f"Filled {filled} missing value(s) in '{target_col}' with column median ({med})."

    if action == "fill_missing_zero" and target_col:
        filled = 0
        out = []
        for r in rows:
            nr = dict(r)
            if to_number(nr.get(target_col)) is None:
                nr[target_col] = 0.0
                filled += 1
            out.append(nr)
        return out, f"Filled {filled} missing value(s) in '{target_col}' with 0."

    if action == "trim_text":
        changed = 0
        out = []
        for r in rows:
            nr = dict(r)
            for k, v in nr.items():
                if isinstance(v, str):
                    trimmed = re.sub(r"\s+", " ", v.strip())
                    if trimmed != v:
                        nr[k] = trimmed
                        changed += 1
            out.append(nr)
        return out, f"Trimmed leading/trailing/repeated whitespace across {changed} text cell(s)."

    raise ValueError(f"Unsupported cleaning action '{action}'.")


def build_dashboard(
    dataset: dict[str, Any],
    all_rows: list[dict[str, Any]],
    filtered_rows: list[dict[str, Any]],
    active_filter: dict[str, str] | None = None,
    trend_params: dict[str, str] | None = None,
    rel_params: dict[str, str] | None = None,
) -> dict[str, Any]:
    cols = dataset["columns"]
    category = dataset.get("category") or "ecommerce"
    diagnostics = dataset.get("diagnostics") or {}
    custom_units = dataset.get("custom_units") or {}
    role_overrides = dataset.get("role_overrides") or {}
    trend_params = trend_params or {}
    rel_params = rel_params or {}

    health = build_dataset_health(filtered_rows, cols, diagnostics)
    numeric_summary = build_numeric_summary(filtered_rows, cols, custom_units)
    category_breakdown = build_category_breakdown(filtered_rows, cols)
    time_trends = build_time_trends(
        filtered_rows,
        cols,
        date_col=trend_params.get("date_col"),
        measure_col=trend_params.get("measure_col"),
        grain=trend_params.get("grain"),
        agg=trend_params.get("agg"),
    )
    relationships = build_relationships(
        filtered_rows,
        cols,
        x_col=rel_params.get("x_col"),
        y_col=rel_params.get("y_col"),
        show_corr=rel_params.get("show_corr", "1") != "0",
    )
    cat_analysis = build_category_analysis(filtered_rows, cols, category, role_overrides)
    default_chart = build_chart_data(filtered_rows, cols)

    # Concise top-level Dataset Facts (replacing generic unexplained KPI cards)
    num_count = sum(1 for c in cols if c["kind"] == "number" and not c.get("is_identifier"))
    cat_count = sum(1 for c in cols if c["kind"] == "category" and not c.get("is_identifier"))
    date_count = sum(1 for c in cols if c["kind"] == "date")
    id_count = sum(1 for c in cols if c.get("is_identifier") or c["kind"] == "id")

    facts = [
        {
            "label": "Rows Analyzed",
            "value": f"{len(filtered_rows):,}",
            "sub": f"of {len(all_rows):,} total stored rows" if len(filtered_rows) != len(all_rows) else "100% of stored dataset",
            "how": "Count of active records passing any applied filter.",
        },
        {
            "label": "Columns & Types",
            "value": f"{len(cols)} columns",
            "sub": f"{num_count} numeric • {cat_count} category • {date_count} date • {id_count} ID",
            "how": "Inferred from >=75% valid value parsing and ID uniqueness checks.",
        },
        {
            "label": "Data Completeness",
            "value": f"{100.0 - health['missing_pct']:.1f}%",
            "sub": f"{health['missing_cells']:,} missing cells • {health['complete_rows']:,} complete rows",
            "how": "(Non-empty cells / Total cells) × 100 across active rows.",
        },
        {
            "label": "Duplicate Rows",
            "value": f"{health['duplicate_rows']:,}",
            "sub": f"{health['duplicate_pct']}% of active rows are exact duplicates",
            "how": "Exact full-row match across all columns.",
        },
    ]

    return {
        "dataset": {
            "id": dataset["id"],
            "name": dataset["name"],
            "category": category,
            "category_label": CATEGORIES.get(category, CATEGORIES["general"])["label"],
            "created_at": dataset["created_at"],
            "row_count": len(all_rows),
            "filtered_row_count": len(filtered_rows),
            "parent_id": dataset.get("parent_id"),
            "transform_note": dataset.get("transform_note") or "",
            "columns": cols,
            "custom_units": custom_units,
            "role_overrides": role_overrides,
        },
        "active_filter": active_filter,
        "facts": facts,
        "health": health,
        "numeric_summary": numeric_summary,
        "category_breakdown": category_breakdown,
        "time_trends": time_trends,
        "relationships": relationships,
        "category_analysis": cat_analysis,
        "default_chart": default_chart,
        # Backwards-compatible aliases
        "profile": health,
        "summary": facts,
    }


def parse_multipart(content_type: str, body: bytes) -> tuple[dict[str, str], tuple[str, bytes] | None]:
    fields: dict[str, str] = {}
    file_tuple: tuple[str, bytes] | None = None
    header = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
    msg = email.parser.BytesParser(policy=email.policy.default).parsebytes(header + body)
    if not msg.is_multipart():
        return fields, file_tuple
    for part in msg.iter_parts():
        disp = part.get("Content-Disposition", "")
        name_m = re.search(r'name="([^"]+)"', disp)
        if not name_m:
            continue
        field_name = name_m.group(1)
        fn_m = re.search(r'filename="([^"]*)"', disp)
        payload = part.get_payload(decode=True) or b""
        if fn_m is not None:
            fn = fn_m.group(1) or "upload.csv"
            file_tuple = (fn, payload)
        else:
            fields[field_name] = payload.decode("utf-8", errors="replace").strip()
    return fields, file_tuple


class PlatformHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: Any) -> None:
        sys.stderr.write(f"[DataPrism] {self.address_string()} - {fmt % args}\n")

    def _send_json(self, payload: dict[str, Any], status: int = 200, cookie_header: str | None = None) -> None:
        raw = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        if cookie_header:
            self.send_header("Set-Cookie", cookie_header)
        self.end_headers()
        self.wfile.write(raw)

    def _current_user(self, con: sqlite3.Connection) -> dict[str, Any] | None:
        raw_cookie = self.headers.get("Cookie", "")
        if not raw_cookie:
            return None
        cookie = SimpleCookie()
        try:
            cookie.load(raw_cookie)
        except Exception:
            return None
        morsel = cookie.get(COOKIE_NAME)
        if not morsel or not morsel.value:
            return None
        token = morsel.value
        now = now_iso()
        row = con.execute(
            """
            SELECT u.id, u.email, u.category, u.created_at
            FROM sessions s
            JOIN users u ON u.id = s.user_id
            WHERE s.token = ? AND s.expires_at > ?
            """,
            (token, now),
        ).fetchone()
        return dict(row) if row else None

    def _make_session_cookie(self, con: sqlite3.Connection, user_id: int) -> str:
        token = secrets.token_urlsafe(32)
        now = datetime.now(timezone.utc)
        expires = now + timedelta(days=SESSION_DAYS)
        con.execute(
            "INSERT INTO sessions (token, user_id, created_at, expires_at) VALUES (?, ?, ?, ?)",
            (token, user_id, now.strftime("%Y-%m-%dT%H:%M:%SZ"), expires.strftime("%Y-%m-%dT%H:%M:%SZ")),
        )
        max_age = SESSION_DAYS * 86400
        return f"{COOKIE_NAME}={token}; HttpOnly; Path=/; SameSite=Lax; Max-Age={max_age}"

    def _read_json_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8", errors="replace") or "{}")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path == "/api/health":
            self._send_json({"ok": True, "status": "healthy", "storage": storage_info()})
            return

        if path == "/api/me":
            with connect() as con:
                user = self._current_user(con)
                categories_list = [
                    {"id": k, "label": v["label"], "icon": v["icon"], "tagline": v["tagline"]}
                    for k, v in CATEGORIES.items()
                    if k in {"ecommerce", "retail", "education", "marketing", "customers", "finance", "product", "general"}
                ]
                self._send_json({"user": user, "categories": categories_list, "storage": storage_info()})
            return

        if path == "/api/dashboard":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required. Please sign in."}, 401)
                    return
                ds_Rows = con.execute(
                    "SELECT id, name, category, created_at, row_count, parent_id, transform_note FROM datasets WHERE user_id = ? ORDER BY id DESC",
                    (user["id"],),
                ).fetchall()
                datasets_list = [
                    {
                        **dict(r),
                        "category_label": CATEGORIES.get(r["category"], CATEGORIES["general"])["label"],
                    }
                    for r in ds_Rows
                ]
                req_ds_id = int(qs.get("dataset", ["0"])[0] or 0) or None
                ds, all_rows = load_dataset_and_rows(con, user["id"], req_ds_id)
                if not ds:
                    self._send_json(
                        {
                            "datasets": datasets_list,
                            "dataset": None,
                            "storage": storage_info(),
                        }
                    )
                    return

                f_field = qs.get("filter_field", [""])[0]
                f_op = qs.get("filter_op", [""])[0]
                f_val = qs.get("filter_val", [""])[0]
                active_filter = {"field": f_field, "op": f_op, "value": f_val} if f_field and f_op else None
                filtered_rows = apply_filter(all_rows, f_field, f_op, f_val) if active_filter else all_rows

                trend_params = {
                    "date_col": qs.get("trend_date", [None])[0],
                    "measure_col": qs.get("trend_measure", [None])[0],
                    "grain": qs.get("trend_grain", [None])[0],
                    "agg": qs.get("trend_agg", [None])[0],
                }
                rel_params = {
                    "x_col": qs.get("rel_x", [None])[0],
                    "y_col": qs.get("rel_y", [None])[0],
                    "show_corr": qs.get("rel_corr", ["1"])[0],
                }

                dash = build_dashboard(ds, all_rows, filtered_rows, active_filter, trend_params, rel_params)
                dash["datasets"] = datasets_list
                dash["storage"] = storage_info()
                self._send_json(dash)
            return

        if path == "/api/preview":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                req_ds_id = int(qs.get("dataset", ["0"])[0] or 0) or None
                ds, all_rows = load_dataset_and_rows(con, user["id"], req_ds_id)
                if not ds:
                    self._send_json({"dataset_id": None, "page": 1, "page_size": 25, "total": 0, "rows": [], "columns": []})
                    return

                f_field = qs.get("filter_field", [""])[0]
                f_op = qs.get("filter_op", [""])[0]
                f_val = qs.get("filter_val", [""])[0]
                rows = apply_filter(all_rows, f_field, f_op, f_val) if (f_field and f_op) else all_rows

                search_q = qs.get("q", [""])[0].strip().lower()
                if search_q:
                    rows = [
                        r
                        for r in rows
                        if any(not is_missing(v) and search_q in str(v).lower() for v in r.values())
                    ]

                page = max(1, int(qs.get("page", ["1"])[0] or 1))
                page_size = max(5, min(100, int(qs.get("page_size", ["25"])[0] or 25)))
                total = len(rows)
                start = (page - 1) * page_size
                end = start + page_size
                self._send_json(
                    {
                        "dataset_id": ds["id"],
                        "columns": ds["columns"],
                        "page": page,
                        "page_size": page_size,
                        "total": total,
                        "total_pages": max(1, math.ceil(total / page_size)),
                        "rows": rows[start:end],
                    }
                )
            return

        if path == "/api/chart":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                req_ds_id = int(qs.get("dataset", ["0"])[0] or 0) or None
                ds, all_rows = load_dataset_and_rows(con, user["id"], req_ds_id)
                if not ds:
                    self._send_json({"valid": False, "error": "No active dataset found."})
                    return
                f_field = qs.get("filter_field", [""])[0]
                f_op = qs.get("filter_op", [""])[0]
                f_val = qs.get("filter_val", [""])[0]
                rows = apply_filter(all_rows, f_field, f_op, f_val) if (f_field and f_op) else all_rows

                chart_res = build_chart_data(
                    rows=rows,
                    cols=ds["columns"],
                    chart_type=qs.get("type", ["bar"])[0],
                    x_col=qs.get("x", [None])[0],
                    y_col=qs.get("y", [None])[0],
                    agg=qs.get("agg", ["sum"])[0],
                    group_col=qs.get("group", [None])[0] or None,
                    date_grain=qs.get("grain", ["month"])[0],
                )
                self._send_json(chart_res)
            return

        if path == "/api/export":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                req_ds_id = int(qs.get("dataset", ["0"])[0] or 0) or None
                ds, all_rows = load_dataset_and_rows(con, user["id"], req_ds_id)
                if not ds:
                    self._send_json({"error": "Dataset not found."}, 404)
                    return
                fmt = qs.get("format", ["csv"])[0].lower()
                safe_name = re.sub(r"[^a-zA-Z0-9_-]+", "_", ds["name"]).strip("_") or "dataset"
                if fmt == "json":
                    raw = json.dumps(all_rows, indent=2).encode("utf-8")
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json; charset=utf-8")
                    self.send_header("Content-Disposition", f'attachment; filename="{safe_name}.json"')
                    self.send_header("Content-Length", str(len(raw)))
                    self.end_headers()
                    self.wfile.write(raw)
                    return
                buf = io.StringIO()
                col_names = [c["name"] for c in ds["columns"]]
                col_labels = [c.get("label", c["name"]) for c in ds["columns"]]
                writer = csv.writer(buf)
                writer.writerow(col_labels)
                for r in all_rows:
                    writer.writerow(["" if r.get(c) is None else r.get(c) for c in col_names])
                raw = buf.getvalue().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", f'attachment; filename="{safe_name}.csv"')
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
            return

        if path in {"/", "/platform.html", "/index.html"}:
            html_file = ROOT / ("index.html" if path == "/index.html" else "platform.html")
            if not html_file.exists():
                html_file = ROOT / "platform.html"
            raw = html_file.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        if path == "/favicon.svg":
            fav = ROOT / "favicon.svg"
            if fav.exists():
                raw = fav.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "image/svg+xml")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
                return

        self._send_json({"error": f"Not found: {path}"}, 404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/auth/signup":
            try:
                data = self._read_json_body()
                email_addr = str(data.get("email") or "").strip().lower()
                password = str(data.get("password") or "")
                category = str(data.get("category") or "ecommerce").strip().lower()
                if category not in CATEGORIES:
                    category = "ecommerce"
                if not email_addr or "@" not in email_addr:
                    self._send_json({"error": "Enter a valid email address."}, 400)
                    return
                if len(password) < 4:
                    self._send_json({"error": "Password must be at least 4 characters."}, 400)
                    return
                salt, p_hash = hash_password(password)
                with connect() as con:
                    existing = con.execute("SELECT id FROM users WHERE email = ?", (email_addr,)).fetchone()
                    if existing:
                        self._send_json({"error": "An account with this email already exists. Try signing in instead."}, 409)
                        return
                    cur = con.execute(
                        "INSERT INTO users (email, salt, pass_hash, password_hash, category, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                        (email_addr, salt, p_hash, p_hash, category, now_iso()),
                    )
                    user_id = int(cur.lastrowid)
                    cookie_hdr = self._make_session_cookie(con, user_id)
                    user = dict(con.execute("SELECT id, email, category, created_at FROM users WHERE id = ?", (user_id,)).fetchone())
                self._send_json({"user": user}, 201, cookie_header=cookie_hdr)
            except Exception as exc:
                self._send_json({"error": f"Signup failed: {exc}"}, 400)
            return

        if path == "/api/auth/login":
            try:
                data = self._read_json_body()
                email_addr = str(data.get("email") or "").strip().lower()
                password = str(data.get("password") or "")
                if not email_addr or not password:
                    self._send_json({"error": "Email and password are required."}, 400)
                    return
                with connect() as con:
                    row = con.execute("SELECT * FROM users WHERE email = ?", (email_addr,)).fetchone()
                    if not row:
                        self._send_json({"error": "No account found for that email. Use 'Create account' to register."}, 401)
                        return
                    _, calc_hash = hash_password(password, row["salt"])
                    if not secrets.compare_digest(calc_hash, row["pass_hash"]):
                        self._send_json({"error": "Incorrect password."}, 401)
                        return
                    cookie_hdr = self._make_session_cookie(con, int(row["id"]))
                    user = {
                        "id": row["id"],
                        "email": row["email"],
                        "category": row["category"],
                        "created_at": row["created_at"],
                    }
                self._send_json({"user": user}, 200, cookie_header=cookie_hdr)
            except Exception as exc:
                self._send_json({"error": f"Login failed: {exc}"}, 400)
            return

        if path == "/api/auth/logout":
            raw_cookie = self.headers.get("Cookie", "")
            cookie = SimpleCookie()
            try:
                cookie.load(raw_cookie)
                morsel = cookie.get(COOKIE_NAME)
                if morsel and morsel.value:
                    with connect() as con:
                        con.execute("DELETE FROM sessions WHERE token = ?", (morsel.value,))
            except Exception:
                pass
            clear_hdr = f"{COOKIE_NAME}=; HttpOnly; Path=/; SameSite=Lax; Max-Age=0"
            self._send_json({"ok": True}, 200, cookie_header=clear_hdr)
            return

        if path == "/api/user/category":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                data = self._read_json_body()
                cat = str(data.get("category") or "ecommerce").strip().lower()
                if cat not in CATEGORIES:
                    self._send_json({"error": f"Unknown category '{cat}'."}, 400)
                    return
                con.execute("UPDATE users SET category = ? WHERE id = ?", (cat, user["id"]))
                dataset_id = data.get("dataset_id")
                if dataset_id:
                    con.execute("UPDATE datasets SET category = ? WHERE id = ? AND user_id = ?", (cat, int(dataset_id), user["id"]))
                self._send_json({"ok": True, "category": cat, "category_label": CATEGORIES[cat]["label"]})
            return

        if path == "/api/upload":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required. Please sign in before uploading."}, 401)
                    return
                length = int(self.headers.get("Content-Length", "0"))
                if length <= 0:
                    self._send_json({"error": "Empty upload request."}, 400)
                    return
                if length > MAX_UPLOAD + 65536:
                    self._send_json({"error": "Uploaded file exceeds the 50 MB size limit."}, 413)
                    return
                body = self.rfile.read(length)
                ctype = self.headers.get("Content-Type", "")
                try:
                    fields, file_item = parse_multipart(ctype, body)
                    if not file_item:
                        self._send_json({"error": "No file was attached in the upload form."}, 400)
                        return
                    filename, raw_bytes = file_item
                    cols, records, diag = decode_upload(filename, raw_bytes)
                    cat = (fields.get("category") or user.get("category") or "ecommerce").strip().lower()
                    if cat not in CATEGORIES:
                        cat = "general"
                    custom_name = (fields.get("name") or "").strip() or Path(filename).stem.replace("_", " ").strip() or "Uploaded Dataset"
                    ds_id = save_dataset(con, user["id"], custom_name, cat, cols, records, diagnostics=diag)
                    con.execute("UPDATE users SET category = ? WHERE id = ?", (cat, user["id"]))
                    self._send_json({"ok": True, "dataset_id": ds_id, "row_count": len(records), "column_count": len(cols)}, 201)
                except ValueError as exc:
                    self._send_json({"error": str(exc)}, 400)
                except Exception as exc:
                    self._send_json({"error": f"Could not process file: {exc}"}, 400)
            return

        if path == "/api/sample":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                data = self._read_json_body()
                cat = str(data.get("category") or user.get("category") or "ecommerce").strip().lower()
                if cat not in CATEGORIES:
                    cat = "ecommerce"
                name, cols, records, diag = make_sample(cat)
                ds_id = save_dataset(con, user["id"], name, cat, cols, records, diagnostics=diag)
                con.execute("UPDATE users SET category = ? WHERE id = ?", (cat, user["id"]))
                self._send_json({"ok": True, "dataset_id": ds_id, "category": cat}, 201)
            return

        if path == "/api/dataset/meta":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                data = self._read_json_body()
                ds_id = int(data.get("dataset_id") or 0)
                ds, _ = load_dataset_and_rows(con, user["id"], ds_id)
                if not ds:
                    self._send_json({"error": "Dataset not found."}, 404)
                    return
                units = dict(ds.get("custom_units") or {})
                if isinstance(data.get("units"), dict):
                    for k, v in data["units"].items():
                        units[str(k)] = str(v or "").strip()
                overrides = dict(ds.get("role_overrides") or {})
                if isinstance(data.get("role_overrides"), dict):
                    for k, v in data["role_overrides"].items():
                        overrides[str(k)] = str(v or "").strip()
                new_cat = str(data.get("category") or ds["category"]).strip().lower()
                if new_cat not in CATEGORIES:
                    new_cat = ds["category"]
                con.execute(
                    "UPDATE datasets SET category = ?, custom_units_json = ?, role_overrides_json = ? WHERE id = ? AND user_id = ?",
                    (new_cat, json.dumps(units), json.dumps(overrides), ds["id"], user["id"]),
                )
                self._send_json({"ok": True, "dataset_id": ds["id"], "custom_units": units, "role_overrides": overrides})
            return

        if path == "/api/dataset/delete":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                data = self._read_json_body()
                ds_id = int(data.get("dataset_id") or 0)
                con.execute("DELETE FROM datasets WHERE id = ? AND user_id = ?", (ds_id, user["id"]))
                self._send_json({"ok": True})
            return

        if path == "/api/transform":
            with connect() as con:
                user = self._current_user(con)
                if not user:
                    self._send_json({"error": "Authentication required."}, 401)
                    return
                data = self._read_json_body()
                ds_id = int(data.get("dataset_id") or 0)
                action = str(data.get("action") or "").strip()
                target_col = str(data.get("column") or "").strip() or None
                ds, rows = load_dataset_and_rows(con, user["id"], ds_id)
                if not ds:
                    self._send_json({"error": "Source dataset not found."}, 404)
                    return
                try:
                    new_rows, note = transform_records(ds["columns"], rows, action, target_col)
                    if not new_rows:
                        self._send_json({"error": "This cleaning action would remove all rows from the dataset."}, 400)
                        return
                    new_name = f"{ds['name']} (Cleaned)"
                    new_id = save_dataset(
                        con,
                        user["id"],
                        new_name,
                        ds["category"],
                        ds["columns"],
                        new_rows,
                        parent_id=ds["id"],
                        transform_note=note,
                        diagnostics=ds.get("diagnostics"),
                        custom_units=ds.get("custom_units"),
                        role_overrides=ds.get("role_overrides"),
                    )
                    self._send_json({"ok": True, "dataset_id": new_id, "transform_note": note, "row_count": len(new_rows)}, 201)
                except ValueError as exc:
                    self._send_json({"error": str(exc)}, 400)
            return

        self._send_json({"error": f"Not found: {path}"}, 404)


def main() -> None:
    port = int(sys.argv[1]) if len(sys.argv) > 1 else int(os.environ.get("PORT", "8780"))
    default_host = "0.0.0.0" if os.environ.get("PORT") or os.environ.get("RENDER") else "127.0.0.1"
    host = os.environ.get("DATAPRISM_HOST", default_host)
    with connect():
        pass
    server = HTTPServer((host, port), PlatformHandler)
    print(f"DataPrism Platform listening on http://{host}:{port} (DB: {DB_PATH})")
    server.serve_forever()


if __name__ == "__main__":
    main()
