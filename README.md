# DataPrism — Transparent Data Analytics Platform

DataPrism is a full-stack portfolio web application for exploring small tabular datasets (`.csv`, `.tsv`, `.xlsx`, `.json`, `.jsonl`). It combines a zero-framework Python HTTP + SQLite backend with a responsive HTML/CSS/SVG/JavaScript frontend designed to make data analysis clear to beginners while reflecting how a careful data analyst works.

Every analysis section and chart explicitly documents:
- **What the app calculated**
- **Which columns it used**
- **How each result was calculated**
- **What the result may mean**
- **What assumptions or limitations apply**

## Repository Architecture

| File | Role |
| :--- | :--- |
| `platform_server.py` | Primary multi-user backend server (Python standard library + `sqlite3` + `openpyxl` for `.xlsx`). Handles authentication, session isolation, file parsing, type inference, the 6-section analytics engine, chart validation, and data cleaning snapshots. |
| `platform.html` | Primary single-page frontend interface for `platform_server.py`, featuring the 6 guided analytics views, interactive Chart Builder (with PNG, SVG, and CSV exports), and Data Preview / Dataset Manager. |
| `server.py` & `index.html` | Standalone local single-workspace e-commerce prototype (`margimap.sqlite3`). |
| `Dockerfile` & `compose.yaml` | Container image definition (`python:3.12-slim` on port `8780`) and Docker Compose setup with a named volume for local persistence. |
| `render.yaml` | Render Blueprint configuration (`plan: free`, Docker runtime, health check at `/api/health`, SQLite path at `/tmp/dataprism/platform.sqlite3`). |
| `start_dataprism.bat` | One-click Windows launcher that checks Python + `openpyxl`, starts `platform_server.py` on port `8785`, waits for `/api/health`, and opens the browser. |

## Core Analytics Workflow

1. **Dataset Health (`1. Dataset Health`)**
   - Reports row count, column count, completeness percentage, missing cell counts, and exact full-row duplicates.
   - Infers column types (`number`, `date`, `category`, `id`, `text`) using a $\ge 75\%$ valid-parse threshold and identifier heuristics (column naming patterns like `_id`/`sku` or near-100% unique codes).
   - Flags unparseable numeric/date cells and ambiguous slash dates (`MM/DD/YYYY` vs. `DD/MM/YYYY`).
   - Clearly separates **read-only diagnostic warnings** from **data cleaning actions** (removing duplicates, dropping missing rows, filling numeric blanks with the median, or trimming whitespace), which always create a **new derived dataset copy** so the original upload is never overwritten.

2. **Numeric Summary (`2. Numeric Summary`)**
   - Computes valid count, missing count, arithmetic mean, median ($Q_2$), sample standard deviation ($s$), sample variance ($s^2$), minimum, 25th percentile ($Q_1$), 75th percentile ($Q_3$), and maximum for quantitative numeric columns.
   - Excludes identifier columns (`order_id`, `student_id`, sequential row numbers) from math summaries.
   - Labels units as `Unknown (not specified)` by default and lets users save custom units (e.g., `USD ($)`, `units`, `%`, `days`).
   - Explains when extreme outliers pull the mean and variance away from the median and interquartile range ($Q_1\text{–}Q_3$).

3. **Category Breakdown (`3. Category Breakdown`)**
   - Shows category counts, percentage of valid (non-missing) rows, percentage of total dataset rows, distinct category count, and most common values.
   - Suppresses identifier columns from category frequency charts and explains why unique IDs are not meaningful analytical groupings.

4. **Trends Over Time (`4. Trends Over Time`)**
   - Enabled only when at least one valid date column exists.
   - Lets users choose the date column, numeric measure (or row count), time grouping (`Daily`, `Weekly`, `Monthly`, `Quarterly`, `Yearly`), and aggregation (`Sum`, `Average`, `Median`, `Record Count`, `Min`, `Max`).
   - Explains the exact date normalization (`YYYY-MM-DD`), bucketing rule, and skipped missing rows.

5. **Relationships (`5. Relationships`)**
   - Lets users pick any two quantitative numeric columns for a scatter plot and toggle **Pearson correlation ($r$)**.
   - Validates suitability ($n \ge 3$ paired numeric rows and non-zero variance in both columns) before calculating $r$ and $R^2$.
   - Prominently explains that **correlation describes linear association and does not prove causation**.

6. **Category-Specific Analysis (`6. Category Analysis`)**
   - Suggests domain analyses based on the selected category (**Sales & E-commerce**, **Retail & Inventory**, **Education & Academics**, **Marketing & Campaigns**, **Customer Analytics**, **Finance & Accounting**, **Product & SaaS**, or **General**).
   - Displays an **editable Column Role Mapping** panel so assumptions are visible and customizable.
   - Never invents metrics from unrelated columns: for example, Sales profit/margin is calculated only when a `Profit` or `Cost/COGS` column exists; if absent, DataPrism explains which column is missing and provides a general breakdown instead.

7. **Validated Chart Builder (`7. Chart Builder`)**
   - Supports **Bar**, **Line**, **Scatter**, **Donut**, and **Histogram** charts with X-axis, Y-axis (or Row Count), aggregation method, optional Grouping/Color field, and custom titles.
   - Validates compatibility and blocks misleading combinations (such as scatter plots on text columns, donut charts with negative numbers, or summing an ID column).
   - Exports charts as **PNG**, **SVG**, or **aggregated CSV data**.

## Deployment Persistence & Limitations

- **Render Free Plan Ephemeral Storage (`render.yaml`)**:
  The Render Blueprint uses `plan: free` and stores SQLite data at `/tmp/dataprism/platform.sqlite3`. On Render's free tier, container filesystem storage is **temporary (ephemeral)**: saved user accounts and uploaded datasets will disappear whenever the instance sleeps after inactivity, restarts, or redeploys. Both the `/api/health` endpoint and a top banner in the web app surface this notice automatically.
- **Persistent Local / Docker Storage**:
  When running locally with Python (`data/platform.sqlite3`) or via `docker compose up` (which mounts the `margimap-data` volume to `/app/data`), your accounts and datasets persist across restarts.
- **Scale & Security Scope**:
  DataPrism is designed for small-to-medium exploratory datasets (up to 50 MB / 100,000 rows per upload) processed in Python memory and stored as JSON records in SQLite. Passwords use PBKDF2-HMAC-SHA256 (310,000 iterations) and sessions use `HttpOnly` cookies (`SameSite=Lax`). If deploying over HTTPS in production behind a reverse proxy, consider enabling TLS termination and persistent disk storage.

## Running Locally

### Option 1: Python CLI (macOS, Linux, or Windows)

```bash
pip install -r requirements.txt
python platform_server.py 8780
```

Then open `http://127.0.0.1:8780` in your browser.

### Option 2: Windows One-Click Launcher

Double-click `start_dataprism.bat` (starts `platform_server.py` on port `8785`, verifies `/api/health`, and opens `http://127.0.0.1:8785`).

### Option 3: Docker Compose (with Persistent Volume)

```bash
docker compose up --build
```

Then open `http://127.0.0.1:8780`.

## Deploying to Render

1. Push the repository to GitHub.
2. In the Render Dashboard, click **New +** -> **Blueprint** and connect your `DataPrism` repository.
3. Render will read `render.yaml`, build the `Dockerfile`, verify `/api/health`, and deploy the web service on the free plan.
