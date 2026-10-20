<p align="center">
  <img src="favicon.svg" width="76" height="76" alt="DataPrism mark">
</p>

<h1 align="center">DataPrism</h1>
<p align="center"><strong>Transparent data analytics for everyday datasets.</strong></p>
<p align="center">Upload a table. Understand its shape. Explore the evidence. Export a clear chart.</p>

<p align="center">
  <a href="https://www.python.org/"><img alt="Python 3.12" src="https://img.shields.io/badge/Python-3.12-3776AB?logo=python&logoColor=white"></a>
  <img alt="Analytics" src="https://img.shields.io/badge/focus-exploratory%20analytics-17694e">
  <img alt="Deployment" src="https://img.shields.io/badge/deploy-Render-46E3B7?logo=render&logoColor=black">
</p>

---

DataPrism is a portfolio web application for exploring small tabular datasets. It pairs a guided analytics workspace with a lightweight Python and SQLite backend. Results are designed to be inspectable: the workspace shows the fields used, the calculation performed, and important limits alongside the result.

> **Demo scope:** DataPrism is intended for learning, demonstrations, and small exploratory datasets. Do not upload confidential, regulated, or otherwise sensitive data.

## What you can do

| Workspace | What it helps you explore |
| --- | --- |
| **Dataset Health** | Row and column counts, detected field types, missing values, duplicate rows, and import warnings |
| **Numeric Summary** | Mean, median, sample variance and standard deviation, quartiles, range, and possible outliers |
| **Category Breakdown** | Counts, shares, and common values for useful categorical fields |
| **Trends Over Time** | Date-based series with selectable time grain and aggregation |
| **Relationships** | Numeric scatter plots and optional Pearson correlation |
| **Category Analysis** | Role-based views for common business and institutional datasets |
| **Chart Builder** | Bar, line, scatter, donut, and histogram charts with downloadable PNG, SVG, or chart data CSV |
| **Data Preview** | Search and inspect uploaded rows; create a cleaned copy using supported transformations |

Analysis categories include Sales & E-commerce, Retail & Inventory, Education & Academics, Marketing & Campaigns, Customer Analytics, Finance & Accounting, Product & SaaS, and General / Custom Dataset.

## See the analysis, not just the number

DataPrism applies practical checks before showing a result:

- Numeric and date summaries use fields inferred from the uploaded rows; identifier-like columns are kept out of ordinary metric rankings.
- Missing and unparseable values are counted and explained. They are not silently changed to zero.
- Cleaning actions create a derived dataset copy so the uploaded source remains available.
- Chart choices are checked against field types to help prevent misleading combinations.
- Category-specific measures are shown only when the required source fields exist.

See [Analytics Methodology](docs/ANALYTICS.md) for calculation details and assumptions.

## Accepted files

CSV, TSV, Excel workbooks in XLSX/XLSM format, JSON arrays or records, and JSON Lines / NDJSON. Uploads are limited by the current server configuration to **50 MB** and **100,000 rows**; larger row sets are truncated at the configured row limit.

## Run it locally

### Windows

1. Install Python 3.12.
2. Open the project folder.
3. Double-click start_dataprism.bat.
4. The launcher starts the analytics server and opens http://127.0.0.1:8785.

### Any platform

From the project directory:

    pip install -r requirements.txt
    python platform_server.py 8780

Then open http://127.0.0.1:8780. The application and its API are served from the same origin; do not open platform.html directly as a file.

### Docker Compose

From the project directory:

    docker compose up --build

Then open http://127.0.0.1:8780. Compose uses a named volume so the local SQLite database survives container restarts.

## Deploy to Render

This repository includes a Render Blueprint and Dockerfile. In Render, create or connect a Blueprint using this repository and allow Render to read render.yaml. The service builds from Dockerfile and uses /api/health as its health check.

To publish a later change, commit and push it to the branch connected to Render. Automatic deploys are enabled by default for linked repositories; a deploy can also be started from the service’s Deploys page.

**Render free-plan storage is temporary.** The current Blueprint points SQLite at /tmp/dataprism/platform.sqlite3. Accounts and uploaded datasets can be lost when the service sleeps, restarts, or redeploys. This is acceptable for a portfolio demo, but not for durable user data. See [Deployment & Operations](docs/DEPLOYMENT.md).

## How the project is put together

    Browser
      platform.html  →  same-origin HTTP API
                            ↓
      platform_server.py  →  parsing + analytics + validation
                            ↓
      SQLite  ←  accounts, sessions, datasets, and derived copies

The frontend is a single HTML file with CSS, SVG, and JavaScript. The backend uses Python’s standard library, SQLite, and openpyxl for Excel parsing. There is no separate JavaScript build step.

| File | Purpose |
| --- | --- |
| platform.html | Main DataPrism site and analytics workspace |
| platform_server.py | Main multi-user application server, API, parsing, and analytics |
| Dockerfile | Container image used by Render and Docker |
| render.yaml | Render Blueprint, health check, and environment configuration |
| compose.yaml | Local Docker setup with persistent named volume |
| start_dataprism.bat | Windows launcher for the local app |
| server.py and index.html | Earlier standalone prototype; not the primary Render service |
| docs/ANALYTICS.md | Definitions, formulas, and interpretation limits |
| docs/DEPLOYMENT.md | Local, Docker, Render, and troubleshooting notes |

## API landmarks

The browser uses the server’s same-origin API. Useful routes include:

| Method | Route | Purpose |
| --- | --- | --- |
| GET | /api/health | Service health and storage status |
| GET | /api/me | Current account and available categories |
| GET | /api/dashboard | Dataset overview and guided analyses |
| GET | /api/preview | Paginated or searched dataset preview |
| GET | /api/chart | Validated chart data |
| GET | /api/export | Export a dataset as CSV |
| POST | /api/auth/signup | Create a demo account |
| POST | /api/auth/login | Start an account session |
| POST | /api/upload | Parse and save an uploaded dataset |
| POST | /api/transform | Create a cleaned dataset copy |

## Project guides

- [Analytics methodology](docs/ANALYTICS.md)
- [Architecture and API](docs/ARCHITECTURE.md)
- [Deployment and troubleshooting](docs/DEPLOYMENT.md)

---