# Architecture and API

## Runtime shape

DataPrism is a same-origin web application. The browser loads platform.html, and its JavaScript requests JSON from platform_server.py. The server parses files, performs analytics, validates chart requests, and stores account and dataset records in SQLite.

    Browser
      ├── platform.html: HTML, CSS, SVG, and JavaScript
      └── /api/* requests
             ↓
    platform_server.py
      ├── Python standard library HTTP server
      ├── CSV / TSV / JSON / JSONL parsers
      ├── openpyxl for XLSX / XLSM
      ├── type inference and analytics functions
      └── SQLite persistence

The main service has no frontend package manager or compilation step. Docker copies the project into a Python 3.12 image and starts platform_server.py. Render supplies the PORT environment variable; the server binds to DATAPRISM_HOST (0.0.0.0 in the Render Blueprint).

## Main components

| Component | Responsibility |
| --- | --- |
| platform.html | Public overview, sign-in, category selection, dataset workspace, chart rendering and downloads |
| platform_server.py | HTTP routes, sessions, upload parsing, type inference, analysis, chart validation, and SQLite storage |
| requirements.txt | Python dependency: openpyxl for Excel workbook parsing |
| Dockerfile | Installs requirements, copies application files, and starts the main server |
| render.yaml | Render web service Blueprint and health-check configuration |
| compose.yaml | Local Docker service and named database volume |
| server.py / index.html | Legacy single-workspace prototype; not used by the Render Blueprint |

## Data path

1. The user creates an account or signs in.
2. The user selects an analysis category and uploads a supported file or loads a sample.
3. The server parses and normalizes the rows, stores the dataset under the account, and returns schema and diagnostics.
4. Dashboard, preview, analysis, and chart requests run against the selected dataset.
5. Cleaning operations create a derived dataset copy; chart and CSV exports are generated when requested.

The current API uses cookie-based sessions. Session cookies are HttpOnly and SameSite=Lax. Passwords are stored as salted PBKDF2-HMAC-SHA256 digests. The local app uses SQLite at data/platform.sqlite3 by default; DATAPRISM_DB can point to another database path.

## API reference

| Method | Path | Purpose |
| --- | --- | --- |
| GET | /api/health | Health status and storage persistence information |
| GET | /api/me | Current session and supported categories |
| GET | /api/dashboard | Dataset overview and calculated analysis sections |
| GET | /api/preview | Dataset rows, columns, search, and pagination |
| GET | /api/chart | Validated chart data for the selected dataset |
| GET | /api/export | Dataset CSV download |
| POST | /api/auth/signup | Create a user and session |
| POST | /api/auth/login | Authenticate and start a session |
| POST | /api/auth/logout | End the current session |
| POST | /api/user/category | Save a user’s selected analysis category |
| POST | /api/upload | Upload and parse a dataset |
| POST | /api/sample | Load a generated category sample |
| POST | /api/dataset/meta | Save dataset metadata such as labels or role overrides |
| POST | /api/dataset/delete | Delete a dataset from the current account |
| POST | /api/transform | Create a cleaned dataset copy |

Responses use JSON for application data and errors, except direct CSV exports. Authentication is required for account-specific dataset routes. The browser calls same-origin relative paths so frontend and backend must be hosted together for this setup.

## Change boundaries

The primary deployment uses platform.html and platform_server.py. index.html and server.py are retained as an earlier prototype. When adding frontend assets, either embed them in platform.html or add explicit safe static-file routes in platform_server.py; arbitrary files are not automatically served by the current route handler.