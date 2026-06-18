# Datasheets Scraper

This project helps map part/model records to datasheet PDFs, then copy the matched files into a target folder structure.
It also contains a small FastAPI app for storing user configuration in SQLite/PostgreSQL.

## What This Project Does

Main workflow:
1. Index local datasheet PDFs into a SQLite database.
2. Search PDF names and first-page text using part/model values from Excel.
3. Write found PDF paths back into a result Excel file.
4. Copy matched local files into organized output folders.

There is also a separate API service for user records and deployment to Railway.

## Quick Start

1. Run `run_pipeline_ui.cmd`.
2. Choose your local Datasheets Library folder and Excel file.
3. Click `Run All`.

Step 1 will automatically create/update your local DB (per user/computer), then run the Excel mapping search.
DB build/update runs with prune enabled, so files removed from disk are also removed from the DB.

## Files and Functions

### `01_build_pdf_map.py`
- Scans the local Datasheets Library for PDF files (filtered by folder names matching `datasheet`/`datasheets`).
- Extracts first-page text using PyMuPDF (`fitz`).
- Stores metadata in SQLite (`pdf_index` table): file path, size, SHA-256 hash, normalized filename/text, modified time.
- Applies deduplication logic (keeps newer duplicate version by hash/mtime).
- Supports prune mode (`--prune-missing`) to remove stale DB rows for files that no longer exist.
- Writes progress/errors to `pdf_index_log.txt`.

### `02_file_search_PDF.py`
- Loads `PART_MODEL_LIST.xlsm`/`.xlsx` (or custom `--excel-file`).
- Reads indexed PDF metadata from SQLite (`--database-file` or auto-detected DB).
- Normalizes keywords from Excel column B and matches against:
  - normalized PDF filename, then
  - normalized first-page text.
- Writes results to column C and saves as `PART_MODEL_with_results_PDF.xlsx` (or custom `--output-excel`).

### `03_download_sharepoint_links.ps1`
- Reads the output Excel and resolves source paths from the mapped link column.
- Operates in local-path mode (copy from local/synced datasheets root).
- Creates output folders, copies files, and handles name collisions (`Skip`, `Rename`, `Overwrite`).
- Produces CSV log with status (`OK`, `FAIL`, `SKIP`).
- Supports dry run and status-based row filtering.

### `pipeline_ui.py`
- Desktop UI built with Tkinter to run the pipeline without command-line usage.
- Lets user select:
  - local datasheets root,
  - part/model Excel file,
  - optional database file.
- Executes Step 1 (build/update local DB using `01_build_pdf_map.py`, then run `02_file_search_PDF.py`) and Step 2 (`03_download_sharepoint_links.ps1`) with live logs.
- Installs missing Python packages for pipeline operation when needed.

### `app_api.py`
- FastAPI service for managing app users in a database.
- Endpoints:
  - `GET /health`
  - `POST /users`
  - `GET /users`
  - `DELETE /users/{user_id}`
- Uses SQLAlchemy ORM model `AppUser`.
- Supports SQLite locally and PostgreSQL on Railway.
- Normalizes `postgres://` and `postgresql://` URLs to `postgresql+psycopg://`.

### Supporting Files
- `requirements.txt`: base requirements entry point.
- `requirements_api.txt`: pinned API/database package versions (FastAPI, SQLAlchemy, psycopg, etc.).
- `run_pipeline_ui.cmd`: creates/activates `.venv`, installs dependencies, sets local defaults (`PDF_INDEX_DB`, `PDF_LOCAL_LIBRARY_ROOT`), launches `pipeline_ui.py`.
- `run_api.cmd`: creates/activates `.venv`, installs API dependencies, runs Uvicorn locally.
- `railway.json`: Railway start command and restart policy.
- `RAILWAY_NOTES.md`: local test and deployment notes.

## Technologies Used

### Languages and Scripting
- Python 3
- PowerShell
- Windows batch (`.cmd`)

### Python Frameworks and Libraries
- FastAPI (REST API)
- Uvicorn (ASGI server)
- SQLAlchemy 2.x (ORM/database access)
- psycopg 3 (PostgreSQL driver)
- Pydantic + email-validator (request/data validation)
- pandas + openpyxl (Excel I/O)
- PyMuPDF / `fitz` (PDF text extraction)
- Tkinter (desktop GUI)
- sqlite3 (built-in local database)

### Data and Storage
- SQLite for local PDF index and local API DB.
- PostgreSQL (Railway) for deployed API.

### Platform and Deployment
- Railway (cloud deployment)
- OneDrive-synced SharePoint library path as local datasheet source

## Typical Usage

### Option A: UI (recommended)
1. Run `run_pipeline_ui.cmd`.
2. Select local datasheets folder, Excel file, and optional DB file.
3. Click `Run All`.

Step 1 now creates/updates a local per-user DB automatically before search.

### Option B: Script by script
1. Build or refresh index:
  - `python 01_build_pdf_map.py --database-file "%LOCALAPPDATA%\DatasheetsScraper\pdf_text_map.db" --root-folder "<datasheets-root>" --prune-missing`
2. Map Excel entries to PDF paths:
   - `python 02_file_search_PDF.py --excel-file <path-to-xlsx-or-xlsm> --local-library-root <datasheets-root> --database-file "%LOCALAPPDATA%\DatasheetsScraper\pdf_text_map.db"`
3. Copy matched files:
   - `powershell -File 03_download_sharepoint_links.ps1 -ExcelFile <PART_MODEL_with_results_PDF.xlsx> -DatasheetsLibraryRoot <datasheets-root>`

### API local run
1. Run `run_api.cmd`.
2. Open Swagger UI at `http://127.0.0.1:8000/docs`.

## Notes
- The downloader script currently works with local file paths from the Excel results.
- For best performance and matching quality, keep the PDF index database up to date by re-running `01_build_pdf_map.py` after library changes.
- If `.venv` was removed, run `run_pipeline_ui.cmd` or `run_api.cmd`; each script recreates `.venv` automatically when missing.
- The PDF index DB should be local per user (default: `%LOCALAPPDATA%\DatasheetsScraper\pdf_text_map.db`) because it stores local machine file paths.
- Avoid storing `pdf_text_map.db` on shared locations/drives, because mapped file paths are machine-specific.
- When prune is enabled (`--prune-missing`), deleted PDFs are removed from the DB on the next run.
