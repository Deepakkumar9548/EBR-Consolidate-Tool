# EBR Consolidate Tool

An AI-assisted raw data consolidation and master converter dashboard for automotive dealership DMS reports (**Enquiry**, **Booking**, **Retail**, **Finance**, and **Accessories**), combining a browser-based interactive UI with a Python chunk-streaming backend engine.

---

## Project Structure

| File | Description |
| :--- | :--- |
| [`start_app.bat`](start_app.bat) | Windows launcher that opens `http://127.0.0.1:8000` and starts `server.py`. |
| [`server.py`](server.py) | Python HTTP server (`http.server`) + `openpyxl` read-only chunk-streaming pipeline backed by SQLite WAL (`ebr_enterprise.db`). |
| [`ebr_zip_converter_modern (2).html`](ebr_zip_converter_modern%20(2).html) | Complete frontend UI and client-side parser/standardizer using SheetJS (`xlsx`) and `JSZip`. |
| [`.gitignore`](.gitignore) | Excludes runtime SQLite databases (`*.db`, `*.db-shm`, `*.db-wal`), `__pycache__/`, `uploads/`, and `exports/`. |

---

## Key Features

### 1. Multi-Format Ingestion (Frontend UI)
- Supports single/multiple files, entire dealership folders, and nested `.zip` archives containing `.xlsx`, `.xls`, `.xlsm`, or `.csv` files.
- **1-Click Auto-Consolidate & Process** or step-by-step categorization and cleaning.

### 2. Intelligent Header Detection & Semantic Column Mapping
- Scans the first 30 rows to detect the real header row and strips repeated mid-sheet headers.
- Matches DMS column variations using exact matching, substring matching, Levenshtein fuzzy similarity, and data-pattern inspection (e.g., Indian 10-digit mobile numbers, `EN-`/`ENQ`/`LD-`, `BK`/`OTF`/`SO`, `INV`/`TAX`/`BL` patterns).
- **Enquiry Type Priority (Enquiry Module):** Explicitly prioritizes **`Enquiry Type`** (`ENQUIRY_TYPE`) over `Enquiry Source` (`ENQUIRY_SOURCE`) in the Enquiry module, with a safe fallback to `Enquiry Source` when `Enquiry Type` is absent or empty.

### 3. Standardized Output Schemas
- **Enquiry:** `Month`, `Model Group`, `Model Varraint`, `Contact_Number`, `Source`, `Enquiry_Date`, `Enquiry Number`, `Dealer`, `Location`, `Region`, `PV/CV/EV`
- **Booking:** `Month`, `Booking Date and Time`, `Booking Status`, `Contact_Number`, `Model Group`, `Model Variant`, `Dealership`, `Location`, `Region`, `PV/CV/EV`
- **Retail:** `Month`, `Invoice Date and Time`, `Invoice Status`, `Contact_Number`, `Model Group`, `Model Variant`, `DELR_NAME_DESC`, `Location`, `Region`, `PV/CV/EV`
- **Finance / Accessories / Raw Consolidated:** Preserves raw source columns alongside `__SourceFile` and `__SourceSheet`.

### 4. Interactive Review, Editing & Export
- **Conversion Funnel & KPIs:** Displays Enquiry → Booking → Retail counts/conversion %, top model, top dealership, and 10-digit mobile validity quality score.
- **In-Line Sheet Viewer/Editor:** Edit cell values, add/delete rows, or delete columns directly in the browser before exporting.
- **Custom Mapping Rules:** Editable & JSON importable/exportable mappings for **Dealer → Region** and **Model Keyword → PV/CV/EV**.
- **Audit Logs:** Per-sheet **Processing Log** (with clickable inspection modals for source rows, imported records, empty rows, and missing critical fields) and **Error & Anomaly Review Log** (e.g., invalid dates, repeated headers).
- **Multiple Export Formats:** Single Excel Workbook (`.xlsx`), individual `.csv` files, or a `.zip` bundle of CSVs.

---

## Python Backend Engine (`server.py`)

Runs locally on `http://127.0.0.1:8000` and streams large `.xlsx` files in **50,000-row chunks** into SQLite (`ebr_enterprise.db`) with checkpoint tracking (`jobs`, `job_checkpoints`, `master_data`, `error_data`).

### REST API Endpoints
- `GET /` — Serves `ebr_zip_converter_modern (2).html`.
- `POST /api/upload` — Accepts `multipart/form-data` Excel upload and starts a background chunk-processing thread.
- `GET /api/jobs/<job_id>/status` — Returns job progress, processed row counts, and chunk status.
- `GET /api/records?job_id=...&category=...&page=1&limit=50&search=...` — Paginated record query & search.
- `GET /api/reports?job_id=...` — Aggregated counts by category, top 10 models, and top 10 dealers.

---

## Quick Start

### Option 1: Run with Local Python Backend (Recommended)
1. Ensure **Python 3** and `openpyxl` are installed:
   ```powershell
   pip install openpyxl
   ```
2. Double-click [`start_app.bat`](start_app.bat) or run from terminal:
   ```powershell
   python server.py
   ```
3. Open `http://127.0.0.1:8000` in your browser.

### Option 2: Standalone Browser Mode
Open [`ebr_zip_converter_modern (2).html`](ebr_zip_converter_modern%20(2).html) directly in Chrome/Edge (requires internet access for the CDN-hosted `JSZip`, `SheetJS`, and Google Fonts scripts). All browser-mode processing stays locally in browser memory.
