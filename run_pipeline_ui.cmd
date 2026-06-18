@echo off
setlocal

if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install -r requirements_api.txt
python -m pip install pandas openpyxl pymupdf

if not defined PDF_INDEX_DB (
  set "PDF_INDEX_DB=%LOCALAPPDATA%\DatasheetsScraper\pdf_text_map.db"
)
if not defined PDF_LOCAL_LIBRARY_ROOT (
  set "PDF_LOCAL_LIBRARY_ROOT=C:\Users\H588238\OneDrive - Honeywell\TSI Team Site (GPS Sofia) - General\3 DATASHEETS\Datasheets Library"
)

python pipeline_ui.py

endlocal
