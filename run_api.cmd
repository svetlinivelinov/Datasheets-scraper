@echo off
setlocal

if not exist .venv (
  py -m venv .venv
)

call .venv\Scripts\activate.bat
pip install -r requirements_api.txt

set "DATABASE_URL=sqlite:///./app_users.db"
uvicorn app_api:app --host 127.0.0.1 --port 8000 --reload

endlocal
