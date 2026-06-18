# Datasheets User DB API - Local and Railway

## Local Test

1. Run `run_api.cmd`
2. Open `http://127.0.0.1:8000/docs`
3. Test endpoints:
   - `GET /health`
   - `POST /users`
   - `GET /users`
   - `DELETE /users/{user_id}`

## Railway Deploy

This repo is prepared for Railway with:

- `railway.json` start command
- `requirements.txt` (which includes `requirements_api.txt`)
- database URL normalization in `app_api.py` for Railway Postgres URLs

Steps:

1. Create a new Railway project from this folder/repo.
2. Add a PostgreSQL service in Railway.
3. Set env var `DATABASE_URL` from Railway Postgres connection string.
4. Deploy.

Railway will use:

- Start command: `uvicorn app_api:app --host 0.0.0.0 --port ${PORT:-8000}`
- Dependencies from `requirements.txt`

## Cleanup Guidance

The following are now ignored by `.gitignore` and should stay local-only:

- `.venv/`
- `*.db`, `*.sqlite`, `*.sqlite3`
- generated logs and output files

If these files are already tracked in git, untrack them once using:

```bash
git rm --cached app_users.db pdf_text_map.db pdf_index_log.txt download_log_legacy.csv
```

## Sample User Payload

```json
{
  "name": "Datasheets Admin",
  "email": "admin@example.com",
  "role": "admin",
  "datasheets_root": "C:\\Users\\your-user\\OneDrive - Honeywell\\TSI Team Site (GPS Sofia) - General\\3 DATASHEETS\\Datasheets Library",
  "sharepoint_library_url": "https://honeywellprod.sharepoint.com/:f:/r/teams/TSITeamGPSSofia/Shared%20Documents/General/3%20DATASHEETS/Datasheets%20Library?csf=1&web=1&e=F1G3KV"
}
```

## Better Options (for production)

1. Keep FastAPI + Railway Postgres if you want quick setup and low ops.
2. Use Supabase (managed Postgres + auth + dashboard) if you want built-in auth quickly.
3. Use Azure App Service + Azure Database for PostgreSQL if corporate governance and AD integration matter most.

For your current workflow and speed, Railway Postgres is good for MVP.
