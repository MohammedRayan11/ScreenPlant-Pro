# ScreenPlant Pro

Flask-based ID card submission, admin review, card generation, analytics, and export tool.

## Production Stack

- App hosting: Railway
- Database: Neon PostgreSQL
- Image/file storage: Cloudflare R2
- Backend: Flask + Gunicorn
- Frontend: Single-page HTML app served by Flask

## Required Railway Environment Variables

Copy `.env.example` into Railway variables and fill real values there. Do not commit real secrets.

```env
DATABASE_URL=postgresql://...
SCREENPLANT_STORAGE_BACKEND=r2
R2_ACCOUNT_ID=...
R2_ACCESS_KEY_ID=...
R2_SECRET_ACCESS_KEY=...
R2_BUCKET_NAME=...
SCREENPLANT_SECRET_KEY=...
SCREENPLANT_ADMIN_PASSWORD=...
SCREENPLANT_HTTPS=1
```

After deploying, verify:

```text
https://your-app.up.railway.app/api/health
```

Expected production health:

```json
{
  "db": "postgres",
  "db_ok": true,
  "storage": "r2",
  "r2_configured": true
}
```

## Local Development

Install dependencies:

```bash
pip install -r backend/requirements.txt
```

Run locally:

```bash
python backend/app.py
```

Open:

```text
http://127.0.0.1:5000
```

Without `DATABASE_URL`, local development uses SQLite at `backend/screenplant.sqlite3`.

## Scale Check

Run the isolated 40k synthetic test:

```bash
python backend/tools/scale_check.py --students 40000 --schools 40
```

This does not touch production or local app data.

## Repository Layout

```text
backend/
  app.py                 Flask backend
  requirements.txt       Python dependencies
  tools/scale_check.py   Isolated scale test
frontend/
  index.html             Frontend SPA
Procfile                 Railway/Gunicorn start command
docker-compose.yml       Local PostgreSQL helper
.env.example             Environment variable template
.gitignore               Keeps runtime/private files out of Git
```

## Do Not Commit

These contain private or generated runtime data and are ignored:

- `backend/uploads/`
- `backend/backups/`
- `backend/generated/`
- `backend/templates_store/`
- `backend/database.json`
- `backend/screenplant.sqlite3`
- `.env`

