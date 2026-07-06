# AISHAH CRM

A Streamlit-based CRM with AI features powered by the Anthropic Claude API.
It tracks clients, opportunities, engagements, channel partners, proposals,
eminence/content plans, and per-record AI chat histories.

## Tech stack

- **Frontend / app:** [Streamlit](https://streamlit.io) (`app.py`)
- **Data layer:** `database.py` — auto-detects the backend:
  - **PostgreSQL** when `DATABASE_URL` is set (production — recommended)
  - **SQLite** (`crm.db`) as a local dev fallback when it isn't
- **AI:** Anthropic Claude API (`anthropic`)
- **Docs parsing:** `pdfplumber`, `python-docx`, `openpyxl`

The schema is created automatically on first run by `init_db()`
(`app.py` calls it once per server process). All table creation is
idempotent (`CREATE TABLE IF NOT EXISTS`) and column migrations are
safe to re-run, so pointing the app at a fresh empty database is all
that's needed to provision it.

## Configuration (secrets)

The app reads these from `st.secrets` (Streamlit Cloud) or environment
variables (local):

| Key | Purpose | Required |
| --- | --- | --- |
| `DATABASE_URL` | PostgreSQL connection string. If omitted, falls back to local SQLite `crm.db`. | Recommended for any real deployment |
| `ANTHROPIC_API_KEY` | Claude API key — enables all AI features | Required for AI features |
| `APP_PASSWORD` | Password gate for the app | Optional (recommended) |

> **Note:** On Streamlit Community Cloud the filesystem is ephemeral, so the
> SQLite fallback is wiped on every restart/redeploy. Use a hosted Postgres
> (`DATABASE_URL`) for persistent data.

## Run locally

```bash
pip install -r requirements.txt

# Option A — quick local run with the SQLite fallback (no DATABASE_URL):
export ANTHROPIC_API_KEY="sk-ant-..."
streamlit run app.py

# Option B — run against Postgres:
export DATABASE_URL="postgresql://user:pass@host:5432/dbname"
export ANTHROPIC_API_KEY="sk-ant-..."
streamlit run app.py
```

Alternatively, create `.streamlit/secrets.toml` (git-ignored) instead of
exporting env vars:

```toml
DATABASE_URL     = "postgresql://user:pass@host:5432/dbname"
ANTHROPIC_API_KEY = "sk-ant-..."
APP_PASSWORD      = "your-password"
```

## Deploy to Streamlit Community Cloud

1. Sign in at <https://share.streamlit.io> with GitHub.
2. **Create app → Deploy a public app from GitHub.**
3. Settings:
   - **Repository:** `githubberpro/aishah`
   - **Branch:** `master`
   - **Main file path:** `app.py`
4. **Advanced settings → Secrets** — paste (TOML):
   ```toml
   DATABASE_URL      = "postgresql://user:pass@host:5432/dbname"
   ANTHROPIC_API_KEY = "sk-ant-..."
   APP_PASSWORD      = "your-password"
   ```
5. **Deploy.** The first build installs `requirements.txt`; `init_db()`
   provisions the schema on first load.

## Provisioning a PostgreSQL database

Any managed Postgres works. Free options that give you a `DATABASE_URL`:

- **[Neon](https://neon.tech)** — serverless Postgres
- **[Supabase](https://supabase.com)** — Postgres with a free tier
- **[Railway](https://railway.app)** / **[Render](https://render.com)**

Create a database, copy its connection string into `DATABASE_URL`, and the
app will create every table on first run.

## Repository layout

```
app.py                  # Streamlit UI and all pages/features
database.py             # Data access layer + schema + migrations
requirements.txt        # Python dependencies
.streamlit/config.toml  # Theme + server config
```
