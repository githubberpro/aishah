# Migrating the AISHAH CRM database (Supabase → Neon)

Both Supabase and Neon are plain PostgreSQL, so this is a standard
`pg_dump` → `pg_restore`. The app only uses tables in the **`public`**
schema, so dump only `public` — this avoids Supabase's internal schemas
(`auth`, `storage`, `realtime`, …) that Neon neither has nor needs.

> You can run the steps below by hand, or use the helper script:
> `./migrate.sh dump` then `./migrate.sh restore` (see `migrate.sh`).

## Prerequisites

- `pg_dump` / `pg_restore` client tools, version **>= your Supabase
  Postgres major version** (e.g. Postgres 15/16 client). Check with
  `pg_dump --version`.
- Both connection strings (see below).

## 1. Get connection strings

- **Supabase:** Project → **Settings → Database → Connection string → URI**.
  Use the **direct** connection on port `5432` (not the pooled `6543` one)
  for dump/restore.
- **Neon:** Project dashboard → **Connection Details** → connection string
  (already includes `?sslmode=require`).

## 2. Dump the `public` schema from Supabase

```bash
pg_dump \
  --no-owner --no-privileges \
  --schema=public \
  -Fc \
  -d "postgresql://postgres:[PW]@db.[REF].supabase.co:5432/postgres" \
  -f aishah_backup.dump
```

- `-Fc` — compressed custom format (works with `pg_restore`).
- `--no-owner --no-privileges` — strips Supabase-specific role grants that
  Neon won't recognize.

## 3. Restore into Neon

```bash
pg_restore \
  --no-owner --no-privileges \
  -d "postgresql://[user]:[PW]@[endpoint].neon.tech/[db]?sslmode=require" \
  aishah_backup.dump
```

If Neon already has the tables (e.g. the app booted against it once and
`init_db()` created them), add `--clean --if-exists` to drop-and-recreate,
or use the **data-only** approach below instead.

## 4. Point the app at Neon

Update `DATABASE_URL` in your Streamlit secrets to the Neon connection
string and reboot the app.

```toml
DATABASE_URL = "postgresql://[user]:[PW]@[endpoint].neon.tech/[db]?sslmode=require"
```

## 5. Verify

Compare row counts on both databases:

```bash
for url in "$SUPABASE_URL" "$NEON_URL"; do
  psql "$url" -c "SELECT
    (SELECT count(*) FROM clients)       AS clients,
    (SELECT count(*) FROM opportunities) AS opportunities,
    (SELECT count(*) FROM engagements)   AS engagements;"
done
```

The numbers should match.

## Alternative: data-only migration

Since `init_db()` creates all tables automatically, you can let the app
build the schema on Neon first (just boot it once against the Neon
`DATABASE_URL`), then migrate **only the data**:

```bash
pg_dump --data-only --schema=public \
  -d "$SUPABASE_URL" \
  | psql "$NEON_URL"
```

After a data-only load, reset the auto-increment sequences so new inserts
don't collide with existing IDs:

```sql
-- run against Neon, once, after the data load
SELECT setval(pg_get_serial_sequence(t, 'id'),
              COALESCE((SELECT max(id) FROM ONLY "public".t2), 1))
FROM (VALUES
  ('clients'), ('opportunities'), ('engagements'), ('eminence'),
  ('chat_history'), ('app_context'), ('channel_partners'),
  ('opportunity_files'), ('proposals'), ('opp_chat_history'),
  ('cp_chat_history'), ('eminence_chat_history'), ('content_plan')
) AS x(t), LATERAL (SELECT t::regclass AS t2) y;
```

> A full (schema + data) `pg_restore` — steps 2–3 — carries sequences over
> automatically, so the reset above is only needed for the data-only path.

## Gotchas

- **Version mismatch:** `pg_dump: server version X; pg_dump version Y` means
  your client is older than the server — upgrade the Postgres client tools.
- **Pooled vs direct:** dumps against Supabase's pooled port `6543` can fail
  or hang; always use the direct `5432` endpoint.
- **SSL:** Neon requires `sslmode=require` (already in its connection
  string). The app also enforces this in `database.py`.
