#!/usr/bin/env bash
#
# migrate.sh — dump the AISHAH CRM database from one Postgres and restore
# it into another (e.g. Supabase -> Neon). See MIGRATION.md for details.
#
# Usage:
#   export SUPABASE_URL="postgresql://postgres:[PW]@db.[REF].supabase.co:5432/postgres"
#   export NEON_URL="postgresql://[user]:[PW]@[endpoint].neon.tech/[db]?sslmode=require"
#
#   ./migrate.sh dump      # dump public schema from SUPABASE_URL -> aishah_backup.dump
#   ./migrate.sh restore   # restore aishah_backup.dump -> NEON_URL
#   ./migrate.sh data-only # stream data (no schema) from SUPABASE_URL -> NEON_URL
#   ./migrate.sh verify    # compare row counts on both databases
#   ./migrate.sh all       # dump then restore
#
# Requires: pg_dump / pg_restore / psql, version >= the Supabase server's
# Postgres major version.

set -euo pipefail

DUMP_FILE="${DUMP_FILE:-aishah_backup.dump}"
CMD="${1:-}"

die() { echo "error: $*" >&2; exit 1; }

require() {
  for bin in "$@"; do
    command -v "$bin" >/dev/null 2>&1 || die "'$bin' not found on PATH"
  done
}

do_dump() {
  require pg_dump
  [ -n "${SUPABASE_URL:-}" ] || die "set SUPABASE_URL (the source connection string)"
  echo ">> dumping public schema from source -> $DUMP_FILE"
  pg_dump \
    --no-owner --no-privileges \
    --schema=public \
    -Fc \
    -d "$SUPABASE_URL" \
    -f "$DUMP_FILE"
  echo ">> done: $DUMP_FILE ($(du -h "$DUMP_FILE" | cut -f1))"
}

do_restore() {
  require pg_restore
  [ -n "${NEON_URL:-}" ] || die "set NEON_URL (the destination connection string)"
  [ -f "$DUMP_FILE" ] || die "$DUMP_FILE not found — run './migrate.sh dump' first"
  echo ">> restoring $DUMP_FILE -> destination"
  # --clean --if-exists lets you re-run safely if the tables already exist.
  pg_restore \
    --no-owner --no-privileges \
    --clean --if-exists \
    -d "$NEON_URL" \
    "$DUMP_FILE"
  echo ">> restore complete"
}

do_data_only() {
  require pg_dump psql
  [ -n "${SUPABASE_URL:-}" ] || die "set SUPABASE_URL (the source connection string)"
  [ -n "${NEON_URL:-}" ]     || die "set NEON_URL (the destination connection string)"
  echo ">> streaming data-only (public schema) from source -> destination"
  echo ">> (tables must already exist on the destination; boot the app once to create them)"
  pg_dump --data-only --schema=public -d "$SUPABASE_URL" | psql "$NEON_URL"
  echo ">> data load complete — remember to reset sequences (see MIGRATION.md)"
}

do_verify() {
  require psql
  [ -n "${SUPABASE_URL:-}" ] || die "set SUPABASE_URL"
  [ -n "${NEON_URL:-}" ]     || die "set NEON_URL"
  local sql="SELECT
    (SELECT count(*) FROM clients)       AS clients,
    (SELECT count(*) FROM opportunities) AS opportunities,
    (SELECT count(*) FROM engagements)   AS engagements;"
  echo ">> source (SUPABASE_URL):"
  psql "$SUPABASE_URL" -c "$sql"
  echo ">> destination (NEON_URL):"
  psql "$NEON_URL" -c "$sql"
  echo ">> row counts should match."
}

case "$CMD" in
  dump)       do_dump ;;
  restore)    do_restore ;;
  data-only)  do_data_only ;;
  verify)     do_verify ;;
  all)        do_dump; do_restore ;;
  *)
    grep -E '^#( |$)' "$0" | sed -E 's/^# ?//'
    exit 1
    ;;
esac
