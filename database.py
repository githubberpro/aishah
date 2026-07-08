import os
import re
import json
import threading
from datetime import datetime, date, timedelta
from pathlib import Path

# ── Backend detection (lazy so env var can be injected before first use) ───────
try:
    import psycopg2
    import psycopg2.extras
    import psycopg2.pool as _pgpool
    _HAS_PG = True
except ImportError:
    _HAS_PG = False

import sqlite3
DB_PATH = Path(__file__).parent / "crm.db"

_USE_PG: bool | None = None
_DB_URL: str = ""

# ── Connection pool (PostgreSQL) / thread-local (SQLite) ───────────────────────
_pg_pool: "_pgpool.ThreadedConnectionPool | None" = None
_pool_lock = threading.Lock()
_thread_local = threading.local()


def _setup() -> bool:
    global _USE_PG, _DB_URL
    if _USE_PG is None:
        _DB_URL = os.environ.get("DATABASE_URL", "")
        _USE_PG = bool(_DB_URL and _HAS_PG)
    return _USE_PG


def _get_pg_pool():
    """Return the shared connection pool, creating it once on first call."""
    global _pg_pool
    with _pool_lock:
        if _pg_pool is None or _pg_pool.closed:
            url = _DB_URL
            if "sslmode" not in url:
                url += ("&" if "?" in url else "?") + "sslmode=require"
            _pg_pool = _pgpool.ThreadedConnectionPool(
                minconn=1,
                maxconn=10,
                dsn=url,
                cursor_factory=psycopg2.extras.RealDictCursor,
                # TCP keepalives prevent Supabase from timing out idle connections
                keepalives=1,
                keepalives_idle=30,
                keepalives_interval=10,
                keepalives_count=5,
            )
    return _pg_pool


def _ph() -> str:
    return "%s" if _setup() else "?"


def get_conn():
    if _setup():
        pool = _get_pg_pool()
        # Serverless Postgres (e.g. Neon) suspends the compute when idle and
        # closes its connections. Those sockets linger in the pool as stale
        # entries that fail mid-query with "SSL connection has been closed
        # unexpectedly". A rollback() alone doesn't reliably detect a socket the
        # server killed, so validate each pooled connection with a real
        # round-trip (SELECT 1) and discard dead ones until we get a live one.
        last_err = None
        for _ in range(10):
            conn = pool.getconn()
            try:
                if conn.closed:
                    raise Exception("stale pooled connection")
                conn.rollback()  # clear any dangling transaction
                with conn.cursor() as cur:
                    cur.execute("SELECT 1")
                    cur.fetchone()
                return conn
            except Exception as e:
                last_err = e
                try:
                    pool.putconn(conn, close=True)
                except Exception:
                    pass
        # Couldn't validate any pooled connection — surface the last error.
        raise last_err if last_err else Exception("no live DB connection")
    # SQLite: reuse a single persistent connection per thread
    conn = getattr(_thread_local, "sqlite_conn", None)
    if conn is None:
        conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        _thread_local.sqlite_conn = conn
    return conn


def _n(sql: str) -> str:
    """Convert :name placeholders → %(name)s for psycopg2."""
    if _setup():
        return re.sub(r":(\w+)", r"%(\1)s", sql)
    return sql


def _q(conn, sql: str, params=()):
    """Execute and return all rows as dicts."""
    if _setup():
        with conn.cursor() as cur:
            cur.execute(sql, params or None)
            return [dict(r) for r in cur.fetchall()]
    return [dict(r) for r in conn.execute(sql, params).fetchall()]


def _q1(conn, sql: str, params=()):
    """Execute and return first row as dict, or None."""
    if _setup():
        with conn.cursor() as cur:
            cur.execute(sql, params or None)
            row = cur.fetchone()
            return dict(row) if row else None
    row = conn.execute(sql, params).fetchone()
    return dict(row) if row else None


def _run(conn, sql: str, params=()):
    """Execute DML; returns lastrowid for INSERT statements."""
    if _setup():
        is_insert = sql.strip().upper().startswith("INSERT") and "RETURNING" not in sql.upper()
        exec_sql = (sql.rstrip(";") + " RETURNING id") if is_insert else sql
        with conn.cursor() as cur:
            cur.execute(exec_sql, params or None)
            if is_insert:
                row = cur.fetchone()
                return row["id"] if row else None
        return None
    cur = conn.execute(sql, params)
    return cur.lastrowid


def _commit(conn):
    conn.commit()


def _close(conn):
    if _setup():
        # Return to pool rather than closing — keeps the TCP connection alive
        pool = _pg_pool
        if pool is not None and not pool.closed:
            try:
                if not conn.closed:
                    conn.rollback()  # Clean up any uncommitted state
                pool.putconn(conn)
                return
            except Exception:
                pass
        # Pool gone or error — just close
        try:
            conn.close()
        except Exception:
            pass
    # SQLite: leave open; thread-local connection is reused next call


# ── Schema ─────────────────────────────────────────────────────────────────────

_SQLITE_SCHEMA = """
CREATE TABLE IF NOT EXISTS clients (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    company            TEXT NOT NULL,
    sector             TEXT NOT NULL,
    sub_sector         TEXT,
    company_size       TEXT,
    buyer_type         TEXT NOT NULL CHECK(buyer_type IN ('Institutional','Owner')),
    country            TEXT DEFAULT 'Singapore',
    key_contact        TEXT,
    contact_title      TEXT,
    relationship_score INTEGER DEFAULT 3 CHECK(relationship_score BETWEEN 1 AND 5),
    ai_maturity        TEXT,
    notes              TEXT,
    annual_revenue     TEXT,
    employee_count     TEXT,
    website            TEXT,
    business_strategy  TEXT,
    business_challenges TEXT,
    competitive_landscape TEXT,
    account_tier       TEXT,
    budget_cycle       TEXT,
    executive_sponsor  TEXT,
    champion           TEXT,
    account_goals      TEXT,
    white_space        TEXT,
    created_at         TEXT,
    updated_at         TEXT
);
CREATE TABLE IF NOT EXISTS opportunities (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id           INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    title               TEXT NOT NULL,
    value_sgd           REAL DEFAULT 0,
    stage               TEXT NOT NULL DEFAULT 'Prospect'
                        CHECK(stage IN ('Prospect','Qualified','Proposal','Negotiation','Won','Lost')),
    ai_service_type     TEXT,
    probability         INTEGER DEFAULT 20 CHECK(probability BETWEEN 0 AND 100),
    expected_close_date TEXT,
    contract_type       TEXT,
    next_action         TEXT,
    next_action_date    TEXT,
    decision_maker      TEXT,
    influencers         TEXT,
    engagement_manager  TEXT,
    engagement_partner  TEXT,
    project_start_date  TEXT,
    project_end_date    TEXT,
    pwc_revenue         REAL DEFAULT 0,
    non_pwc_revenue     REAL DEFAULT 0,
    wip                 REAL DEFAULT 0,
    notes               TEXT,
    created_at          TEXT,
    updated_at          TEXT
);
CREATE TABLE IF NOT EXISTS engagements (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    opportunity_id INTEGER REFERENCES opportunities(id) ON DELETE SET NULL,
    client_id      INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    activity_type  TEXT NOT NULL,
    activity_date  TEXT NOT NULL,
    participants   TEXT,
    summary        TEXT,
    outcomes       TEXT,
    next_steps     TEXT,
    created_at     TEXT
);
CREATE TABLE IF NOT EXISTS eminence (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    type         TEXT NOT NULL
                 CHECK(type IN ('Publication','Speaking','Event','Award','Media','Advisory')),
    title        TEXT NOT NULL,
    date         TEXT NOT NULL,
    sector       TEXT,
    platform     TEXT,
    description  TEXT,
    impact_score INTEGER DEFAULT 3 CHECK(impact_score BETWEEN 1 AND 5),
    url          TEXT,
    created_at   TEXT
);
CREATE TABLE IF NOT EXISTS chat_history (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    role       TEXT NOT NULL CHECK(role IN ('user','assistant')),
    content    TEXT NOT NULL,
    created_at TEXT
);
CREATE TABLE IF NOT EXISTS app_context (
    key        TEXT PRIMARY KEY,
    value      TEXT NOT NULL,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS channel_partners (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    name                TEXT NOT NULL,
    partner_type        TEXT,
    program_name        TEXT,
    status              TEXT DEFAULT 'Active',
    tier                TEXT DEFAULT 'Standard',
    engagement_partner  TEXT,
    engagement_manager  TEXT,
    primary_contact     TEXT,
    contact_title       TEXT,
    contact_email       TEXT,
    mou_date            TEXT,
    renewal_date        TEXT,
    focus_sectors       TEXT,
    joint_pipeline_value REAL DEFAULT 0,
    referrals_received  INTEGER DEFAULT 0,
    referrals_converted INTEGER DEFAULT 0,
    notes               TEXT,
    created_at          TEXT,
    updated_at          TEXT
);
CREATE TABLE IF NOT EXISTS opportunity_files (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    opportunity_id  INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
    filename        TEXT NOT NULL,
    file_type       TEXT,
    extracted_text  TEXT,
    file_size_kb    REAL,
    created_at      TEXT
);
CREATE TABLE IF NOT EXISTS proposals (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    opportunity_id  INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
    version         INTEGER DEFAULT 1,
    content         TEXT NOT NULL,
    created_at      TEXT,
    updated_at      TEXT
);
CREATE TABLE IF NOT EXISTS opp_chat_history (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    opportunity_id  INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
    role            TEXT NOT NULL CHECK(role IN ('user','assistant')),
    content         TEXT NOT NULL,
    created_at      TEXT
);
CREATE TABLE IF NOT EXISTS cp_chat_history (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    cp_id           INTEGER NOT NULL REFERENCES channel_partners(id) ON DELETE CASCADE,
    role            TEXT NOT NULL CHECK(role IN ('user','assistant')),
    content         TEXT NOT NULL,
    created_at      TEXT
);
CREATE TABLE IF NOT EXISTS eminence_chat_history (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    role       TEXT NOT NULL CHECK(role IN ('user','assistant')),
    content    TEXT NOT NULL,
    created_at TEXT
);
CREATE TABLE IF NOT EXISTS content_plan (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    archetype   TEXT NOT NULL,
    title       TEXT NOT NULL,
    platform    TEXT,
    target_date TEXT,
    status      TEXT DEFAULT 'Draft',
    notes       TEXT,
    created_at  TEXT
);
"""

_PG_TABLES = [
    """CREATE TABLE IF NOT EXISTS clients (
        id                 SERIAL PRIMARY KEY,
        company            TEXT NOT NULL,
        sector             TEXT NOT NULL,
        sub_sector         TEXT,
        company_size       TEXT,
        buyer_type         TEXT NOT NULL CHECK(buyer_type IN ('Institutional','Owner')),
        country            TEXT DEFAULT 'Singapore',
        key_contact        TEXT,
        contact_title      TEXT,
        relationship_score INTEGER DEFAULT 3 CHECK(relationship_score BETWEEN 1 AND 5),
        ai_maturity        TEXT,
        notes              TEXT,
        annual_revenue     TEXT,
        employee_count     TEXT,
        website            TEXT,
        business_strategy  TEXT,
        business_challenges TEXT,
        competitive_landscape TEXT,
        account_tier       TEXT,
        budget_cycle       TEXT,
        executive_sponsor  TEXT,
        champion           TEXT,
        account_goals      TEXT,
        white_space        TEXT,
        created_at         TEXT,
        updated_at         TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS opportunities (
        id                  SERIAL PRIMARY KEY,
        client_id           INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
        title               TEXT NOT NULL,
        value_sgd           DOUBLE PRECISION DEFAULT 0,
        stage               TEXT NOT NULL DEFAULT 'Prospect'
                            CHECK(stage IN ('Prospect','Qualified','Proposal','Negotiation','Won','Lost')),
        ai_service_type     TEXT,
        probability         INTEGER DEFAULT 20 CHECK(probability BETWEEN 0 AND 100),
        expected_close_date TEXT,
        contract_type       TEXT,
        next_action         TEXT,
        next_action_date    TEXT,
        decision_maker      TEXT,
        influencers         TEXT,
        engagement_manager  TEXT,
        engagement_partner  TEXT,
        project_start_date  TEXT,
        project_end_date    TEXT,
        pwc_revenue         DOUBLE PRECISION DEFAULT 0,
        non_pwc_revenue     DOUBLE PRECISION DEFAULT 0,
        wip                 DOUBLE PRECISION DEFAULT 0,
        notes               TEXT,
        created_at          TEXT,
        updated_at          TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS engagements (
        id             SERIAL PRIMARY KEY,
        opportunity_id INTEGER REFERENCES opportunities(id) ON DELETE SET NULL,
        client_id      INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
        activity_type  TEXT NOT NULL,
        activity_date  TEXT NOT NULL,
        participants   TEXT,
        summary        TEXT,
        outcomes       TEXT,
        next_steps     TEXT,
        created_at     TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS eminence (
        id           SERIAL PRIMARY KEY,
        type         TEXT NOT NULL
                     CHECK(type IN ('Publication','Speaking','Event','Award','Media','Advisory')),
        title        TEXT NOT NULL,
        date         TEXT NOT NULL,
        sector       TEXT,
        platform     TEXT,
        description  TEXT,
        impact_score INTEGER DEFAULT 3 CHECK(impact_score BETWEEN 1 AND 5),
        url          TEXT,
        created_at   TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS chat_history (
        id         SERIAL PRIMARY KEY,
        role       TEXT NOT NULL CHECK(role IN ('user','assistant')),
        content    TEXT NOT NULL,
        created_at TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS app_context (
        key        TEXT PRIMARY KEY,
        value      TEXT NOT NULL,
        updated_at TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS channel_partners (
        id                   SERIAL PRIMARY KEY,
        name                 TEXT NOT NULL,
        partner_type         TEXT,
        program_name         TEXT,
        status               TEXT DEFAULT 'Active',
        tier                 TEXT DEFAULT 'Standard',
        engagement_partner   TEXT,
        engagement_manager   TEXT,
        primary_contact      TEXT,
        contact_title        TEXT,
        contact_email        TEXT,
        mou_date             TEXT,
        renewal_date         TEXT,
        focus_sectors        TEXT,
        joint_pipeline_value DOUBLE PRECISION DEFAULT 0,
        referrals_received   INTEGER DEFAULT 0,
        referrals_converted  INTEGER DEFAULT 0,
        notes                TEXT,
        created_at           TEXT,
        updated_at           TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS opportunity_files (
        id              SERIAL PRIMARY KEY,
        opportunity_id  INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
        filename        TEXT NOT NULL,
        file_type       TEXT,
        extracted_text  TEXT,
        file_size_kb    DOUBLE PRECISION,
        created_at      TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS proposals (
        id              SERIAL PRIMARY KEY,
        opportunity_id  INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
        version         INTEGER DEFAULT 1,
        content         TEXT NOT NULL,
        created_at      TEXT,
        updated_at      TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS opp_chat_history (
        id              SERIAL PRIMARY KEY,
        opportunity_id  INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
        role            TEXT NOT NULL CHECK(role IN ('user','assistant')),
        content         TEXT NOT NULL,
        created_at      TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS cp_chat_history (
        id              SERIAL PRIMARY KEY,
        cp_id           INTEGER NOT NULL REFERENCES channel_partners(id) ON DELETE CASCADE,
        role            TEXT NOT NULL CHECK(role IN ('user','assistant')),
        content         TEXT NOT NULL,
        created_at      TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS eminence_chat_history (
        id         SERIAL PRIMARY KEY,
        role       TEXT NOT NULL CHECK(role IN ('user','assistant')),
        content    TEXT NOT NULL,
        created_at TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS content_plan (
        id          SERIAL PRIMARY KEY,
        archetype   TEXT NOT NULL,
        title       TEXT NOT NULL,
        platform    TEXT,
        target_date TEXT,
        status      TEXT DEFAULT 'Draft',
        notes       TEXT,
        created_at  TEXT
    )""",
]


def init_db():
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                for stmt in _PG_TABLES:
                    cur.execute(stmt)
        else:
            conn.executescript(_SQLITE_SCHEMA)
        _commit(conn)
    finally:
        _close(conn)
    # Migrate: add new columns to existing databases that pre-date them
    _migrate_add_columns("opportunities", [
        "decision_maker TEXT", "influencers TEXT",
        "engagement_manager TEXT", "engagement_partner TEXT",
        "project_start_date TEXT", "project_end_date TEXT",
        "pwc_revenue REAL DEFAULT 0",
        "non_pwc_revenue REAL DEFAULT 0",
        "wip REAL DEFAULT 0",
    ])
    _migrate_add_columns("clients", [
        "annual_revenue TEXT", "employee_count TEXT", "website TEXT",
        "business_strategy TEXT", "business_challenges TEXT", "competitive_landscape TEXT",
        "account_tier TEXT", "budget_cycle TEXT", "executive_sponsor TEXT",
        "champion TEXT", "account_goals TEXT", "white_space TEXT",
    ])
    # Migrate: scale value_sgd from thousands to full S$ for records entered before unit change
    _migrate_value_sgd_to_full()
    # Create tables introduced after initial deployment
    _migrate_create_channel_partners()
    _migrate_create_scope_tables()
    _migrate_create_cp_chat()
    _migrate_create_eminence_tables()
    _migrate_add_columns("channel_partners", [
        "health_score INTEGER DEFAULT 3",
        "last_meeting_date TEXT",
        "next_meeting_date TEXT",
        "meeting_purpose TEXT",
    ])


def _migrate_value_sgd_to_full():
    """Multiply value_sgd by 1000 for records entered as thousands (e.g. 150 → 150000).
    Safe to run repeatedly: only touches values < 10000 (raw S$ amounts, not K entries).
    Errors are swallowed so a transient DB fault cannot prevent _init_db_once caching."""
    conn = None
    try:
        conn = get_conn()
        if _setup():
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE opportunities SET value_sgd = value_sgd * 1000 "
                    "WHERE value_sgd > 0 AND value_sgd < 10000"
                )
        else:
            conn.execute(
                "UPDATE opportunities SET value_sgd = value_sgd * 1000 "
                "WHERE value_sgd > 0 AND value_sgd < 10000"
            )
        _commit(conn)
    except Exception:
        try:
            if conn is not None:
                conn.rollback()
        except Exception:
            pass
    finally:
        if conn is not None:
            _close(conn)


def _migrate_create_channel_partners():
    """Create channel_partners table on existing databases that pre-date it."""
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS channel_partners (
                        id                   SERIAL PRIMARY KEY,
                        name                 TEXT NOT NULL,
                        partner_type         TEXT,
                        program_name         TEXT,
                        status               TEXT DEFAULT 'Active',
                        tier                 TEXT DEFAULT 'Standard',
                        engagement_partner   TEXT,
                        engagement_manager   TEXT,
                        primary_contact      TEXT,
                        contact_title        TEXT,
                        contact_email        TEXT,
                        mou_date             TEXT,
                        renewal_date         TEXT,
                        focus_sectors        TEXT,
                        joint_pipeline_value DOUBLE PRECISION DEFAULT 0,
                        referrals_received   INTEGER DEFAULT 0,
                        referrals_converted  INTEGER DEFAULT 0,
                        notes                TEXT,
                        created_at           TEXT,
                        updated_at           TEXT
                    )
                """)
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS channel_partners (
                    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                    name                 TEXT NOT NULL,
                    partner_type         TEXT,
                    program_name         TEXT,
                    status               TEXT DEFAULT 'Active',
                    tier                 TEXT DEFAULT 'Standard',
                    engagement_partner   TEXT,
                    engagement_manager   TEXT,
                    primary_contact      TEXT,
                    contact_title        TEXT,
                    contact_email        TEXT,
                    mou_date             TEXT,
                    renewal_date         TEXT,
                    focus_sectors        TEXT,
                    joint_pipeline_value REAL DEFAULT 0,
                    referrals_received   INTEGER DEFAULT 0,
                    referrals_converted  INTEGER DEFAULT 0,
                    notes                TEXT,
                    created_at           TEXT,
                    updated_at           TEXT
                )
            """)
        _commit(conn)
    finally:
        _close(conn)


def _migrate_create_scope_tables():
    """Create opportunity_files, proposals, opp_chat_history on existing DBs."""
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS opportunity_files (
                    id SERIAL PRIMARY KEY,
                    opportunity_id INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
                    filename TEXT NOT NULL, file_type TEXT, extracted_text TEXT,
                    file_size_kb DOUBLE PRECISION, created_at TEXT)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS proposals (
                    id SERIAL PRIMARY KEY,
                    opportunity_id INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
                    version INTEGER DEFAULT 1, content TEXT NOT NULL,
                    created_at TEXT, updated_at TEXT)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS opp_chat_history (
                    id SERIAL PRIMARY KEY,
                    opportunity_id INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
                    role TEXT NOT NULL CHECK(role IN ('user','assistant')),
                    content TEXT NOT NULL, created_at TEXT)""")
        else:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS opportunity_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    opportunity_id INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
                    filename TEXT NOT NULL, file_type TEXT, extracted_text TEXT,
                    file_size_kb REAL, created_at TEXT);
                CREATE TABLE IF NOT EXISTS proposals (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    opportunity_id INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
                    version INTEGER DEFAULT 1, content TEXT NOT NULL,
                    created_at TEXT, updated_at TEXT);
                CREATE TABLE IF NOT EXISTS opp_chat_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    opportunity_id INTEGER NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
                    role TEXT NOT NULL CHECK(role IN ('user','assistant')),
                    content TEXT NOT NULL, created_at TEXT);
            """)
        _commit(conn)
    finally:
        _close(conn)


def _migrate_create_cp_chat():
    """Create cp_chat_history table on existing databases that pre-date it."""
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS cp_chat_history (
                    id SERIAL PRIMARY KEY,
                    cp_id INTEGER NOT NULL REFERENCES channel_partners(id) ON DELETE CASCADE,
                    role TEXT NOT NULL CHECK(role IN ('user','assistant')),
                    content TEXT NOT NULL, created_at TEXT)""")
        else:
            conn.execute("""CREATE TABLE IF NOT EXISTS cp_chat_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cp_id INTEGER NOT NULL REFERENCES channel_partners(id) ON DELETE CASCADE,
                role TEXT NOT NULL CHECK(role IN ('user','assistant')),
                content TEXT NOT NULL, created_at TEXT)""")
        _commit(conn)
    finally:
        _close(conn)


def _migrate_create_eminence_tables():
    """Create eminence_chat_history and content_plan tables on existing databases."""
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS eminence_chat_history (
                    id SERIAL PRIMARY KEY,
                    role TEXT NOT NULL CHECK(role IN ('user','assistant')),
                    content TEXT NOT NULL, created_at TEXT)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS content_plan (
                    id SERIAL PRIMARY KEY, archetype TEXT NOT NULL, title TEXT NOT NULL,
                    platform TEXT, target_date TEXT, status TEXT DEFAULT 'Draft',
                    notes TEXT, created_at TEXT)""")
        else:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS eminence_chat_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    role TEXT NOT NULL CHECK(role IN ('user','assistant')),
                    content TEXT NOT NULL, created_at TEXT);
                CREATE TABLE IF NOT EXISTS content_plan (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, archetype TEXT NOT NULL,
                    title TEXT NOT NULL, platform TEXT, target_date TEXT,
                    status TEXT DEFAULT 'Draft', notes TEXT, created_at TEXT);
            """)
        _commit(conn)
    finally:
        _close(conn)


def _migrate_add_columns(table: str, col_defs: list):
    """Add columns that may not exist in older database instances (idempotent)."""
    conn = get_conn()
    try:
        for col_def in col_defs:
            col_name = col_def.split()[0]
            try:
                if _setup():
                    with conn.cursor() as cur:
                        cur.execute(f"ALTER TABLE {table} ADD COLUMN {col_def}")
                else:
                    conn.execute(f"ALTER TABLE {table} ADD COLUMN {col_def}")
                _commit(conn)
            except Exception:
                # PostgreSQL aborts the transaction on error; rollback so the
                # next ALTER TABLE in the loop can proceed cleanly.
                if _setup():
                    conn.rollback()
    finally:
        _close(conn)


# ── Clients ────────────────────────────────────────────────────────────────────

def get_clients(sector=None, buyer_type=None):
    ph = _ph()
    sql = "SELECT * FROM clients WHERE 1=1"
    params: list = []
    if sector and sector != "All":
        sql += f" AND sector = {ph}"
        params.append(sector)
    if buyer_type and buyer_type != "All":
        sql += f" AND buyer_type = {ph}"
        params.append(buyer_type)
    sql += " ORDER BY relationship_score DESC, company"
    conn = get_conn()
    try:
        return _q(conn, sql, tuple(params))
    finally:
        _close(conn)


def get_client(client_id):
    ph = _ph()
    conn = get_conn()
    try:
        return _q1(conn, f"SELECT * FROM clients WHERE id = {ph}", (client_id,))
    finally:
        _close(conn)


def upsert_client(data: dict):
    now = datetime.now().isoformat()
    data = dict(data)  # don't mutate caller's dict
    data.setdefault("sub_sector", "")
    data.setdefault("company_size", "")
    data.setdefault("country", "Singapore")
    data.setdefault("key_contact", "")
    data.setdefault("contact_title", "")
    data.setdefault("relationship_score", 3)
    data.setdefault("ai_maturity", "")
    data.setdefault("notes", "")
    data.setdefault("annual_revenue", "")
    data.setdefault("employee_count", "")
    data.setdefault("website", "")
    data.setdefault("business_strategy", "")
    data.setdefault("business_challenges", "")
    data.setdefault("competitive_landscape", "")
    data.setdefault("account_tier", "")
    data.setdefault("budget_cycle", "")
    data.setdefault("executive_sponsor", "")
    data.setdefault("champion", "")
    data.setdefault("account_goals", "")
    data.setdefault("white_space", "")
    data["updated_at"] = now
    conn = get_conn()
    try:
        if data.get("id"):
            _run(conn, _n("""
                UPDATE clients
                SET company=:company, sector=:sector, sub_sector=:sub_sector,
                    company_size=:company_size, buyer_type=:buyer_type, country=:country,
                    key_contact=:key_contact, contact_title=:contact_title,
                    relationship_score=:relationship_score, ai_maturity=:ai_maturity,
                    notes=:notes,
                    annual_revenue=:annual_revenue, employee_count=:employee_count,
                    website=:website, business_strategy=:business_strategy,
                    business_challenges=:business_challenges,
                    competitive_landscape=:competitive_landscape,
                    account_tier=:account_tier, budget_cycle=:budget_cycle,
                    executive_sponsor=:executive_sponsor, champion=:champion,
                    account_goals=:account_goals, white_space=:white_space,
                    updated_at=:updated_at
                WHERE id=:id
            """), data)
        else:
            data["created_at"] = now
            data["id"] = _run(conn, _n("""
                INSERT INTO clients
                    (company, sector, sub_sector, company_size, buyer_type, country,
                     key_contact, contact_title, relationship_score, ai_maturity, notes,
                     annual_revenue, employee_count, website, business_strategy,
                     business_challenges, competitive_landscape, account_tier,
                     budget_cycle, executive_sponsor, champion, account_goals, white_space,
                     created_at, updated_at)
                VALUES
                    (:company, :sector, :sub_sector, :company_size, :buyer_type, :country,
                     :key_contact, :contact_title, :relationship_score, :ai_maturity, :notes,
                     :annual_revenue, :employee_count, :website, :business_strategy,
                     :business_challenges, :competitive_landscape, :account_tier,
                     :budget_cycle, :executive_sponsor, :champion, :account_goals, :white_space,
                     :created_at, :updated_at)
            """), data)
        _commit(conn)
        return data["id"]
    finally:
        _close(conn)


def delete_client(client_id):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM clients WHERE id = {ph}", (client_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Opportunities ──────────────────────────────────────────────────────────────

def get_opportunities(client_id=None, stage=None):
    ph = _ph()
    sql = """
        SELECT o.*, c.company, c.sector, c.buyer_type
        FROM opportunities o JOIN clients c ON o.client_id = c.id
        WHERE 1=1
    """
    params: list = []
    if client_id:
        sql += f" AND o.client_id = {ph}"
        params.append(client_id)
    if stage and stage != "All":
        sql += f" AND o.stage = {ph}"
        params.append(stage)
    sql += " ORDER BY o.value_sgd DESC"
    conn = get_conn()
    try:
        return _q(conn, sql, tuple(params))
    finally:
        _close(conn)


def get_opportunity(opp_id):
    ph = _ph()
    conn = get_conn()
    try:
        return _q1(conn, f"""
            SELECT o.*, c.company
            FROM opportunities o JOIN clients c ON o.client_id = c.id
            WHERE o.id = {ph}
        """, (opp_id,))
    finally:
        _close(conn)


def upsert_opportunity(data: dict):
    now = datetime.now().isoformat()
    data = dict(data)  # don't mutate caller's dict
    data.setdefault("value_sgd", 0)
    data.setdefault("ai_service_type", "")
    data.setdefault("probability", 50)
    data.setdefault("expected_close_date", "")
    data.setdefault("contract_type", "")
    data.setdefault("next_action", "")
    data.setdefault("next_action_date", "")
    data.setdefault("decision_maker", "")
    data.setdefault("influencers", "")
    data.setdefault("engagement_manager", "")
    data.setdefault("engagement_partner", "")
    data.setdefault("project_start_date", "")
    data.setdefault("project_end_date", "")
    data.setdefault("pwc_revenue", 0)
    data.setdefault("non_pwc_revenue", 0)
    data.setdefault("wip", 0)
    data.setdefault("notes", "")
    data["updated_at"] = now
    conn = get_conn()
    try:
        if data.get("id"):
            _run(conn, _n("""
                UPDATE opportunities
                SET client_id=:client_id, title=:title, value_sgd=:value_sgd,
                    stage=:stage, ai_service_type=:ai_service_type, probability=:probability,
                    expected_close_date=:expected_close_date, contract_type=:contract_type,
                    next_action=:next_action, next_action_date=:next_action_date,
                    decision_maker=:decision_maker, influencers=:influencers,
                    engagement_manager=:engagement_manager, engagement_partner=:engagement_partner,
                    project_start_date=:project_start_date, project_end_date=:project_end_date,
                    pwc_revenue=:pwc_revenue, non_pwc_revenue=:non_pwc_revenue, wip=:wip,
                    notes=:notes, updated_at=:updated_at
                WHERE id=:id
            """), data)
        else:
            data["created_at"] = now
            data["id"] = _run(conn, _n("""
                INSERT INTO opportunities
                    (client_id, title, value_sgd, stage, ai_service_type, probability,
                     expected_close_date, contract_type, next_action, next_action_date,
                     decision_maker, influencers, engagement_manager, engagement_partner,
                     project_start_date, project_end_date,
                     pwc_revenue, non_pwc_revenue, wip,
                     notes, created_at, updated_at)
                VALUES
                    (:client_id, :title, :value_sgd, :stage, :ai_service_type, :probability,
                     :expected_close_date, :contract_type, :next_action, :next_action_date,
                     :decision_maker, :influencers, :engagement_manager, :engagement_partner,
                     :project_start_date, :project_end_date,
                     :pwc_revenue, :non_pwc_revenue, :wip,
                     :notes, :created_at, :updated_at)
            """), data)
        _commit(conn)
        return data["id"]
    finally:
        _close(conn)


def delete_opportunity(opp_id):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM opportunities WHERE id = {ph}", (opp_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Engagements ────────────────────────────────────────────────────────────────

def get_engagements(client_id=None, opportunity_id=None, limit=50):
    ph = _ph()
    sql = """
        SELECT e.*, c.company, o.title AS opp_title
        FROM engagements e
        JOIN clients c ON e.client_id = c.id
        LEFT JOIN opportunities o ON e.opportunity_id = o.id
        WHERE 1=1
    """
    params: list = []
    if client_id:
        sql += f" AND e.client_id = {ph}"
        params.append(client_id)
    if opportunity_id:
        sql += f" AND e.opportunity_id = {ph}"
        params.append(opportunity_id)
    sql += f" ORDER BY e.activity_date DESC LIMIT {ph}"
    params.append(limit)
    conn = get_conn()
    try:
        return _q(conn, sql, tuple(params))
    finally:
        _close(conn)


def add_engagement(data: dict):
    data = dict(data)  # don't mutate caller's dict
    data.setdefault("opportunity_id", None)
    data.setdefault("participants", "")
    data.setdefault("summary", "")
    data.setdefault("outcomes", "")
    data.setdefault("next_steps", "")
    data.setdefault("created_at", datetime.now().isoformat())
    conn = get_conn()
    try:
        eid = _run(conn, _n("""
            INSERT INTO engagements
                (opportunity_id, client_id, activity_type, activity_date,
                 participants, summary, outcomes, next_steps, created_at)
            VALUES
                (:opportunity_id, :client_id, :activity_type, :activity_date,
                 :participants, :summary, :outcomes, :next_steps, :created_at)
        """), data)
        _commit(conn)
        return eid
    finally:
        _close(conn)


# ── Eminence ───────────────────────────────────────────────────────────────────

def get_eminence(type_filter=None):
    ph = _ph()
    sql = "SELECT * FROM eminence WHERE 1=1"
    params: list = []
    if type_filter and type_filter != "All":
        sql += f" AND type = {ph}"
        params.append(type_filter)
    sql += " ORDER BY date DESC"
    conn = get_conn()
    try:
        return _q(conn, sql, tuple(params))
    finally:
        _close(conn)


def add_eminence(data: dict):
    data = dict(data)  # don't mutate caller's dict
    data.setdefault("sector", "")
    data.setdefault("platform", "")
    data.setdefault("description", "")
    data.setdefault("impact_score", 3)
    data.setdefault("url", "")
    data.setdefault("created_at", datetime.now().isoformat())
    conn = get_conn()
    try:
        eid = _run(conn, _n("""
            INSERT INTO eminence
                (type, title, date, sector, platform, description, impact_score, url, created_at)
            VALUES
                (:type, :title, :date, :sector, :platform, :description, :impact_score, :url, :created_at)
        """), data)
        _commit(conn)
        return eid
    finally:
        _close(conn)


def delete_eminence(em_id):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM eminence WHERE id = {ph}", (em_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Eminence Chat ─────────────────────────────────────────────────────────────

def save_eminence_message(role: str, content: str):
    ph = _ph()
    now = datetime.now().isoformat()
    conn = get_conn()
    try:
        _run(conn, f"INSERT INTO eminence_chat_history (role, content, created_at) VALUES ({ph},{ph},{ph})",
             (role, content, now))
        _commit(conn)
    finally:
        _close(conn)


def get_eminence_chat_history(limit=50):
    ph = _ph()
    conn = get_conn()
    try:
        rows = _q(conn, f"SELECT role, content FROM eminence_chat_history ORDER BY id DESC LIMIT {ph}", (limit,))
        return list(reversed(rows))
    finally:
        _close(conn)


def clear_eminence_chat_history():
    conn = get_conn()
    try:
        _run(conn, "DELETE FROM eminence_chat_history")
        _commit(conn)
    finally:
        _close(conn)


# ── Content Plan ──────────────────────────────────────────────────────────────

def add_content_plan(data: dict):
    data = dict(data)
    data.setdefault("platform", "")
    data.setdefault("target_date", "")
    data.setdefault("status", "Draft")
    data.setdefault("notes", "")
    data.setdefault("created_at", datetime.now().isoformat())
    conn = get_conn()
    try:
        eid = _run(conn, _n("""
            INSERT INTO content_plan (archetype, title, platform, target_date, status, notes, created_at)
            VALUES (:archetype, :title, :platform, :target_date, :status, :notes, :created_at)
        """), data)
        _commit(conn)
        return eid
    finally:
        _close(conn)


def get_content_plan(status_filter=None):
    ph = _ph()
    sql = "SELECT * FROM content_plan WHERE 1=1"
    params: list = []
    if status_filter and status_filter != "All":
        sql += f" AND status = {ph}"
        params.append(status_filter)
    sql += " ORDER BY target_date, id"
    conn = get_conn()
    try:
        return _q(conn, sql, tuple(params))
    finally:
        _close(conn)


def update_content_plan_status(item_id: int, status: str):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"UPDATE content_plan SET status = {ph} WHERE id = {ph}", (status, item_id))
        _commit(conn)
    finally:
        _close(conn)


def delete_content_plan(item_id: int):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM content_plan WHERE id = {ph}", (item_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Channel Partners ───────────────────────────────────────────────────────────

def get_channel_partners(status=None):
    ph = _ph()
    sql = "SELECT * FROM channel_partners WHERE 1=1"
    params: list = []
    if status and status != "All":
        sql += f" AND status = {ph}"
        params.append(status)
    sql += " ORDER BY tier, name"
    conn = get_conn()
    try:
        return _q(conn, sql, tuple(params))
    finally:
        _close(conn)


def get_channel_partner(cp_id):
    ph = _ph()
    conn = get_conn()
    try:
        return _q1(conn, f"SELECT * FROM channel_partners WHERE id = {ph}", (cp_id,))
    finally:
        _close(conn)


def upsert_channel_partner(data: dict):
    now = datetime.now().isoformat()
    data = dict(data)
    data.setdefault("partner_type", "")
    data.setdefault("program_name", "")
    data.setdefault("status", "Active")
    data.setdefault("tier", "Standard")
    data.setdefault("engagement_partner", "")
    data.setdefault("engagement_manager", "")
    data.setdefault("primary_contact", "")
    data.setdefault("contact_title", "")
    data.setdefault("contact_email", "")
    data.setdefault("mou_date", "")
    data.setdefault("renewal_date", "")
    data.setdefault("focus_sectors", "")
    data.setdefault("joint_pipeline_value", 0)
    data.setdefault("referrals_received", 0)
    data.setdefault("referrals_converted", 0)
    data.setdefault("notes", "")
    data.setdefault("health_score", 3)
    data.setdefault("last_meeting_date", "")
    data.setdefault("next_meeting_date", "")
    data.setdefault("meeting_purpose", "")
    data["updated_at"] = now
    conn = get_conn()
    try:
        if data.get("id"):
            _run(conn, _n("""
                UPDATE channel_partners
                SET name=:name, partner_type=:partner_type, program_name=:program_name,
                    status=:status, tier=:tier,
                    engagement_partner=:engagement_partner, engagement_manager=:engagement_manager,
                    primary_contact=:primary_contact, contact_title=:contact_title,
                    contact_email=:contact_email,
                    mou_date=:mou_date, renewal_date=:renewal_date,
                    focus_sectors=:focus_sectors,
                    joint_pipeline_value=:joint_pipeline_value,
                    referrals_received=:referrals_received,
                    referrals_converted=:referrals_converted,
                    notes=:notes,
                    health_score=:health_score, last_meeting_date=:last_meeting_date,
                    next_meeting_date=:next_meeting_date, meeting_purpose=:meeting_purpose,
                    updated_at=:updated_at
                WHERE id=:id
            """), data)
        else:
            data["created_at"] = now
            data["id"] = _run(conn, _n("""
                INSERT INTO channel_partners
                    (name, partner_type, program_name, status, tier,
                     engagement_partner, engagement_manager,
                     primary_contact, contact_title, contact_email,
                     mou_date, renewal_date, focus_sectors,
                     joint_pipeline_value, referrals_received, referrals_converted,
                     notes, health_score, last_meeting_date, next_meeting_date, meeting_purpose,
                     created_at, updated_at)
                VALUES
                    (:name, :partner_type, :program_name, :status, :tier,
                     :engagement_partner, :engagement_manager,
                     :primary_contact, :contact_title, :contact_email,
                     :mou_date, :renewal_date, :focus_sectors,
                     :joint_pipeline_value, :referrals_received, :referrals_converted,
                     :notes, :health_score, :last_meeting_date, :next_meeting_date, :meeting_purpose,
                     :created_at, :updated_at)
            """), data)
        _commit(conn)
        return data["id"]
    finally:
        _close(conn)


def delete_channel_partner(cp_id):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM channel_partners WHERE id = {ph}", (cp_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Opportunity Files ──────────────────────────────────────────────────────────

def add_opportunity_file(data: dict):
    data = dict(data)
    data.setdefault("file_type", "")
    data.setdefault("extracted_text", "")
    data.setdefault("file_size_kb", 0)
    data["created_at"] = datetime.now().isoformat()
    conn = get_conn()
    try:
        fid = _run(conn, _n("""
            INSERT INTO opportunity_files
                (opportunity_id, filename, file_type, extracted_text, file_size_kb, created_at)
            VALUES (:opportunity_id, :filename, :file_type, :extracted_text, :file_size_kb, :created_at)
        """), data)
        _commit(conn)
        return fid
    finally:
        _close(conn)


def get_opportunity_files(opportunity_id):
    ph = _ph()
    conn = get_conn()
    try:
        return _q(conn,
                  f"SELECT * FROM opportunity_files WHERE opportunity_id = {ph} ORDER BY created_at DESC",
                  (opportunity_id,))
    finally:
        _close(conn)


def delete_opportunity_file(file_id):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM opportunity_files WHERE id = {ph}", (file_id,))
        _commit(conn)
    finally:
        _close(conn)


def deduplicate_opportunity_files(opportunity_id: int) -> int:
    """Keep only the most recent file per filename; return number of rows removed."""
    files = get_opportunity_files(opportunity_id)
    seen: dict = {}
    to_delete = []
    for f in sorted(files, key=lambda x: x["created_at"] or "", reverse=True):
        name = f["filename"]
        if name in seen:
            to_delete.append(f["id"])
        else:
            seen[name] = f["id"]
    for fid in to_delete:
        delete_opportunity_file(fid)
    return len(to_delete)


# ── Proposals ──────────────────────────────────────────────────────────────────

def upsert_proposal(opportunity_id: int, content: str, version: int = None):
    now = datetime.now().isoformat()
    ph = _ph()
    conn = get_conn()
    try:
        existing = _q1(conn,
                       f"SELECT * FROM proposals WHERE opportunity_id = {ph} ORDER BY version DESC LIMIT 1",
                       (opportunity_id,))
        if existing:
            _run(conn, _n("""
                UPDATE proposals SET content=:content, version=:version, updated_at=:updated_at
                WHERE id=:id
            """), {"id": existing["id"], "content": content,
                   "version": version or existing["version"],
                   "updated_at": now})
        else:
            _run(conn, _n("""
                INSERT INTO proposals (opportunity_id, version, content, created_at, updated_at)
                VALUES (:opportunity_id, :version, :content, :created_at, :updated_at)
            """), {"opportunity_id": opportunity_id, "version": version or 1,
                   "content": content, "created_at": now, "updated_at": now})
        _commit(conn)
    finally:
        _close(conn)


def get_proposal(opportunity_id: int):
    ph = _ph()
    conn = get_conn()
    try:
        return _q1(conn,
                   f"SELECT * FROM proposals WHERE opportunity_id = {ph} ORDER BY version DESC LIMIT 1",
                   (opportunity_id,))
    finally:
        _close(conn)


def delete_proposal(opportunity_id: int):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM proposals WHERE opportunity_id = {ph}", (opportunity_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Per-opportunity chat ────────────────────────────────────────────────────────

def save_opp_message(opportunity_id: int, role: str, content: str):
    ph = _ph()
    now = datetime.now().isoformat()
    conn = get_conn()
    try:
        _run(conn,
             f"INSERT INTO opp_chat_history (opportunity_id, role, content, created_at) VALUES ({ph},{ph},{ph},{ph})",
             (opportunity_id, role, content, now))
        _commit(conn)
    finally:
        _close(conn)


def get_opp_chat(opportunity_id: int, limit: int = 40):
    ph = _ph()
    conn = get_conn()
    try:
        rows = _q(conn,
                  f"SELECT role, content FROM opp_chat_history WHERE opportunity_id = {ph} ORDER BY id DESC LIMIT {ph}",
                  (opportunity_id, limit))
        return list(reversed(rows))
    finally:
        _close(conn)


def clear_opp_chat(opportunity_id: int):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM opp_chat_history WHERE opportunity_id = {ph}", (opportunity_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Per-partner chat ────────────────────────────────────────────────────────────

def save_cp_message(cp_id: int, role: str, content: str):
    ph = _ph()
    now = datetime.now().isoformat()
    conn = get_conn()
    try:
        _run(conn,
             f"INSERT INTO cp_chat_history (cp_id, role, content, created_at) VALUES ({ph},{ph},{ph},{ph})",
             (cp_id, role, content, now))
        _commit(conn)
    finally:
        _close(conn)


def get_cp_chat(cp_id: int, limit: int = 40):
    ph = _ph()
    conn = get_conn()
    try:
        rows = _q(conn,
                  f"SELECT role, content FROM cp_chat_history WHERE cp_id = {ph} ORDER BY id DESC LIMIT {ph}",
                  (cp_id, limit))
        return list(reversed(rows))
    finally:
        _close(conn)


def clear_cp_chat(cp_id: int):
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM cp_chat_history WHERE cp_id = {ph}", (cp_id,))
        _commit(conn)
    finally:
        _close(conn)


# ── Chat history ───────────────────────────────────────────────────────────────

def save_message(role: str, content: str):
    ph = _ph()
    now = datetime.now().isoformat()
    conn = get_conn()
    try:
        _run(conn, f"INSERT INTO chat_history (role, content, created_at) VALUES ({ph}, {ph}, {ph})",
             (role, content, now))
        _commit(conn)
    finally:
        _close(conn)


def get_chat_history(limit=50):
    ph = _ph()
    conn = get_conn()
    try:
        rows = _q(conn, f"SELECT role, content FROM chat_history ORDER BY id DESC LIMIT {ph}", (limit,))
        return list(reversed(rows))
    finally:
        _close(conn)


def clear_chat_history():
    conn = get_conn()
    try:
        _run(conn, "DELETE FROM chat_history")
        _commit(conn)
    finally:
        _close(conn)


# ── Dashboard stats ────────────────────────────────────────────────────────────

def get_pipeline_summary():
    conn = get_conn()
    try:
        return _q(conn, """
            SELECT stage,
                   COUNT(*) AS count,
                   SUM(value_sgd) AS total_value,
                   SUM(value_sgd * probability / 100.0) AS weighted_value
            FROM opportunities
            GROUP BY stage
        """)
    finally:
        _close(conn)


def get_upcoming_actions(days=14):
    ph = _ph()
    target_date = (date.today() + timedelta(days=days)).isoformat()
    sql = f"""
        SELECT o.id, o.title, o.next_action, o.next_action_date,
               o.stage, o.value_sgd, c.company, c.buyer_type
        FROM opportunities o JOIN clients c ON o.client_id = c.id
        WHERE o.next_action_date IS NOT NULL
          AND o.next_action_date <= {ph}
          AND o.stage NOT IN ('Won', 'Lost')
        ORDER BY o.next_action_date
    """
    conn = get_conn()
    try:
        return _q(conn, sql, (target_date,))
    finally:
        _close(conn)


def get_eminence_summary():
    conn = get_conn()
    try:
        return _q(conn, """
            SELECT type, COUNT(*) AS count, AVG(impact_score) AS avg_impact
            FROM eminence
            GROUP BY type
        """)
    finally:
        _close(conn)


def get_client_sectors():
    conn = get_conn()
    try:
        rows = _q(conn, "SELECT DISTINCT sector FROM clients ORDER BY sector")
        return [r["sector"] for r in rows]
    finally:
        _close(conn)


# ── Demo data cleanup ─────────────────────────────────────────────────────────

_DEMO_COMPANIES = [
    "DBS Bank", "GovTech Singapore", "Grab Holdings", "Pacific International Lines",
    "Wilmar International", "Singtel", "Raffles Medical Group", "MAS (Monetary Authority)",
    "CapitaLand Investment", "Dyson Singapore",
]

_DEMO_EMINENCE_TITLES = [
    "AI Governance in Singapore Financial Services: A PwC Perspective",
    "Keynote: Responsible AI for Asia's Financial Sector",
    "Panel: GenAI Implementation Challenges in Government",
    "Interview: The Business Times — 'AI Advisors Are The New Management Consultants'",
    "MAS FEAT Principles Working Group — External Advisor",
    "PwC AI Day Singapore 2026",
    "Singapore AI Readiness Index 2026",
    "Singapore Business Review — Top 40 Under 40 in Tech",
]


def delete_demo_data():
    """Remove demo-seeded companies (cascades to their deals and activities) and demo eminence items."""
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                cur.execute("DELETE FROM clients WHERE company = ANY(%s)", (_DEMO_COMPANIES,))
                cur.execute("DELETE FROM eminence WHERE title = ANY(%s)", (_DEMO_EMINENCE_TITLES,))
        else:
            ph_c = ",".join(["?" for _ in _DEMO_COMPANIES])
            ph_e = ",".join(["?" for _ in _DEMO_EMINENCE_TITLES])
            conn.execute(f"DELETE FROM clients WHERE company IN ({ph_c})", _DEMO_COMPANIES)
            conn.execute(f"DELETE FROM eminence WHERE title IN ({ph_e})", _DEMO_EMINENCE_TITLES)
        _commit(conn)
    finally:
        _close(conn)


# ── Persistent app context (survives reboots) ──────────────────────────────────

def save_pending_doc(doc: dict):
    """Upsert the pending document context to the DB so it survives reboots."""
    now = datetime.now().isoformat()
    value = json.dumps(doc)
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                cur.execute(
                    "INSERT INTO app_context (key, value, updated_at) VALUES (%s, %s, %s) "
                    "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = EXCLUDED.updated_at",
                    ("pending_doc", value, now),
                )
        else:
            conn.execute(
                "INSERT OR REPLACE INTO app_context (key, value, updated_at) VALUES (?, ?, ?)",
                ("pending_doc", value, now),
            )
        _commit(conn)
    finally:
        _close(conn)


def get_pending_doc():
    """Return the pending document context dict from the DB, or None."""
    ph = _ph()
    conn = get_conn()
    try:
        row = _q1(conn, f"SELECT value FROM app_context WHERE key = {ph}", ("pending_doc",))
        return json.loads(row["value"]) if row else None
    finally:
        _close(conn)


def clear_pending_doc():
    """Delete the pending document context from the DB."""
    ph = _ph()
    conn = get_conn()
    try:
        _run(conn, f"DELETE FROM app_context WHERE key = {ph}", ("pending_doc",))
        _commit(conn)
    finally:
        _close(conn)


def set_context(key: str, value: str):
    """Upsert a string value into app_context by key."""
    now = datetime.now().isoformat()
    conn = get_conn()
    try:
        if _setup():
            with conn.cursor() as cur:
                cur.execute(
                    "INSERT INTO app_context (key, value, updated_at) VALUES (%s, %s, %s) "
                    "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = EXCLUDED.updated_at",
                    (key, value, now),
                )
        else:
            conn.execute(
                "INSERT OR REPLACE INTO app_context (key, value, updated_at) VALUES (?, ?, ?)",
                (key, value, now),
            )
        _commit(conn)
    finally:
        _close(conn)


def get_context(key: str) -> str | None:
    """Retrieve a string value from app_context by key, or None if missing."""
    ph = _ph()
    conn = get_conn()
    try:
        row = _q1(conn, f"SELECT value FROM app_context WHERE key = {ph}", (key,))
        return row["value"] if row else None
    finally:
        _close(conn)
