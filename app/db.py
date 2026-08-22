"""
Database access for API clients (who's allowed to call this service) and jobs
(one row per generation run, used for usage limits and audit history).

Uses Neon Postgres (or any standard Postgres) via DATABASE_URL. Connections
are opened fresh per call rather than pooled - Neon's free/pro tiers
autosuspend the underlying compute after a period of inactivity, so a
long-lived pool would end up holding dead connections that fail on first use
after an idle period. Short-lived connections reconnect cleanly every time,
which costs a bit of latency per request but is much simpler to get right at
this scale. If usage grows enough that connection setup time becomes a
bottleneck, swap this for a proper pool (e.g. psycopg2.pool + a keep-alive
ping) then - not before.
"""
import os
import hashlib
import secrets
import logging
from contextlib import contextmanager

import psycopg2
import psycopg2.extras

logger = logging.getLogger(__name__)


def _dsn():
    dsn = os.getenv('DATABASE_URL')
    if not dsn:
        raise RuntimeError(
            "DATABASE_URL is not set. Copy your Neon connection string "
            "(Neon dashboard -> Connection Details) into the DATABASE_URL env var."
        )
    return dsn


@contextmanager
def get_connection():
    conn = psycopg2.connect(_dsn())
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_schema():
    """Idempotent - safe to call on every app startup."""
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS statement_api_clients (
                    id SERIAL PRIMARY KEY,
                    label TEXT UNIQUE NOT NULL,
                    key_hash TEXT UNIQUE NOT NULL,
                    plan TEXT NOT NULL DEFAULT 'free',
                    monthly_job_limit INTEGER,     -- NULL = unlimited
                    is_active BOOLEAN NOT NULL DEFAULT TRUE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                );
            """)
            cur.execute("""
                CREATE TABLE IF NOT EXISTS statement_jobs (
                    temp_id UUID PRIMARY KEY,
                    client_id INTEGER NOT NULL REFERENCES statement_api_clients(id),
                    status TEXT NOT NULL DEFAULT 'processing',
                    row_count INTEGER,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    completed_at TIMESTAMPTZ
                );
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_jobs_client_created ON statement_jobs(client_id, created_at);")
    logger.info("Database schema is up to date.")


def hash_key(raw_key):
    return hashlib.sha256(raw_key.encode('utf-8')).hexdigest()


def get_client_by_key(raw_key):
    """Returns a dict {id, label, plan, monthly_job_limit} or None."""
    with get_connection() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                "SELECT id, label, plan, monthly_job_limit FROM statement_api_clients "
                "WHERE key_hash = %s AND is_active = TRUE",
                (hash_key(raw_key),)
            )
            return cur.fetchone()


def count_jobs_this_month(client_id):
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT COUNT(*) FROM statement_jobs "
                "WHERE client_id = %s AND created_at >= date_trunc('month', now())",
                (client_id,)
            )
            return cur.fetchone()[0]


def create_job(temp_id, client_id):
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO statement_jobs (temp_id, client_id, status) VALUES (%s, %s, 'processing')",
                (temp_id, client_id)
            )


def update_job_status(temp_id, status):
    with get_connection() as conn:
        with conn.cursor() as cur:
            if status in ('completed', 'error'):
                cur.execute(
                    "UPDATE statement_jobs SET status = %s, completed_at = now() WHERE temp_id = %s",
                    (status, temp_id)
                )
            else:
                cur.execute("UPDATE statement_jobs SET status = %s WHERE temp_id = %s", (status, temp_id))


def create_api_client(label, plan='free', monthly_job_limit=None):
    """Creates a new client and returns the RAW api key.

    The raw key is only ever returned here, at creation time - only its hash
    is stored. If it's lost, the only fix is issuing a new key.
    """
    raw_key = f"sg_live_{secrets.token_urlsafe(32)}"
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO statement_api_clients (label, key_hash, plan, monthly_job_limit) "
                "VALUES (%s, %s, %s, %s)",
                (label, hash_key(raw_key), plan, monthly_job_limit)
            )
    return raw_key
