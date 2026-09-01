"""
Database access for API clients (who's allowed to call this service) and jobs
(one row per generation run, used for usage limits and audit history).

Uses a real asyncpg connection pool, created once at app startup (see
main.py's lifespan) and closed at shutdown - same pattern Vett uses, so the
whole product suite has one consistent operational story rather than each
product inventing its own. The pool handles reconnecting after Neon's
autosuspend; a failed query on a stale connection surfaces as a normal
retryable error rather than something this module needs to special-case.
"""
import os
import hashlib
import secrets
import logging
from contextlib import asynccontextmanager

import asyncpg

logger = logging.getLogger(__name__)

_pool: asyncpg.Pool | None = None


def _dsn() -> str:
    dsn = os.getenv('DATABASE_URL')
    if not dsn:
        raise RuntimeError(
            "DATABASE_URL is not set. Copy your Neon connection string "
            "(Neon dashboard -> Connection Details) into the DATABASE_URL env var."
        )
    return dsn


async def init_pool() -> asyncpg.Pool:
    """Call once, from the FastAPI lifespan startup. Idempotent - safe to
    call again (returns the existing pool) if something calls it twice."""
    global _pool
    if _pool is None:
        _pool = await asyncpg.create_pool(_dsn(), min_size=2, max_size=10, command_timeout=15)
        logger.info("Database pool created.")
    return _pool


async def close_pool() -> None:
    global _pool
    if _pool is not None:
        await _pool.close()
        _pool = None
        logger.info("Database pool closed.")


def get_pool() -> asyncpg.Pool:
    if _pool is None:
        raise RuntimeError("Database pool not initialized - init_pool() must run at app startup first.")
    return _pool


@asynccontextmanager
async def get_connection(client_id: int | None = None):
    """When client_id is given, the connection is pinned to that tenant for
    its whole transaction (via a session-local Postgres setting), and Row-
    Level Security on statement_jobs enforces it - so even a future query
    against that table that forgets its own WHERE client_id = ... clause
    still can't read or write another tenant's rows. This is enforced by
    Postgres itself, not by this function remembering to filter correctly.
    """
    pool = get_pool()
    async with pool.acquire() as conn:
        async with conn.transaction():
            if client_id is not None:
                # set_config (not a raw `SET ... = $1`) is what accepts a
                # normal parameterized value here, so this can never become
                # a SQL-injection vector even though client_id ultimately
                # traces back to a request. Must run *inside* the
                # transaction (is_local=true scopes to "rest of the current
                # transaction" - it has no effect if set before one starts).
                await conn.execute("SELECT set_config('app.current_client_id', $1, true)", str(client_id))
            yield conn


async def _ensure_jobs_rls(conn: asyncpg.Connection) -> None:
    """Row-Level Security on statement_jobs, scoped to the session's
    app.current_client_id. FORCE makes it apply even to the role that owns
    the table (Postgres exempts table owners from RLS by default, and the
    role in your Neon connection string is typically the owner since it's
    the one that ran CREATE TABLE) - without FORCE, this app's own normal
    connection could silently bypass the policy it just created.
    """
    await conn.execute("ALTER TABLE statement_jobs ENABLE ROW LEVEL SECURITY;")
    await conn.execute("ALTER TABLE statement_jobs FORCE ROW LEVEL SECURITY;")
    await conn.execute("""
        DO $$
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM pg_policies
                WHERE tablename = 'statement_jobs' AND policyname = 'tenant_isolation'
            ) THEN
                CREATE POLICY tenant_isolation ON statement_jobs
                    USING (client_id = current_setting('app.current_client_id', true)::integer)
                    WITH CHECK (client_id = current_setting('app.current_client_id', true)::integer);
            END IF;
        END
        $$;
    """)


async def init_schema() -> None:
    """Idempotent - safe to call on every app startup."""
    async with get_connection() as conn:
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS statement_api_clients (
                id SERIAL PRIMARY KEY,
                label TEXT UNIQUE NOT NULL,
                key_hash TEXT UNIQUE NOT NULL,
                plan TEXT NOT NULL DEFAULT 'free',
                monthly_job_limit INTEGER,     -- NULL = unlimited
                max_rows_per_job INTEGER,       -- NULL = unlimited
                password_protection_allowed BOOLEAN NOT NULL DEFAULT FALSE,
                is_active BOOLEAN NOT NULL DEFAULT TRUE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT now()
            );
        """)
        await conn.execute("ALTER TABLE statement_api_clients ADD COLUMN IF NOT EXISTS max_rows_per_job INTEGER;")
        await conn.execute("ALTER TABLE statement_api_clients ADD COLUMN IF NOT EXISTS password_protection_allowed BOOLEAN NOT NULL DEFAULT FALSE;")
        await conn.execute("""
            CREATE TABLE IF NOT EXISTS statement_jobs (
                temp_id UUID PRIMARY KEY,
                client_id INTEGER NOT NULL REFERENCES statement_api_clients(id),
                status TEXT NOT NULL DEFAULT 'processing',
                row_count INTEGER,
                created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                completed_at TIMESTAMPTZ
            );
        """)
        await conn.execute("CREATE INDEX IF NOT EXISTS idx_jobs_client_created ON statement_jobs(client_id, created_at);")
        await _ensure_jobs_rls(conn)
    logger.info("Database schema is up to date.")


def hash_key(raw_key: str) -> str:
    return hashlib.sha256(raw_key.encode('utf-8')).hexdigest()


async def get_client_by_key(raw_key: str) -> dict | None:
    """Returns a dict {id, label, plan, monthly_job_limit, max_rows_per_job,
    password_protection_allowed} or None.

    Not RLS-scoped: this is the auth bootstrap itself, run before we know
    who's calling. It only ever matches the one row whose key_hash equals
    the hash of the exact key supplied, so it can't be used to read or
    enumerate any other client's data.
    """
    async with get_connection() as conn:
        row = await conn.fetchrow(
            "SELECT id, label, plan, monthly_job_limit, max_rows_per_job, password_protection_allowed "
            "FROM statement_api_clients WHERE key_hash = $1 AND is_active = TRUE",
            hash_key(raw_key),
        )
        return dict(row) if row else None


async def count_jobs_this_month(client_id: int) -> int:
    async with get_connection(client_id=client_id) as conn:
        return await conn.fetchval(
            "SELECT COUNT(*) FROM statement_jobs "
            "WHERE client_id = $1 AND created_at >= date_trunc('month', now())",
            client_id,
        )


async def create_job(temp_id: str, client_id: int, row_count: int | None = None) -> None:
    async with get_connection(client_id=client_id) as conn:
        await conn.execute(
            "INSERT INTO statement_jobs (temp_id, client_id, status, row_count) VALUES ($1, $2, 'processing', $3)",
            temp_id, client_id, row_count,
        )


async def update_job_status(temp_id: str, status: str, client_id: int) -> None:
    async with get_connection(client_id=client_id) as conn:
        if status in ('completed', 'error'):
            await conn.execute(
                "UPDATE statement_jobs SET status = $1, completed_at = now() WHERE temp_id = $2",
                status, temp_id,
            )
        else:
            await conn.execute("UPDATE statement_jobs SET status = $1 WHERE temp_id = $2", status, temp_id)


async def create_api_client(label: str, plan: str = 'free', monthly_job_limit: int | None = None,
                             max_rows_per_job: int | None = None,
                             password_protection_allowed: bool = False) -> str:
    """Creates a new client and returns the RAW api key.

    The raw key is only ever returned here, at creation time - only its hash
    is stored. If it's lost, the only fix is issuing a new key.
    """
    raw_key = f"sg_live_{secrets.token_urlsafe(32)}"
    async with get_connection() as conn:
        await conn.execute(
            "INSERT INTO statement_api_clients "
            "(label, key_hash, plan, monthly_job_limit, max_rows_per_job, password_protection_allowed) "
            "VALUES ($1, $2, $3, $4, $5, $6)",
            label, hash_key(raw_key), plan, monthly_job_limit, max_rows_per_job, password_protection_allowed,
        )
    return raw_key
