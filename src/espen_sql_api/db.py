"""Database engines and session factories."""

import os
from pathlib import Path

from dotenv import load_dotenv
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

load_dotenv()

BASE_PATH = Path(__file__).resolve(strict=True).parent

# Connection pool configuration
POOL_SIZE = int(os.getenv("DB_POOL_SIZE", "5"))
POOL_MAX_OVERFLOW = int(os.getenv("DB_POOL_MAX_OVERFLOW", "10"))
POOL_RECYCLE = int(os.getenv("DB_POOL_RECYCLE", "3600"))
POOL_PRE_PING = os.getenv("DB_POOL_PRE_PING", "true").lower() == "true"
POOL_TIMEOUT = int(os.getenv("DB_POOL_TIMEOUT", "30"))
QUERY_TIMEOUT = int(os.getenv("DB_QUERY_TIMEOUT", "30"))
ECHO_POOL = os.getenv("DB_ECHO_POOL", "false")

# PostgreSQL (oncho projection data)
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=POOL_PRE_PING,
    pool_recycle=POOL_RECYCLE,
    pool_size=POOL_SIZE,
    max_overflow=POOL_MAX_OVERFLOW,
    pool_timeout=POOL_TIMEOUT,
    connect_args={"connect_timeout": QUERY_TIMEOUT},
    echo_pool=ECHO_POOL if ECHO_POOL == "debug" else False,
)
SessionLocal = sessionmaker(bind=engine)

# MSSQL (ESPEN analytical tables)
REMOTE_DATABASE_URL = os.getenv("REMOTE_DATABASE_URL")
remote_engine = create_engine(
    REMOTE_DATABASE_URL,
    pool_pre_ping=POOL_PRE_PING,
    pool_recycle=POOL_RECYCLE,
    pool_size=POOL_SIZE,
    max_overflow=POOL_MAX_OVERFLOW,
    pool_timeout=POOL_TIMEOUT,
    connect_args={"timeout": QUERY_TIMEOUT},
    echo_pool=ECHO_POOL if ECHO_POOL == "debug" else False,
)
RemoteSessionLocal = sessionmaker(bind=remote_engine)

# SQLite metadata
meta_db_path = BASE_PATH / "espen.db"
meta_engine = create_engine("sqlite:///" + str(meta_db_path))
MetaSessionLocal = sessionmaker(bind=meta_engine)


def get_pool_status(engine_name: str, db_engine):
    """Get connection pool status for monitoring."""
    pool = db_engine.pool
    return {
        "engine": engine_name,
        "size": pool.size(),
        "checked_in": pool.checkedin(),
        "checked_out": pool.checkedout(),
        "overflow": pool.overflow(),
        "status": "healthy" if pool.checkedin() > 0 else "degraded"
    }
