"""ESPEN SQL API - FastAPI application."""

import logging
import sys
from datetime import datetime

from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.responses import JSONResponse
from sqlalchemy import text

from db import (
    SessionLocal,
    RemoteSessionLocal,
    engine,
    remote_engine,
    get_pool_status,
)
from routers import oncho, espen, campaign

load_dotenv()

# Logging setup
logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

app = FastAPI()


@app.get("/up")
def status():
    return {"status": "up"}


@app.get("/health")
def health_check():
    """Comprehensive health check including database connections."""
    health_status = {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "databases": {}
    }

    # Check PostgreSQL connection
    try:
        with SessionLocal() as session:
            session.execute(text("SELECT 1"))
        health_status["databases"]["postgresql"] = {
            "status": "healthy",
            "pool": get_pool_status("postgresql", engine)
        }
    except Exception as e:
        health_status["status"] = "unhealthy"
        health_status["databases"]["postgresql"] = {
            "status": "unhealthy",
            "error": str(e)
        }

    # Check MSSQL connection
    try:
        with RemoteSessionLocal() as session:
            session.execute(text("SELECT 1"))
        health_status["databases"]["mssql"] = {
            "status": "healthy",
            "pool": get_pool_status("mssql", remote_engine)
        }
    except Exception as e:
        health_status["status"] = "unhealthy"
        health_status["databases"]["mssql"] = {
            "status": "unhealthy",
            "error": str(e)
        }

    status_code = 200 if health_status["status"] == "healthy" else 503
    return JSONResponse(content=health_status, status_code=status_code)


# Include routers
app.include_router(oncho.router)
app.include_router(espen.router)
app.include_router(campaign.router)
