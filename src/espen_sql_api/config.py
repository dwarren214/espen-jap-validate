"""Application configuration."""

import logging
import os
from pathlib import Path

from dotenv import load_dotenv

from .response_guard import GuardrailThresholds

load_dotenv()

BASE_PATH = Path(__file__).resolve(strict=True).parent
PROJECT_ROOT = BASE_PATH.parent.parent

# Guardrail configuration
MAX_RETURNED_ROWS = int(os.getenv("MAX_RETURNED_ROWS", "5000"))
MAX_RETURNED_BYTES = int(os.getenv("MAX_RETURNED_BYTES", "400000"))
PREVIEW_ROW_COUNT = int(os.getenv("PREVIEW_ROW_COUNT", "50"))

guardrail_thresholds = GuardrailThresholds(
    max_rows=MAX_RETURNED_ROWS,
    max_bytes=MAX_RETURNED_BYTES,
    preview_rows=PREVIEW_ROW_COUNT,
)

# Campaign Hub API
ESPEN_CAMPAIGN_HUB_KEY = os.getenv("ESPEN_CAMPAIGN_HUB_KEY")
CAMPAIGN_DB_DIR = Path(os.getenv("CAMPAIGN_DB_DIR", BASE_PATH / "campaign_dbs"))
if not ESPEN_CAMPAIGN_HUB_KEY:
    logging.warning("ESPEN_CAMPAIGN_HUB_KEY environment variable is not set. Campaign data may not be available.")

# JAP validation upload/storage configuration (additive scaffold for Story 1).
JAP_UPLOAD_BACKEND = os.getenv("JAP_UPLOAD_BACKEND", "local_fs")
JAP_UPLOAD_DIR = Path(
    os.getenv(
        "JAP_UPLOAD_DIR",
        str(PROJECT_ROOT / "local/jap_validation/uploads"),
    )
)
JAP_FILE_TTL_SECONDS = int(os.getenv("JAP_FILE_TTL_SECONDS", "86400"))
JAP_MAX_UPLOAD_BYTES = int(os.getenv("JAP_MAX_UPLOAD_BYTES", str(20 * 1024 * 1024)))

# Local workspace conventions used by JAP validation scaffolding.
JAP_METADATA_DIR = JAP_UPLOAD_DIR.parent / "metadata"
JAP_RUNS_DIR = JAP_UPLOAD_DIR.parent / "runs"
