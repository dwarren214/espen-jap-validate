"""Application configuration."""

import logging
import os
from pathlib import Path

from dotenv import load_dotenv

from response_guard import GuardrailThresholds

load_dotenv()

BASE_PATH = Path(__file__).resolve(strict=True).parent

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
if not ESPEN_CAMPAIGN_HUB_KEY:
    logging.warning("ESPEN_CAMPAIGN_HUB_KEY environment variable is not set. Campaign data may not be available.")
