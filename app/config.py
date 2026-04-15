"""
Application configuration constants and path settings.
"""
import os
import logging
from pathlib import Path

# ── Logging ──────────────────────────────────────────────────────────
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, datefmt=LOG_DATE_FORMAT)
logger = logging.getLogger("excel-to-ppt")

# ── Directory paths ──────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"

# Ensure directories exist
for d in [UPLOAD_DIR, OUTPUT_DIR, STATIC_DIR]:
    d.mkdir(exist_ok=True)

# ── Application settings ─────────────────────────────────────────────
APP_TITLE = "Excel to PowerPoint Generator"
APP_DESCRIPTION = "Upload Excel and PPT files to generate reports"
APP_VERSION = "6.0.0"

# File upload limits
MAX_UPLOAD_SIZE_MB = 50
ALLOWED_EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")
ALLOWED_PPT_EXTENSIONS = (".pptx", ".ppt")

# File cleanup settings (seconds)
FILE_CLEANUP_MAX_AGE = 24 * 60 * 60  # 24 hours

# ── Mesh slide layout presets ────────────────────────────────────────
MESH_BACKHAUL_LAYOUT = {
    "left": 0.423,
    "top": 1.267,
    "width": 12.0,
    "height": 4.294,
}

MESH_FRONTHAUL_LAYOUT = {
    "left": 0.208,
    "top": 1.267,
    "width": 12.592,
    "height": 4.344,
}

# Default image placement
DEFAULT_IMAGE_LAYOUT = {
    "left": 0.423,
    "top": 1.4,
    "width": 12.0,
    "height": 5.6,
}

# ── COM automation settings ──────────────────────────────────────────
COM_MAX_RETRIES = 3
COM_RETRY_DELAY = 0.3  # seconds
COM_CLIPBOARD_DELAY = 0.2  # seconds
IMAGE_MIN_SIZE_BYTES = 500
IMAGE_MIN_UNIQUE_COLORS = 10
IMAGE_MIN_STDEV = 5.0
