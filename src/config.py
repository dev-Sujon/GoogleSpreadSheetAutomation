"""
config.py
---------
Single source of configuration for the Python side of the
ERP -> Google Sheets pipeline.

IMPORTANT:
    This project intentionally does NOT use .env / python-dotenv.
    Put the three real Google values below once and keep this file private.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

# ============================================================
# USER CONFIGURATION — EDIT THESE VALUES
# ============================================================

# Google Apps Script Web App deployment URL.
# Must be the deployed /exec URL, not /dev.
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzT3Di2kguDbdfPaongnvyBuw0cp6rjcNJ-Z7pvyW7bG2ZhTAziF32u0IJg_WdOT09K/exec"

# Shared secret used by Python and Apps Script.
API_TOKEN = "11SS%$S***154475MySciptAPI_132213###########SOn"

# Target Google Spreadsheet ID.
# Python does not need this for the POST itself, but keeping it here
# makes the Python project configuration explicit and easy to audit.
SPREADSHEET_ID = "1944VCMThyqliN_VNRyJggHEf6j3cVMUw-d__8j9i6I0"

# Daily reporting timezone.
REPORT_TIMEZONE = "Asia/Dhaka"

# HTTP settings.
HTTP_TIMEOUT_SECONDS = 120
MAX_RETRIES = 2
RETRY_BACKOFF_SECONDS = 2

# Required daily sheet format.
SHEET_DATE_FORMAT = "%d_%b_%Y"


# ============================================================
# INTERNAL CONSTANTS
# ============================================================

MONTHS_EN = (
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
)

PLACEHOLDER_MARKERS = (
    "PASTE_YOUR_",
    "YOUR_",
    "REPLACE_",
    "CHANGE_ME",
)


# ============================================================
# CONFIG VALIDATION
# ============================================================

def validate_config() -> None:
    """Validate configuration before any upload attempt."""

    values = {
        "APPS_SCRIPT_URL": APPS_SCRIPT_URL,
        "API_TOKEN": API_TOKEN,
        "SPREADSHEET_ID": SPREADSHEET_ID,
    }

    for name, value in values.items():
        if not isinstance(value, str) or not value.strip():
            raise RuntimeError(f"{name} is empty. Edit src/config.py.")

        upper_value = value.strip().upper()
        if any(marker in upper_value for marker in PLACEHOLDER_MARKERS):
            raise RuntimeError(
                f"{name} still contains a placeholder value. Edit src/config.py."
            )

    normalized_url = APPS_SCRIPT_URL.strip().rstrip("/")

    if not normalized_url.endswith("/exec"):
        raise RuntimeError(
            "APPS_SCRIPT_URL must be the deployed Apps Script Web App URL "
            "ending in '/exec'."
        )

    if len(API_TOKEN.strip()) < 24:
        raise RuntimeError(
            "API_TOKEN is too short. Use at least 24 characters; "
            "32+ random characters are recommended."
        )

    if len(SPREADSHEET_ID.strip()) < 20:
        raise RuntimeError(
            "SPREADSHEET_ID does not look valid. Copy the spreadsheet ID "
            "from the Google Sheets URL."
        )

    if HTTP_TIMEOUT_SECONDS < 10:
        raise RuntimeError("HTTP_TIMEOUT_SECONDS must be at least 10 seconds.")

    if MAX_RETRIES < 0 or MAX_RETRIES > 5:
        raise RuntimeError("MAX_RETRIES must be between 0 and 5.")

    if RETRY_BACKOFF_SECONDS < 0:
        raise RuntimeError("RETRY_BACKOFF_SECONDS cannot be negative.")

    try:
        ZoneInfo(REPORT_TIMEZONE)
    except Exception as exc:
        raise RuntimeError(
            f"Invalid REPORT_TIMEZONE='{REPORT_TIMEZONE}'."
        ) from exc


# ============================================================
# DATE HELPERS
# ============================================================

def get_now() -> datetime:
    """Return current date/time in the reporting timezone."""
    return datetime.now(ZoneInfo(REPORT_TIMEZONE))


def get_previous_report_date() -> date:
    """Return yesterday's calendar date in the reporting timezone."""
    return get_now().date() - timedelta(days=1)


def format_sheet_name(value: date | datetime | str) -> str:
    """
    Convert a date-like value to the required sheet name.

    Example:
        2026-09-20 -> 20_Sep_2026
    """

    if isinstance(value, str):
        value = date.fromisoformat(value)
    elif isinstance(value, datetime):
        value = value.date()

    if not isinstance(value, date):
        raise TypeError(
            "value must be a date, datetime, or YYYY-MM-DD string."
        )

    return f"{value.day:02d}_{MONTHS_EN[value.month - 1]}_{value.year:04d}"


def format_report_date(value: date | datetime | str) -> str:
    """Convert a date-like value into transport format YYYY-MM-DD."""

    if isinstance(value, str):
        return date.fromisoformat(value).isoformat()

    if isinstance(value, datetime):
        return value.date().isoformat()

    if isinstance(value, date):
        return value.isoformat()

    raise TypeError(
        "value must be a date, datetime, or YYYY-MM-DD string."
    )
