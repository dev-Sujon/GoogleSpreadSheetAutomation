"""
json_payload.py
---------------
Convert independently-shaped pandas reports into a strict JSON-safe payload.

Nothing here performs HTTP. The result is an ordinary Python dict suitable for:
    requests.post(..., json=payload)
"""

from __future__ import annotations

import datetime as dt
import decimal
import json
import math
import uuid

import numpy as np
import pandas as pd

from src.config import format_report_date


def make_json_safe(value):
    """Convert one pandas/numpy/Python cell value into a JSON-safe value."""

    # Missing values first.
    if value is None or value is pd.NaT or value is pd.NA:
        return ""

    # Booleans before integers because bool is a subclass of int in Python.
    if isinstance(value, (bool, np.bool_)):
        return bool(value)

    if isinstance(value, (int, np.integer)):
        return int(value)

    if isinstance(value, (float, np.floating)):
        number = float(value)
        return number if math.isfinite(number) else ""

    if isinstance(value, decimal.Decimal):
        number = float(value)
        return number if math.isfinite(number) else ""

    if isinstance(value, np.datetime64):
        if np.isnat(value):
            return ""
        value = pd.Timestamp(value)

    if isinstance(value, pd.Timestamp):
        if value is pd.NaT:
            return ""
        if value.time() == dt.time(0, 0):
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M:%S")

    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")

    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, (str, int, float, bool)):
        return value

    return str(value)


def _normalize_headers(columns) -> list[str]:
    headers = [make_json_safe(column) for column in columns]
    return [str(header) for header in headers]


def dataframe_to_report(report_name: str, df: pd.DataFrame) -> dict:
    """Convert one DataFrame into one report object."""

    if not isinstance(df, pd.DataFrame):
        raise TypeError(
            f"Report '{report_name}' must be a pandas DataFrame, "
            f"got {type(df).__name__}."
        )

    headers = _normalize_headers(df.columns)
    rows = []

    # `astype(object)` preserves Timestamp/date objects for our converter.
    for row_number, row in enumerate(df.astype(object).itertuples(index=False, name=None), start=1):
        safe_row = [make_json_safe(cell) for cell in row]
        if len(safe_row) != len(headers):
            raise ValueError(
                f"Report '{report_name}' row {row_number} has "
                f"{len(safe_row)} values but has {len(headers)} headers."
            )
        rows.append(safe_row)

    return {
        "name": str(report_name),
        "headers": headers,
        "rows": rows,
    }


def build_payload(reports, report_date) -> dict:
    """
    Build the complete JSON-safe payload.

    `report_date` is transported as ISO `YYYY-MM-DD`. Apps Script converts it
    deterministically to the daily tab name `DD_MMM_YYYY`.
    """

    if not reports:
        raise ValueError("reports is empty - nothing to send")

    payload = {
        "version": 1,
        "run_id": uuid.uuid4().hex,
        "report_date": format_report_date(report_date),
        "reports": [
            dataframe_to_report(name, df)
            for name, df in reports
        ],
    }

    # Strict serialization check: fail locally before any network call.
    json.dumps(payload, allow_nan=False)
    return payload
