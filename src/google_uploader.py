"""
google_uploader.py
------------------
HTTP client for the ERP -> Google Apps Script pipeline.

Configuration is imported from src.config. No .env file is used.
"""

from __future__ import annotations

import time
from typing import Any

import requests

from src.config import (
    API_TOKEN,
    APPS_SCRIPT_URL,
    HTTP_TIMEOUT_SECONDS,
    MAX_RETRIES,
    RETRY_BACKOFF_SECONDS,
    validate_config,
)


RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}


def check_apps_script_health(
    url: str | None = None,
    timeout: int = 30,
) -> dict[str, Any]:
    """
    Check that the deployed Apps Script Web App responds with JSON success.

    This GET check does not modify the target spreadsheet.
    """

    validate_config()
    target_url = (url or APPS_SCRIPT_URL).strip().rstrip("/")

    if not target_url.endswith("/exec"):
        return {
            "success": False,
            "error": "Apps Script URL must end with '/exec'.",
        }

    try:
        response = requests.get(
            target_url,
            timeout=timeout,
            allow_redirects=True,
            headers={"Accept": "application/json"},
        )
    except requests.exceptions.Timeout:
        return {
            "success": False,
            "error": f"Health check timed out after {timeout}s.",
        }
    except requests.exceptions.ConnectionError as exc:
        return {
            "success": False,
            "error": f"Health check connection error: {exc}",
        }
    except requests.exceptions.RequestException as exc:
        return {
            "success": False,
            "error": f"Health check failed: {exc}",
        }

    if response.status_code >= 400:
        return {
            "success": False,
            "error": (
                f"Health check HTTP {response.status_code}: "
                f"{response.text[:500]}"
            ),
        }

    try:
        result = response.json()
    except ValueError:
        return {
            "success": False,
            "error": (
                "Health check did not return JSON. "
                "Confirm that the deployed Web App URL ends in /exec "
                "and that its access setting permits the request."
            ),
        }

    if not isinstance(result, dict):
        return {
            "success": False,
            "error": "Health check returned a non-object JSON response.",
        }

    if not result.get("success"):
        return {
            "success": False,
            "error": result.get("error", "Apps Script health check failed."),
        }

    return result


def upload_to_google_apps_script(
    payload: dict,
    url: str | None = None,
    token: str | None = None,
    timeout: int | None = None,
) -> dict[str, Any]:
    """
    Upload one complete payload to Google Apps Script.

    The normal application should call this without overriding url/token/timeout.
    Optional arguments exist only for controlled tests.
    """

    validate_config()

    target_url = (url or APPS_SCRIPT_URL).strip().rstrip("/")
    target_token = token if token is not None else API_TOKEN
    request_timeout = timeout or HTTP_TIMEOUT_SECONDS

    if not target_url.endswith("/exec"):
        return _fail("Apps Script URL must end with '/exec'.")

    if not target_token or len(target_token.strip()) < 24:
        return _fail("Apps Script API token is missing or too short.")

    if not isinstance(payload, dict):
        return _fail("payload must be a Python dictionary.")

    if not payload.get("reports"):
        return _fail("payload contains no reports.")

    if not payload.get("report_date"):
        return _fail("payload is missing report_date.")

    # Copy rather than mutating the caller's dictionary.
    body = dict(payload)
    body["token"] = target_token

    run_id = body.get("run_id", "")
    total_attempts = MAX_RETRIES + 1

    for attempt in range(1, total_attempts + 1):
        started = time.perf_counter()

        try:
            response = requests.post(
                target_url,
                json=body,
                timeout=request_timeout,
                allow_redirects=True,
                headers={
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                },
            )

        except requests.exceptions.Timeout:
            elapsed = time.perf_counter() - started

            # A timeout is ambiguous: Apps Script may have completed the write.
            # Retrying is safe here because the daily operation is idempotent:
            # the same sheet is cleared and rewritten.
            if attempt < total_attempts:
                _sleep_before_retry(attempt)
                continue

            return _fail(
                f"Upload timed out after {request_timeout}s "
                f"(elapsed {elapsed:.1f}s). "
                "Check the Apps Script Executions log because the server "
                "may have completed the write."
            )

        except requests.exceptions.ConnectionError as exc:
            if attempt < total_attempts:
                _sleep_before_retry(attempt)
                continue

            return _fail(f"Upload connection error: {exc}")

        except requests.exceptions.RequestException as exc:
            return _fail(f"Upload request failed: {exc}")

        if response.status_code in RETRYABLE_STATUS_CODES and attempt < total_attempts:
            _sleep_before_retry(attempt)
            continue

        if response.status_code >= 400:
            return _fail(
                f"Apps Script HTTP {response.status_code}: "
                f"{response.text[:500]}"
            )

        break

    try:
        result = response.json()
    except ValueError:
        return _fail(
            "Apps Script did not return JSON. "
            "This usually indicates a Web App deployment/access problem "
            "or a Google login/HTML response. "
            f"Response preview: {response.text[:500]!r}"
        )

    if not isinstance(result, dict):
        return _fail(
            f"Unexpected Apps Script response type: "
            f"{type(result).__name__}"
        )

    if not result.get("success"):
        return _fail(
            "Apps Script rejected the upload: "
            f"{result.get('error', result)}"
        )

    # Helps ensure the response belongs to this exact request.
    returned_run_id = result.get("run_id")
    if run_id and returned_run_id and returned_run_id != run_id:
        return _fail(
            "Apps Script returned a different run_id than the uploaded payload."
        )

    result.setdefault("run_id", run_id)
    return result


def _sleep_before_retry(attempt: int) -> None:
    """Small deterministic backoff between retry attempts."""
    time.sleep(RETRY_BACKOFF_SECONDS * attempt)


def _fail(message: str) -> dict[str, Any]:
    return {
        "success": False,
        "error": message,
    }
