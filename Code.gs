/**
 * ERP -> Google Sheets Web App
 *
 * Python sends independently-shaped reports in one JSON payload.
 * Apps Script writes them into one previous-day tab.
 *
 * Daily tab rule:
 *     2026-09-20 -> 20_Sep_2026
 *
 * IMPORTANT:
 *     This project intentionally uses hard-coded configuration.
 *     Put the SAME spreadsheet ID and API token used in src/config.py below.
 *     Keep this Apps Script project private if the token is stored here.
 */

// ============================================================
// HARD-CODED CONFIGURATION — EDIT THESE VALUES
// ============================================================

const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";
const API_TOKEN = "PASTE_YOUR_API_TOKEN_HERE";

const BLANK_ROWS_BETWEEN_REPORTS = 2;
const AUTO_RESIZE_COLUMNS = true;
const TEST_SHEET_NAME = "__ERP_AUTOMATION_TEST__";
const EXPECTED_PAYLOAD_VERSION = 1;


function doGet() {
  try {
    // Health check verifies both configuration and target spreadsheet access.
    const config = getConfig_();
    SpreadsheetApp.openById(config.spreadsheetId);
    return jsonResponse_({
      success: true,
      message: "ERP Google Sheets Web App is running"
    });
  } catch (err) {
    return jsonResponse_({
      success: false,
      error: errorMessage_(err)
    });
  }
}


function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    const payload = parsePayload_(e);
    validatePayload_(payload);

    lock.waitLock(30000);

    try {
      return jsonResponse_(writeReports_(payload, sheetNameFromISO_(payload.report_date)));
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    return jsonResponse_({
      success: false,
      error: errorMessage_(err)
    });
  }
}


// ============================================================
// CONFIGURATION
// ============================================================

function getConfig_() {
  const spreadsheetId = String(SPREADSHEET_ID || "").trim();
  const apiToken = String(API_TOKEN || "").trim();

  if (!spreadsheetId || spreadsheetId.indexOf("PASTE_YOUR_") === 0) {
    throw new Error(
      "SPREADSHEET_ID is not configured. Edit Code.gs."
    );
  }

  if (!apiToken || apiToken.indexOf("PASTE_YOUR_") === 0) {
    throw new Error(
      "API_TOKEN is not configured. Edit Code.gs."
    );
  }

  if (apiToken.length < 24) {
    throw new Error(
      "API_TOKEN is too short. Use at least 24 characters."
    );
  }

  return {
    spreadsheetId: spreadsheetId,
    apiToken: apiToken
  };
}


// ============================================================
// PARSE + VALIDATE
// ============================================================

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("Empty request: no JSON body received.");
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error("Invalid JSON payload: " + errorMessage_(err));
  }
}


function validatePayload_(payload) {
  const config = getConfig_();

  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    throw new Error("Payload must be a JSON object.");
  }

  if (payload.token !== config.apiToken) {
    throw new Error("Unauthorized: invalid API token.");
  }

  if (payload.version !== EXPECTED_PAYLOAD_VERSION) {
    throw new Error(
      "Unsupported payload version: " + payload.version +
      ". Expected: " + EXPECTED_PAYLOAD_VERSION
    );
  }

  if (typeof payload.report_date !== "string" ||
      !/^\d{4}-\d{2}-\d{2}$/.test(payload.report_date)) {
    throw new Error(
      "Invalid report_date. Expected YYYY-MM-DD, e.g. 2026-09-20."
    );
  }

  // Reject impossible dates such as 2026-02-31.
  const normalizedDate = normalizeISODate_(payload.report_date);
  if (normalizedDate !== payload.report_date) {
    throw new Error("Invalid calendar date: " + payload.report_date);
  }

  if (!Array.isArray(payload.reports) || payload.reports.length === 0) {
    throw new Error("reports must be a non-empty list.");
  }

  payload.reports.forEach(function(report, index) {
    const label = "Report #" + (index + 1);

    if (!report || typeof report !== "object" || Array.isArray(report)) {
      throw new Error(label + " must be an object.");
    }

    if (typeof report.name !== "string" || report.name.trim() === "") {
      throw new Error(label + " is missing a name.");
    }

    if (!Array.isArray(report.headers)) {
      throw new Error(label + " headers must be a list.");
    }

    if (!Array.isArray(report.rows)) {
      throw new Error(label + " rows must be a list.");
    }

    report.headers.forEach(function(header, headerIndex) {
      if (header === null || header === undefined) {
        throw new Error(
          label + " header #" + (headerIndex + 1) + " is empty/null."
        );
      }
    });

    report.rows.forEach(function(row, rowIndex) {
      if (!Array.isArray(row)) {
        throw new Error(
          label + " row " + (rowIndex + 1) + " is not a list."
        );
      }

      if (row.length !== report.headers.length) {
        throw new Error(
          label + " row " + (rowIndex + 1) +
          " has " + row.length +
          " values but the report has " + report.headers.length + " headers."
        );
      }

      row.forEach(function(value, colIndex) {
        if (typeof value === "number" && !isFinite(value)) {
          throw new Error(
            label + " row " + (rowIndex + 1) +
            ", column " + (colIndex + 1) +
            " contains non-finite number."
          );
        }
      });
    });
  });
}


// ============================================================
// WRITE
// ============================================================

function writeReports_(payload, targetSheetName) {
  const config = getConfig_();
  const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);

  // Build everything before touching the target sheet. If payload conversion
  // fails, the existing daily sheet remains intact.
  const blocks = payload.reports.map(buildBlock_);

  let totalRows = 0;
  let maxCols = 1;

  blocks.forEach(function(block, index) {
    totalRows += block.values.length;
    maxCols = Math.max(maxCols, block.width);
    if (index < blocks.length - 1) {
      totalRows += BLANK_ROWS_BETWEEN_REPORTS;
    }
  });

  const sheet = getOrCreateSheet_(spreadsheet, targetSheetName);

  // Clear only after validation + block construction succeeded.
  sheet.clear();
  ensureGridSize_(sheet, totalRows, maxCols);

  let startRow = 1;

  blocks.forEach(function(block, index) {
    const range = sheet.getRange(
      startRow,
      1,
      block.values.length,
      block.width
    );

    range.setValues(block.values);
    range.setFontSize(10);

    // Report title
    sheet
      .getRange(startRow, 1)
      .setFontWeight("bold")
      .setFontSize(11);

    // Header
    if (block.hasHeader) {
      sheet
        .getRange(startRow + 1, 1, 1, block.width)
        .setFontWeight("bold")
        .setFontSize(10);
    }

    startRow += block.values.length;

    // Exactly two blank rows BETWEEN reports; none are required after the last.
    if (index < blocks.length - 1) {
      startRow += BLANK_ROWS_BETWEEN_REPORTS;
    }
  });

  if (AUTO_RESIZE_COLUMNS && maxCols > 0) {
    sheet.autoResizeColumns(1, maxCols);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    message: "Reports uploaded successfully",
    sheet: targetSheetName,
    report_count: blocks.length,
    rows_written: totalRows,
    run_id: payload.run_id || ""
  };
}


function getOrCreateSheet_(spreadsheet, name) {
  let sheet = spreadsheet.getSheetByName(name);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }

  return sheet;
}


function ensureGridSize_(sheet, rowsNeeded, colsNeeded) {
  if (rowsNeeded < 1) rowsNeeded = 1;
  if (colsNeeded < 1) colsNeeded = 1;

  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (maxRows < rowsNeeded) {
    sheet.insertRowsAfter(maxRows, rowsNeeded - maxRows);
  }

  if (maxCols < colsNeeded) {
    sheet.insertColumnsAfter(maxCols, colsNeeded - maxCols);
  }
}


function buildBlock_(report) {
  const width = Math.max(report.headers.length, 1);
  const values = [];

  // Title row
  values.push(padRow_([toCell_(report.name)], width));

  if (report.headers.length === 0) {
    values.push(["(no data)"]);

    return {
      values: values,
      width: width,
      hasHeader: false
    };
  }

  // Header row
  values.push(report.headers.map(toCell_));

  // Data rows
  report.rows.forEach(function(row) {
    values.push(row.map(toCell_));
  });

  // Final rectangular-shape guard.
  values.forEach(function(row, index) {
    if (row.length !== width) {
      throw new Error(
        "Range size mismatch in report '" + report.name +
        "', block row " + (index + 1) +
        ". Expected " + width + " cells, got " + row.length + "."
      );
    }
  });

  return {
    values: values,
    width: width,
    hasHeader: true
  };
}


function padRow_(row, width) {
  const result = row.slice();
  while (result.length < width) {
    result.push("");
  }
  return result;
}


function toCell_(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  const text = String(value);

  // Prevent report data beginning with '=' from being interpreted as a formula.
  // Normal strings remain untouched, so values such as 00123 remain text.
  if (/^[=+\-@]/.test(text)) {
    return "'" + text;
  }

  return text;
}


// ============================================================
// DATE HANDLING
// ============================================================

function normalizeISODate_(isoDate) {
  const parts = isoDate.split("-");
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);

  const utc = new Date(Date.UTC(year, month - 1, day));

  if (isNaN(utc.getTime())) {
    throw new Error("Invalid report_date: " + isoDate);
  }

  const normalized = Utilities.formatDate(utc, "UTC", "yyyy-MM-dd");
  return normalized;
}


function sheetNameFromISO_(isoDate) {
  // Manual month names make the tab name independent of spreadsheet locale.
  const months = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  const parts = isoDate.split("-");
  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);

  if (month < 1 || month > 12) {
    throw new Error("Invalid month in report_date: " + isoDate);
  }

  return (
    String(day).padStart(2, "0") + "_" +
    months[month - 1] + "_" +
    String(year).padStart(4, "0")
  );
}


// ============================================================
// TEST
// ============================================================

function testWrite() {
  const payload = {
    version: EXPECTED_PAYLOAD_VERSION,
    run_id: "TEST-RUN",
    report_date: "2026-09-20",
    reports: [
      {
        name: "Employee Report",
        headers: ["Employee ID", "Name", "Department", "Joining Date"],
        rows: [
          [101, "A", "IT", "2026-01-01"],
          [102, "B", "HR", "2026-02-15"]
        ]
      },
      {
        name: "Production Report",
        headers: ["Factory", "Production Date", "Production Qty"],
        rows: [
          ["MGL", "2026-09-20", 25000],
          ["MGSL", "2026-09-20", 30000]
        ]
      }
    ]
  };

  validatePayload_({
    version: payload.version,
    token: getConfig_().apiToken,
    report_date: payload.report_date,
    reports: payload.reports
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const result = writeReports_(payload, TEST_SHEET_NAME);
    Logger.log(JSON.stringify(result));
  } finally {
    lock.releaseLock();
  }
}


function deleteTestSheet() {
  const config = getConfig_();
  const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(TEST_SHEET_NAME);

  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
}


// ============================================================
// RESPONSE / ERROR HELPERS
// ============================================================

function errorMessage_(err) {
  return String(err && err.message ? err.message : err);
}


function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
