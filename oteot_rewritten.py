
# ========================================================================
# NOTEBOOK CELL 0
# ========================================================================

import pandas as pd
import numpy as np
import os



# ========================================================================
# NOTEBOOK CELL 1
# ========================================================================

oteot = pd.read_excel("Data\CategoryWiseOverTime.xlsx", sheet_name="Report_CategoryWiseOTWoven")



# ========================================================================
# NOTEBOOK CELL 2
# ========================================================================

# =============================================================================
# Category Wise Over Time report  ->  clean `oteot` DataFrame (+ formatted Excel)
# =============================================================================
import re
import warnings
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]   # fixed list -> no locale surprises
_EXCEL_EPOCH = pd.Timestamp("1899-12-30")               # Excel's 1900 date system (serial 1 = 1900-01-01)
_MIN_SERIAL, _MAX_SERIAL = 36526, 73415                 # 2000-01-01 .. 2100-12-31: only whole numbers in this
                                                        # window are treated as dates (guards normal numbers)
_DATE_RE = re.compile(r"^\d{2}-[A-Z][a-z]{2}-\d{4}$")   # dd-MMM-yyyy
_HOUR_KEYS = {"hr", "hour", "hours"}                    # 'Hr.' in day groups, 'Hour' in the Month Total group
_AMNT_KEYS = {"amnt", "amount"}


# -----------------------------------------------------------------------------
# Small helpers
# -----------------------------------------------------------------------------
def _clean_cell(value):
    """Normalise one cell: trim, unwrap, collapse whitespace; blanks -> NaN.
    Numbers, dates and other non-text values are returned untouched."""
    if isinstance(value, str):
        value = re.sub(r"[\u200b\ufeff]", "", value)     # zero-width characters
        value = re.sub(r"\s+", " ", value).strip()      # \n, \t, NBSP, double spaces -> one space
        return np.nan if value == "" else value
    if value is None or pd.isna(value):
        return np.nan
    return value


def _key(value):
    """Comparison key for header text: 'Hr.' -> 'hr', 'Month Total' -> 'monthtotal'."""
    return re.sub(r"[^a-z]", "", str(value).lower())


def _fmt_date(d):
    return f"{d.day:02d}-{_MONTHS[d.month - 1]}-{d.year}"


def _to_date_label(value):
    """Return 'dd-MMM-yyyy' for a real date or an Excel serial number, else None."""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return _fmt_date(value)
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        if float(value).is_integer() and _MIN_SERIAL <= value <= _MAX_SERIAL:
            return _fmt_date(_EXCEL_EPOCH + pd.Timedelta(days=int(value)))
    return None


def _ffill(values):
    """Forward-fill a list until the next non-blank value (no pandas downcast warnings)."""
    out, last = [], np.nan
    for v in values:
        if pd.isna(v):
            out.append(last)
        else:
            last = v
            out.append(v)
    return out


# -----------------------------------------------------------------------------
# Main function
# -----------------------------------------------------------------------------
def clean_category_wise_overtime(
    input_file,
    source_sheet="Report_CategoryWiseOTWoven",
    output_file=None,
    rows_to_drop=11,
    fill_columns=("Factory", "Type", "Category"),
    month_total_span=None,
    output_sheet="oteot",
):
    """
    Clean the RMG 'Category Wise Over Time' report and return it as DataFrame `oteot`.

    Layout of the returned frame (same as the target sheet, headers stay in the grid):
        row 0  : 'Month Total' and one dd-MMM-yyyy date per day, sitting on that group's 'Emp' column
        row 1  : Factory | Type | Category | Type | Emp | Hour | Amnt. | Emp | Hr. | Amnt. | ...
        row 2+ : data
    Columns are 0..n-1 because the report has two header rows.

    Parameters
    ----------
    rows_to_drop     : report-information rows above the date row (11 -> Excel rows 1-11).
    fill_columns     : header names to forward-fill. If a header appears twice ('Type' is both
                       the Woven/Knit column and the OT/EOT column) only the FIRST, leftmost one
                       is filled. The OT/EOT column is left as-is so that 'Total' rows keep a
                       blank there instead of being mislabelled 'EOT'.
    month_total_span : None -> merge 'Month Total' across its own Emp/Hour/Amnt. columns (3).
                       Pass 2 to merge only 'Month Total' + the next blank cell.
    """
    # ---- 1. Read the whole worksheet (header=None so DataFrame rows == Excel rows) ----------
    input_file = Path(input_file)
    if not input_file.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    with pd.ExcelFile(input_file) as xl:
        if source_sheet not in xl.sheet_names:
            raise ValueError(f"Sheet '{source_sheet}' not found. Available sheets: {xl.sheet_names}")
        # keep_default_na=False: text such as 'NA' or 'None' is kept; blanks are handled below
        raw = xl.parse(source_sheet, header=None, dtype=object, keep_default_na=False)

    # ---- 2. Delete the first `rows_to_drop` rows (zero-based: iloc[11:] == Excel rows 12+) ---
    df = raw.iloc[rows_to_drop:].reset_index(drop=True)
    if df.empty:
        raise ValueError(f"The sheet has no rows after removing the first {rows_to_drop} rows.")
    df.columns = range(df.shape[1])

    # ---- 3. Normalise every cell (whitespace, wrapped text, NaN / None / '' -> NaN) ---------
    elementwise = df.map if hasattr(df, "map") else df.applymap     # pandas >= 2.1 / older
    df = elementwise(_clean_cell).astype(object)

    # ---- 4. Delete totally blank rows and reset the index ------------------------------------
    df = df.dropna(how="all").reset_index(drop=True)

    # ---- 5. Excel serial dates in the first row -> dd-MMM-yyyy -------------------------------
    for c in df.columns:
        label = _to_date_label(df.iat[0, c])
        if label:
            df.iat[0, c] = label

    # ---- 6. Locate the header rows (row 0 = dates, row 1 = Factory/Type/.../Emp/Hr./Amnt.) ---
    header_rows = [r for r in df.index[:5] if any(_key(v) == "factory" for v in df.loc[r].dropna())]
    if header_rows != [1]:
        raise ValueError(
            "Expected the date row at position 0 and the 'Factory ... Emp/Hr./Amnt.' row at "
            f"position 1 after removing {rows_to_drop} rows (found header row(s): {header_rows}). "
            "Check `rows_to_drop`."
        )
    hdr = {c: _key(v) for c, v in df.loc[1].items() if pd.notna(v)}       # {column: header key}

    # ---- 7. Detect the Emp / Hr. / Amnt. groups (each = one day or the Month Total) ----------
    emp_cols = sorted(c for c, k in hdr.items() if k == "emp")
    if not emp_cols:
        raise ValueError("No 'Emp' columns found in the header row.")
    bounds = emp_cols[1:] + [df.shape[1]]
    groups = []
    for emp, nxt in zip(emp_cols, bounds):
        hr = next((c for c in range(emp + 1, nxt) if hdr.get(c) in _HOUR_KEYS), None)
        am = next((c for c in range(emp + 1, nxt) if hdr.get(c) in _AMNT_KEYS and hr and c > hr), None)
        if hr is None or am is None:
            raise ValueError(f"Column group starting at column index {emp} lacks an Hr./Amnt. column.")
        groups.append({"emp": emp, "hr": hr, "amnt": am, "label": None, "kind": None})

    # ---- 8. Attach 'Month Total' / date headers to their group ------------------------------
    # A label belongs to the first group whose Emp column is at or to the right of the label.
    # Assumption: the blank cells after a label are merged-cell leftovers, not data.
    labels = []
    for c, v in df.loc[0].items():
        if pd.isna(v):
            continue
        if _key(v) == "monthtotal":
            labels.append((c, "Month Total", "month_total"))
        elif isinstance(v, str) and _DATE_RE.match(v):
            labels.append((c, v, "date"))
    if not any(kind == "date" for _, _, kind in labels):
        raise ValueError("No date headers found in the first row.")
    if not any(kind == "month_total" for _, _, kind in labels):
        warnings.warn("'Month Total' was not found in the first row.")

    for col, text, kind in labels:
        grp = next((g for g in groups if g["emp"] >= col), None)
        if grp is None or grp["label"] is not None:
            raise ValueError(f"Header '{text}' (column index {col}) cannot be matched to a unique Emp/Hr./Amnt. group.")
        if any(pd.notna(df.iat[0, x]) for x in (grp["hr"], grp["amnt"])):
            raise ValueError(f"Unexpected content in the cells after '{text}'; refusing to merge over data.")
        grp["label"], grp["kind"] = text, kind
        if col != grp["emp"]:                     # park the label on the group's Emp column
            df.iat[0, col] = np.nan
            df.iat[0, grp["emp"]] = text
    for g in groups:
        if g["label"] is None:
            warnings.warn(f"Emp/Hr./Amnt. group at column index {g['emp']} has no date header above it.")

    # ---- 9. Forward-fill Factory / Type / Category (data rows only) --------------------------
    fill_cols = []
    for name in fill_columns:
        col = next((c for c, k in sorted(hdr.items()) if k == _key(name)), None)   # leftmost match
        if col is None:
            raise ValueError(f"Could not find the '{name}' header to forward-fill.")
        fill_cols.append(col)
    for c in fill_cols:
        df.loc[2:, c] = _ffill(df.loc[2:, c].tolist())

    # ---- 10. Drop columns that are completely blank (merged-cell leftovers) ------------------
    keep = [c for c in df.columns if df[c].notna().any()]
    new_pos = {old: i for i, old in enumerate(keep)}
    oteot = df[keep].copy()
    oteot.columns = range(len(keep))
    for g in groups:
        g["emp"], g["hr"], g["amnt"] = new_pos[g["emp"]], new_pos[g["hr"]], new_pos[g["amnt"]]
        if not (g["hr"] == g["emp"] + 1 and g["amnt"] == g["emp"] + 2):
            raise ValueError(f"Group '{g['label']}' is not three adjacent columns after cleaning.")

    # ---- 11. Export + OpenPyXL formatting ----------------------------------------------------
    if output_file is not None:
        output_file = Path(output_file)
        if output_file.suffix.lower() != ".xlsx":
            raise ValueError("`output_file` must end with .xlsx")
        output_file.parent.mkdir(parents=True, exist_ok=True)
        try:
            oteot.to_excel(output_file, sheet_name=output_sheet, header=False, index=False)
        except PermissionError as exc:
            raise PermissionError(f"Cannot write {output_file}. Is it open in Excel?") from exc
        _format_workbook(output_file, output_sheet, oteot, groups, month_total_span)

    return oteot


# -----------------------------------------------------------------------------
# Excel formatting (pandas cannot write merged cells, so this is done with OpenPyXL)
# -----------------------------------------------------------------------------
def _format_workbook(path, sheet_name, oteot, groups, month_total_span):
    wb = load_workbook(path)
    ws = wb[sheet_name]
    n_rows, n_cols = oteot.shape

    centre = Alignment(horizontal="center", vertical="center", wrap_text=False)
    left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    right = Alignment(horizontal="right", vertical="center", wrap_text=False)

    # Alignment: rows 1-2 (date + Emp/Hr./Amnt. headers) centred; body text left, numbers right
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=n_rows, max_col=n_cols), start=1):
        for cell in row:
            if r <= 2:
                cell.alignment = centre
            elif isinstance(cell.value, str):
                cell.alignment = left
            else:
                cell.alignment = right

    # Merge only the header cell of each group: date -> its 3 columns, Month Total -> its group
    for g in groups:
        if g["label"] is None:
            continue
        span = month_total_span if (g["kind"] == "month_total" and month_total_span) else 3
        start = g["emp"] + 1                                   # openpyxl columns are 1-based
        ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=start + span - 1)
        ws.cell(row=1, column=start).alignment = centre        # anchor cell carries the alignment

    # Column widths from content (row 1 skipped: its merged labels span several columns)
    for idx in range(1, n_cols + 1):
        letter = get_column_letter(idx)
        longest = max((len(str(c.value)) for c in ws[letter][1:] if c.value is not None), default=0)
        ws.column_dimensions[letter].width = min(max(longest + 2, 8), 45)

    wb.save(path)


# -----------------------------------------------------------------------------
# Example
# -----------------------------------------------------------------------------
oteot = clean_category_wise_overtime(
    input_file=r"Data\CategoryWiseOverTime.xlsx",
    source_sheet="Report_CategoryWiseOTWoven",
    output_file=r"Data\oteot.xlsx",
)

display(oteot.head(20))



# ========================================================================
# NOTEBOOK CELL 3
# ========================================================================

display(oteot.head(2))



# ========================================================================
# NOTEBOOK CELL 4
# ========================================================================

# Syntax: df.iat[row_position, column_position] = new_value
oteot.iat[1, 1] = "Factory_Type" 
oteot.iat[1, 3] = "OT/EOT"



# ========================================================================
# NOTEBOOK CELL 5
# ========================================================================

display(oteot.head(5))



# ========================================================================
# NOTEBOOK CELL 6
# ========================================================================

# =============================================================================
# oteot  ->  OT / Extra-OT management report
# =============================================================================
import re
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# -----------------------------------------------------------------------------
# Fixed, zero-based column positions inside `oteot`
# row 0 = date headers, row 1 = column headers, data from row 2
# -----------------------------------------------------------------------------
COLS = {
    "Factory": 0, "Type": 1, "Category": 2, "OTEOT": 3,
    "MT_Emp": 4, "MT_Hour": 5, "MT_Amnt": 6,     # Month Total group
    "PD_Emp": 7, "PD_Hour": 8, "PD_Amnt": 9,     # Previous Day group
}

NUMERIC_COLS = [
    "MT_Emp", "MT_Hour", "MT_Amnt",
    "PD_Emp", "PD_Hour", "PD_Amnt"
]

# -----------------------------------------------------------------------------
# Keep the original KPI calculations.
# The management report below will display only the columns requested in the
# example report.
# -----------------------------------------------------------------------------
KPI_SPEC = [
    ("Number of OT Employees",               "PD_Emp",  "OT"),
    ("Number of Extra OT Employees",         "PD_Emp",  "EOT"),
    ("Previous Days' OT Hours",              "PD_Hour", "OT"),
    ("Previous Days' Extra OT Hours",        "PD_Hour", "EOT"),
    ("Monthly Cumulative OT Hours",          "MT_Hour", "OT"),
    ("Monthly Cumulative Extra OT Hours",    "MT_Hour", "EOT"),
    ("Monthly OT Expense",                   "MT_Amnt", "OT"),
    ("Monthly Extra OT Expense",             "MT_Amnt", "EOT"),
]

KPI_COLS = [name for name, _, _ in KPI_SPEC]

# -----------------------------------------------------------------------------
# Management-report columns exactly in the requested order.
# "Number of OT Employees" is intentionally not shown.
# -----------------------------------------------------------------------------
REPORT_KPI_COLS = [
    "Number of Extra OT Employees",
    "Previous Days' OT Hours",
    "Previous Days' Extra OT Hours",
    "Monthly Cumulative OT Hours",
    "Monthly Cumulative Extra OT Hours",
    "Monthly OT Expense",
    "Monthly Extra OT Expense",
]

# Display labels for the management report.
# The spelling of "Cumultiave" follows the example supplied in the request.
REPORT_HEADERS = {
    "Number of Extra OT Employees": "Number of Extra OT Employees",
    "Previous Days' OT Hours": "Previous Days' OT Hours",
    "Previous Days' Extra OT Hours": "Previous Days' Extra OT Hours",
    "Monthly Cumulative OT Hours": "Monthly Cumulative OT Hours",
    "Monthly Cumulative Extra OT Hours": "Monthly Cumultiave Extra OT Hours",
    "Monthly OT Expense": "Monthly OT Expense",
    "Monthly Extra OT Expense": "Monthly Extra OT Expense",
}

# Labels used by the source report for pre-calculated subtotal / total rows.
_AGG_CATEGORY_KEYS = {"total", "factot", "facteot"}
_AGG_FACTORY_KEYS = {"totalot", "totaleot", "grandtotal"}


def _norm_key(v):
    return re.sub(r"[^a-z]", "", str(v).lower())


# -----------------------------------------------------------------------------
# Main function
# -----------------------------------------------------------------------------
def build_ot_kpi_report(oteot, output_file=None):
    """
    Build the OT / Extra-OT management report from cleaned `oteot`.

    Returns
    -------
    detail_kpi
        Original detailed KPI table with all 8 calculated KPIs.
    management_report
        Management table in the requested 8-column layout:
        Factory + 7 report KPIs.
    report_info
        Previous Day and Month Total period information.
    """

    if oteot.shape[0] < 3 or oteot.shape[1] <= max(COLS.values()):
        raise ValueError(
            f"`oteot` must have >=2 header rows, >=1 data row, and "
            f">={max(COLS.values()) + 1} columns."
        )

    # -------------------------------------------------------------------------
    # 1. Preserve row 0 reporting information
    # -------------------------------------------------------------------------
    date_labels = [
        v for v in oteot.iloc[0]
        if isinstance(v, str) and re.match(r"^\d{2}-[A-Za-z]{3}-\d{4}$", v)
    ]

    report_info = {}

    if date_labels:
        parsed = sorted(pd.to_datetime(date_labels, format="%d-%b-%Y"))
        report_info["Previous Day"] = parsed[-1].strftime("%d-%b-%Y")
        report_info["Month Total period"] = (
            f"{parsed[0]:%d-%b-%Y} to {parsed[-1]:%d-%b-%Y}"
        )

    # -------------------------------------------------------------------------
    # 2. Data rows only
    # -------------------------------------------------------------------------
    data = oteot.iloc[2:].reset_index(drop=True)

    data = data.rename(
        columns={v: k for k, v in COLS.items()}
    )[list(COLS.keys())].copy()

    # -------------------------------------------------------------------------
    # 3. Clean text columns
    # -------------------------------------------------------------------------
    for c in ("Factory", "Type", "Category", "OTEOT"):
        data[c] = data[c].apply(
            lambda v:
                re.sub(r"\s+", " ", str(v)).strip()
                if pd.notna(v)
                else np.nan
        )

    data = data.dropna(
        subset=["Factory", "Category", "OTEOT"]
    ).reset_index(drop=True)

    # -------------------------------------------------------------------------
    # 4. Convert numeric columns safely
    # -------------------------------------------------------------------------
    for c in NUMERIC_COLS:
        data[c] = pd.to_numeric(
            data[c],
            errors="coerce"
        ).fillna(0)

    # -------------------------------------------------------------------------
    # 5. Remove source report subtotal / total rows
    # -------------------------------------------------------------------------
    is_agg = (
        data["Category"].apply(_norm_key).isin(_AGG_CATEGORY_KEYS)
        |
        data["Factory"].apply(_norm_key).isin(_AGG_FACTORY_KEYS)
    )

    real = data.loc[~is_agg].reset_index(drop=True)

    if real.empty:
        raise ValueError(
            "No real category rows remained after removing "
            "subtotal/grand-total rows."
        )

    # -------------------------------------------------------------------------
    # 6. Detail KPI table
    # -------------------------------------------------------------------------
    grouped = (
        real.groupby(
            ["Factory", "Type", "Category", "OTEOT"],
            sort=False
        )[NUMERIC_COLS]
        .sum()
        .reset_index()
    )

    pivot = grouped.pivot_table(
        index=["Factory", "Type", "Category"],
        columns="OTEOT",
        values=NUMERIC_COLS,
        fill_value=0,
        aggfunc="sum",
    )

    detail_kpi = pd.DataFrame(index=pivot.index)

    for kpi_name, source_col, ot_eot in KPI_SPEC:
        detail_kpi[kpi_name] = (
            pivot[(source_col, ot_eot)]
            if (source_col, ot_eot) in pivot.columns
            else 0
        )

    detail_kpi = detail_kpi.reset_index()

    # -------------------------------------------------------------------------
    # 7. Roll-ups
    # -------------------------------------------------------------------------
    type_summary = (
        detail_kpi
        .groupby(["Factory", "Type"], sort=False)[KPI_COLS]
        .sum()
        .reset_index()
    )

    factory_summary = (
        detail_kpi
        .groupby(["Factory"], sort=False)[KPI_COLS]
        .sum()
        .reset_index()
    )

    # -------------------------------------------------------------------------
    # 8. Hierarchical management report
    #
    # The requested report treats the base/Woven data as the factory row.
    # Additional types are shown separately as Factory-Type rows.
    #
    # Example:
    #
    # Factory      Extra OT Emp   Prev OT   Prev EOT   Monthly OT ...
    # MGL              360          ...
    # MGL-Knit         192          ...
    # MGNSL             10          ...
    # ...
    # Total                          5129       1081    ...
    #
    # This prevents the factory row from adding MGL-Woven + MGL-Knit together
    # while still keeping the MGL-Knit values visible as a separate row.
    # -------------------------------------------------------------------------
    rows = []

    BASE_TYPE_KEY = "woven"

    for factory in dict.fromkeys(real["Factory"]):

        factory_types = type_summary.loc[
            type_summary["Factory"] == factory,
            "Type"
        ].tolist()

        if not factory_types:
            continue

        # -------------------------------------------------------------
        # Factory row:
        #   1. Prefer the Woven type when it exists.
        #   2. Otherwise, if there is only one type, use that type.
        #   3. Otherwise, use the first type in source order.
        # -------------------------------------------------------------
        woven_type = next(
            (
                t for t in factory_types
                if _norm_key(t) == BASE_TYPE_KEY
            ),
            None
        )

        if woven_type is not None:
            factory_type = woven_type
        elif len(factory_types) == 1:
            factory_type = factory_types[0]
        else:
            factory_type = factory_types[0]

        base_row = type_summary.loc[
            (type_summary["Factory"] == factory)
            &
            (type_summary["Type"] == factory_type)
        ]

        if base_row.empty:
            continue

        rows.append(
            {
                "Factory": factory,
                **base_row.iloc[0][REPORT_KPI_COLS].to_dict()
            }
        )

        # -------------------------------------------------------------
        # Additional type rows:
        # Show every type except the base factory row.
        # Therefore MGL + Woven becomes:
        #   MGL
        #   MGL-Knit
        # and NOT:
        #   MGL
        #   MGL-Woven
        #   MGL-Knit
        # -------------------------------------------------------------
        for extra_type in factory_types:
            if extra_type == factory_type:
                continue

            type_row = type_summary.loc[
                (type_summary["Factory"] == factory)
                &
                (type_summary["Type"] == extra_type)
            ]

            if type_row.empty:
                continue

            rows.append(
                {
                    "Factory": f"{factory}-{extra_type}",
                    **type_row.iloc[0][REPORT_KPI_COLS].to_dict()
                }
            )

    # -------------------------------------------------------------------------
    # Total row:
    # Sum all real Factory + Type + Category rows so additional types such as
    # MGL-Knit are included in the grand total.
    #
    # Number of Extra OT Employees is intentionally blank because the requested
    # example does not show a cumulative employee count in the Total row.
    # -------------------------------------------------------------------------
    total_row = {"Factory": "Total"}

    for col in REPORT_KPI_COLS:
        if col == "Number of Extra OT Employees":
            total_row[col] = np.nan
        else:
            total_row[col] = detail_kpi[col].sum()

    rows.append(total_row)

    management_report = pd.DataFrame(
        rows,
        columns=["Factory"] + REPORT_KPI_COLS
    )

    # Rename only the management-report presentation headers.
    management_report = management_report.rename(
        columns=REPORT_HEADERS
    )

    if output_file is not None:
        _export_workbook(
            output_file,
            detail_kpi,
            management_report,
            report_info
        )

    return detail_kpi, management_report, report_info


# -----------------------------------------------------------------------------
# Excel formatting
# -----------------------------------------------------------------------------
_HEADER_FILL = PatternFill(
    "solid",
    fgColor="1F4E78"
)

_HEADER_FONT = Font(
    name="Arial",
    size=10,
    bold=True,
    color="FFFFFF"
)

_BODY_FONT = Font(
    name="Arial",
    size=10
)

_TOTAL_FONT = Font(
    name="Arial",
    size=10,
    bold=True
)

_TOTAL_FILL = PatternFill(
    "solid",
    fgColor="D9E1F2"
)

_THIN = Side(
    style="thin",
    color="BFBFBF"
)

_BORDER = Border(
    left=_THIN,
    right=_THIN,
    top=_THIN,
    bottom=_THIN
)

_COUNT_FMT = "#,##0"
_HOUR_FMT = "#,##0"
_MONEY_FMT = "#,##0"


def _kpi_format(column_name):
    if "Expense" in column_name:
        return _MONEY_FMT

    if "Hour" in column_name:
        return _HOUR_FMT

    if "Employee" in column_name:
        return _COUNT_FMT

    return _COUNT_FMT


def _write_sheet(
    wb,
    title,
    df,
    label_cols,
    freeze_after_col,
    bold_labels=()
):
    ws = wb.create_sheet(title)

    ws.append(list(df.columns))

    # Header formatting
    for cell in ws[1]:
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = _BORDER

    # Body
    for row in dataframe_to_rows(
        df,
        index=False,
        header=False
    ):
        ws.append(row)

    # Body formatting
    for r in range(2, ws.max_row + 1):
        is_bold_row = (
            ws.cell(r, 1).value in bold_labels
        )

        for c in range(
            1,
            ws.max_column + 1
        ):
            cell = ws.cell(r, c)
            cell.border = _BORDER

            col_name = df.columns[c - 1]

            if col_name in label_cols:
                cell.font = (
                    _TOTAL_FONT
                    if is_bold_row
                    else _BODY_FONT
                )
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )
            else:
                cell.number_format = _kpi_format(col_name)

                cell.font = (
                    _TOTAL_FONT
                    if is_bold_row
                    else _BODY_FONT
                )

                cell.alignment = Alignment(
                    horizontal="right",
                    vertical="center"
                )

            if is_bold_row:
                cell.fill = _TOTAL_FILL

    # Column widths
    for idx in range(
        1,
        ws.max_column + 1
    ):
        letter = get_column_letter(idx)

        longest = max(
            (
                len(str(c.value))
                for c in ws[letter]
                if c.value is not None
            ),
            default=8
        )

        ws.column_dimensions[letter].width = min(
            max(longest + 2, 10),
            42
        )

    # Slightly taller header
    ws.row_dimensions[1].height = 32

    ws.freeze_panes = ws.cell(
        row=2,
        column=freeze_after_col + 1
    ).coordinate

    ws.auto_filter.ref = ws.dimensions

    return ws


def _export_workbook(
    output_file,
    detail_kpi,
    management_report,
    report_info
):
    output_file = Path(output_file)

    if output_file.suffix.lower() != ".xlsx":
        raise ValueError(
            "`output_file` must end with .xlsx"
        )

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    wb = Workbook()
    wb.remove(wb.active)

    # -------------------------------------------------------------------------
    # Management Report
    # -------------------------------------------------------------------------
    management_ws = _write_sheet(
        wb,
        "Management Report",
        management_report,
        label_cols={"Factory"},
        freeze_after_col=1,
        bold_labels=(
            {"Total"}
            |
            {
                r
                for r in management_report["Factory"]
                if "-" not in str(r)
                and r != "Total"
            }
        )
    )

    if report_info:
        banner = "  |  ".join(
            f"{k}: {v}"
            for k, v in report_info.items()
        )

        management_ws.insert_rows(1)

        management_ws.cell(
            row=1,
            column=1,
            value=banner
        ).font = Font(
            name="Arial",
            size=10,
            italic=True
        )

        management_ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=management_ws.max_column
        )

        management_ws.cell(
            row=1,
            column=1
        ).alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        management_ws.freeze_panes = "B3"

    # -------------------------------------------------------------------------
    # Detail KPI
    # -------------------------------------------------------------------------
    _write_sheet(
        wb,
        "Detail KPI",
        detail_kpi,
        label_cols={
            "Factory",
            "Type",
            "Category"
        },
        freeze_after_col=3,
    )

    try:
        wb.save(output_file)

    except PermissionError as exc:
        raise PermissionError(
            f"Cannot write {output_file}. "
            "Is it open in Excel?"
        ) from exc


# -----------------------------------------------------------------------------
# Run
# -----------------------------------------------------------------------------
# `clean_category_wise_overtime()` is already defined in Cell 2, so there is
# no external `clean_ot` module dependency here.
oteot = clean_category_wise_overtime(
    input_file=r"Data\CategoryWiseOverTime.xlsx",
    source_sheet="Report_CategoryWiseOTWoven",
)

detail_kpi, management_report, report_info = build_ot_kpi_report(
    oteot,
    output_file=r"Data\OT_KPI_Report.xlsx"
)

display(management_report)

