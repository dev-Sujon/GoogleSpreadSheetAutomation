"""
master_builder.py
-----------------
Create a human-readable master DataFrame from independently-shaped reports.

The original DataFrames are never modified. Each report keeps its own header,
column order, rows, and column count. Shorter reports are padded only in the
master preview so that the final DataFrame is rectangular.
"""

from __future__ import annotations

import pandas as pd


BLANK_ROWS_BETWEEN_REPORTS = 2


def build_master_dataframe(reports, include_title: bool = True) -> pd.DataFrame:
    """
    Build one vertical preview DataFrame.

    Parameters
    ----------
    reports:
        Iterable of (report_name, DataFrame).
    include_title:
        Put the report name above each report header.

    Returns
    -------
    pandas.DataFrame
        A rectangular preview containing each report and exactly two blank
        rows between reports. The real API payload is built from `reports`
        directly, not from this padded DataFrame.
    """

    if not reports:
        return pd.DataFrame()

    normalized = []

    for name, df in reports:
        if not isinstance(df, pd.DataFrame):
            raise TypeError(
                f"Report '{name}' must be a pandas DataFrame, "
                f"got {type(df).__name__}."
            )
        normalized.append((str(name), df))

    blocks = []
    for name, df in normalized:
        blocks.append(_report_to_rows(name, df, include_title))

    width = max(
        len(row)
        for block in blocks
        for row in block
    )

    all_rows = []
    for index, block in enumerate(blocks):
        if index > 0:
            all_rows.extend([
                [""] * width
                for _ in range(BLANK_ROWS_BETWEEN_REPORTS)
            ])

        for row in block:
            all_rows.append(row + [""] * (width - len(row)))

    return pd.DataFrame(all_rows, columns=range(width), dtype=object)


def _report_to_rows(name: str, df: pd.DataFrame, include_title: bool) -> list[list]:
    rows: list[list] = []

    if include_title:
        rows.append([name])

    if len(df.columns) == 0:
        rows.append(["(no data)"])
        return rows

    rows.append([str(column) for column in df.columns])

    # astype(object) prevents all-datetime DataFrames from turning timestamps
    # into raw integer nanoseconds when converted to Python lists.
    rows.extend(df.astype(object).values.tolist())
    return rows
