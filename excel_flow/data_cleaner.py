"""Clean and reshape tabular data for analysis.

Missing-value policy
--------------------
``summarize_missing_values`` counts only pandas-null values (``NaN`` / ``None`` /
``NaT``). Whitespace-only strings are **not** treated as missing values; they are
handled by blank-row removal and string stripping instead.
"""

from __future__ import annotations

from typing import Any

import pandas as pd

from excel_flow.validators import (
    ExcelFlowError,
    has_cleaning_operations,
    validate_columns_to_drop,
    validate_rename_mapping,
)


def drop_columns(dataframe: pd.DataFrame, columns_to_drop: list[str] | None) -> pd.DataFrame:
    """Return a copy with the selected columns removed.

    Unknown columns are ignored. An empty selection leaves the data unchanged.
    """
    validated = validate_columns_to_drop(dataframe, columns_to_drop)
    result = dataframe.copy()
    if not validated:
        return result
    return result.drop(columns=validated)


def rename_columns(
    dataframe: pd.DataFrame,
    rename_mapping: dict[str, str] | None,
) -> pd.DataFrame:
    """Return a copy with validated column renames applied."""
    validated = validate_rename_mapping(dataframe, rename_mapping)
    result = dataframe.copy()
    if not validated:
        return result
    return result.rename(columns=validated)


def strip_string_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Strip leading/trailing whitespace from object/string columns only."""
    result = dataframe.copy()
    for column in result.columns:
        series = result[column]
        if pd.api.types.is_string_dtype(series) or series.dtype == object:
            result[column] = series.map(_strip_cell_value)
    return result


def _strip_cell_value(value: Any) -> Any:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return value
    if pd.isna(value):
        return value
    if isinstance(value, str):
        return value.strip()
    return value


def _is_blank_cell(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def drop_blank_rows(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Drop rows where every cell is null or whitespace-only.

    Returns the cleaned DataFrame and the number of removed rows.
    """
    result = dataframe.copy()
    if result.empty:
        return result, 0

    blank_mask = result.map(_is_blank_cell).all(axis=1)
    removed = int(blank_mask.sum())
    if removed:
        result = result.loc[~blank_mask].reset_index(drop=True)
    return result, removed


def drop_duplicate_rows(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Drop fully duplicated rows, keeping the first occurrence."""
    result = dataframe.copy()
    before = len(result)
    result = result.drop_duplicates(keep="first").reset_index(drop=True)
    removed = before - len(result)
    return result, removed


def summarize_missing_values(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return missing-value counts and rates per column.

    Only pandas-null values are counted. Whitespace-only strings are excluded.
    """
    if len(dataframe.columns) == 0:
        return pd.DataFrame(columns=["列名", "欠損件数", "欠損率"])

    row_count = max(len(dataframe), 1)
    missing_counts = dataframe.isna().sum()
    rows: list[dict[str, Any]] = []
    for column in dataframe.columns:
        count = int(missing_counts[column])
        rate = count / row_count * 100
        rows.append(
            {
                "列名": str(column),
                "欠損件数": count,
                "欠損率": f"{rate:.1f}%",
            }
        )
    return pd.DataFrame(rows)


def build_cleaning_summary(
    *,
    original: pd.DataFrame,
    cleaned: pd.DataFrame,
    blank_rows_removed: int = 0,
    duplicates_removed: int = 0,
) -> dict[str, int]:
    """Build a before/after cleaning summary."""
    original_rows = len(original)
    cleaned_rows = len(cleaned)
    original_cols = len(original.columns)
    cleaned_cols = len(cleaned.columns)
    return {
        "original_rows": original_rows,
        "cleaned_rows": cleaned_rows,
        "removed_rows": original_rows - cleaned_rows,
        "original_columns": original_cols,
        "cleaned_columns": cleaned_cols,
        "removed_columns": original_cols - cleaned_cols,
        "blank_rows_removed": blank_rows_removed,
        "duplicates_removed": duplicates_removed,
    }


def clean_dataframe(
    dataframe: pd.DataFrame,
    *,
    rename_mapping: dict[str, str] | None = None,
    columns_to_drop: list[str] | None = None,
    strip_whitespace: bool = False,
    remove_blank_rows: bool = False,
    remove_duplicates: bool = False,
) -> tuple[pd.DataFrame, dict[str, int]]:
    """Apply cleaning operations in a fixed, reproducible order.

    Order:
    1. rename columns
    2. drop columns
    3. strip string whitespace
    4. drop blank rows
    5. drop duplicate rows
    """
    if not has_cleaning_operations(
        columns_to_drop=columns_to_drop,
        rename_mapping=rename_mapping,
        strip_whitespace=strip_whitespace,
        drop_blank_rows=remove_blank_rows,
        drop_duplicates=remove_duplicates,
    ):
        raise ExcelFlowError("整形条件が指定されていません。")

    # Work on a copy so the caller's original DataFrame stays intact.
    working = dataframe.copy()
    blank_rows_removed = 0
    duplicates_removed = 0

    working = rename_columns(working, rename_mapping)
    working = drop_columns(working, columns_to_drop)

    if strip_whitespace:
        working = strip_string_columns(working)

    if remove_blank_rows:
        working, blank_rows_removed = drop_blank_rows(working)

    if remove_duplicates:
        working, duplicates_removed = drop_duplicate_rows(working)

    summary = build_cleaning_summary(
        original=dataframe,
        cleaned=working,
        blank_rows_removed=blank_rows_removed,
        duplicates_removed=duplicates_removed,
    )
    return working, summary
