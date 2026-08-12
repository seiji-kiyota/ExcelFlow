"""Validate input files, data, and user settings for ExcelFlow."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

SUPPORTED_EXTENSIONS = {".xlsx", ".csv"}


class ExcelFlowError(Exception):
    """Application error with a user-facing message.

    ``user_message`` is safe to show in the UI.
    ``detail`` is optional developer-oriented context and should not be shown
    as the primary error text.
    """

    def __init__(self, user_message: str, detail: str | None = None) -> None:
        self.user_message = user_message
        self.detail = detail
        super().__init__(user_message)


def get_extension(filename: str | None) -> str:
    """Return the lower-case file extension including the leading dot."""
    if filename is None or not str(filename).strip():
        raise ExcelFlowError("ファイルが選択されていません。")

    extension = Path(str(filename)).suffix.lower()
    if not extension:
        raise ExcelFlowError(
            "対応しているファイル形式は .xlsx と .csv です。",
            detail=f"filename without extension: {filename!r}",
        )
    return extension


def validate_supported_extension(filename: str | None) -> str:
    """Validate that the file extension is supported.

    Returns
    -------
    str
        Normalized extension (``.xlsx`` or ``.csv``).
    """
    extension = get_extension(filename)
    if extension not in SUPPORTED_EXTENSIONS:
        raise ExcelFlowError(
            "対応しているファイル形式は .xlsx と .csv です。",
            detail=f"unsupported extension: {extension}",
        )
    return extension


def validate_sheet_names(sheet_names: list[str]) -> list[str]:
    """Ensure the workbook has at least one usable sheet name."""
    cleaned = [name for name in sheet_names if str(name).strip()]
    if not cleaned:
        raise ExcelFlowError("Excelファイルに有効なシートがありません。")
    return cleaned


def validate_dataframe_not_empty(df: pd.DataFrame) -> None:
    """Raise when the loaded DataFrame has no rows and no columns."""
    if df is None or (df.empty and len(df.columns) == 0):
        raise ExcelFlowError("読み込んだファイルにデータがありません。")
    if df.empty:
        raise ExcelFlowError("読み込んだファイルにデータがありません。")


def validate_columns_to_drop(
    dataframe: pd.DataFrame,
    columns_to_drop: list[str] | None,
) -> list[str]:
    """Validate and normalize columns selected for deletion.

    Unknown column names are ignored safely.
    Deleting every column is rejected.
    """
    if not columns_to_drop:
        return []

    existing = set(dataframe.columns.astype(str))
    normalized = [str(column) for column in columns_to_drop if str(column) in existing]
    # Preserve order while removing duplicates.
    unique_columns = list(dict.fromkeys(normalized))

    if unique_columns and len(unique_columns) >= len(dataframe.columns):
        raise ExcelFlowError("すべての列を削除することはできません。")

    return unique_columns


def validate_rename_mapping(
    dataframe: pd.DataFrame,
    rename_mapping: dict[str, str] | None,
) -> dict[str, str]:
    """Validate column rename mapping.

    Rules:
    - empty new names are rejected
    - source columns must exist
    - new names must not collide with remaining columns or each other
    """
    if not rename_mapping:
        return {}

    current_columns = [str(column) for column in dataframe.columns]
    current_set = set(current_columns)
    cleaned: dict[str, str] = {}

    for source, target in rename_mapping.items():
        source_name = str(source)
        target_name = "" if target is None else str(target).strip()

        if source_name not in current_set:
            raise ExcelFlowError(
                "指定された列が見つかりません。",
                detail=f"missing source column: {source_name}",
            )
        if not target_name:
            raise ExcelFlowError("変更後の列名を空にすることはできません。")
        if source_name == target_name:
            continue
        cleaned[source_name] = target_name

    if not cleaned:
        return {}

    reserved = current_set - set(cleaned.keys())
    new_names = list(cleaned.values())
    if len(new_names) != len(set(new_names)):
        raise ExcelFlowError("変更後の列名が重複しています。")

    for new_name in new_names:
        if new_name in reserved:
            raise ExcelFlowError(
                "変更後の列名が重複しています。",
                detail=f"conflicts with existing column: {new_name}",
            )

    return cleaned


def has_cleaning_operations(
    *,
    columns_to_drop: list[str] | None = None,
    rename_mapping: dict[str, str] | None = None,
    strip_whitespace: bool = False,
    drop_blank_rows: bool = False,
    drop_duplicates: bool = False,
) -> bool:
    """Return True when at least one cleaning operation is requested."""
    return bool(
        columns_to_drop
        or rename_mapping
        or strip_whitespace
        or drop_blank_rows
        or drop_duplicates
    )


SUPPORTED_AGGREGATIONS = ("sum", "mean", "count", "max", "min")


def validate_group_columns(
    dataframe: pd.DataFrame,
    group_columns: list[str] | None,
) -> list[str]:
    """Validate group-by columns (1 or 2 unique existing columns)."""
    if not group_columns:
        raise ExcelFlowError("グループ項目を1つ以上選択してください。")

    normalized = [str(column) for column in group_columns]
    if len(normalized) != len(set(normalized)):
        raise ExcelFlowError("グループ項目に同じ列を重複して指定できません。")
    if len(normalized) > 2:
        raise ExcelFlowError("グループ項目は最大2列まで選択できます。")

    existing = set(dataframe.columns.astype(str))
    for column in normalized:
        if column not in existing:
            raise ExcelFlowError(
                "指定された列が見つかりません。",
                detail=f"missing group column: {column}",
            )
    return normalized


def validate_aggregation_method(aggregation: str | None) -> str:
    """Validate aggregation method identifier."""
    if aggregation is None or not str(aggregation).strip():
        raise ExcelFlowError("集計方法を選択してください。")
    method = str(aggregation).strip().lower()
    if method not in SUPPORTED_AGGREGATIONS:
        raise ExcelFlowError(
            "対応している集計方法は 合計 / 平均 / 件数 / 最大 / 最小 です。",
            detail=f"unsupported aggregation: {method}",
        )
    return method


def validate_value_column(
    dataframe: pd.DataFrame,
    value_column: str | None,
    *,
    numeric_columns: list[str],
) -> str:
    """Validate a numeric value column used by sum/mean/max/min."""
    if value_column is None or not str(value_column).strip():
        raise ExcelFlowError("集計対象列を選択してください。")

    column = str(value_column)
    existing = set(dataframe.columns.astype(str))
    if column not in existing:
        raise ExcelFlowError(
            "指定された列が見つかりません。",
            detail=f"missing value column: {column}",
        )
    if column not in numeric_columns:
        raise ExcelFlowError(
            "合計・平均・最大・最小には数値列のみ指定できます。",
            detail=f"non-numeric column: {column}",
        )
    return column


def validate_aggregation_input(
    dataframe: pd.DataFrame,
    *,
    group_columns: list[str] | None,
    aggregation: str | None,
    value_column: str | None,
    numeric_columns: list[str],
) -> tuple[list[str], str, str | None]:
    """Validate full aggregation settings and return normalized values."""
    if dataframe is None or dataframe.empty:
        raise ExcelFlowError("集計できるデータがありません。")

    groups = validate_group_columns(dataframe, group_columns)
    method = validate_aggregation_method(aggregation)

    if method == "count":
        return groups, method, None

    column = validate_value_column(
        dataframe,
        value_column,
        numeric_columns=numeric_columns,
    )
    return groups, method, column
