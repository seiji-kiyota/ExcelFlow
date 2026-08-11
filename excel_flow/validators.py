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
