"""Tests for excel_flow.validators."""

from __future__ import annotations

import pandas as pd
import pytest

from excel_flow.validators import (
    ExcelFlowError,
    get_extension,
    validate_dataframe_not_empty,
    validate_sheet_names,
    validate_supported_extension,
)


def test_get_extension_normalizes_case() -> None:
    assert get_extension("report.XLSX") == ".xlsx"
    assert get_extension("report.Csv") == ".csv"


def test_validate_supported_extension_ok() -> None:
    assert validate_supported_extension("a.xlsx") == ".xlsx"
    assert validate_supported_extension("b.CSV") == ".csv"


def test_validate_supported_extension_rejects_xls() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        validate_supported_extension("legacy.xls")
    assert "対応しているファイル形式は .xlsx と .csv です。" == exc_info.value.user_message


def test_validate_missing_filename() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        validate_supported_extension(None)
    assert "ファイルが選択されていません。" == exc_info.value.user_message


def test_validate_sheet_names_empty() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        validate_sheet_names([])
    assert "有効なシート" in exc_info.value.user_message


def test_validate_sheet_names_ok() -> None:
    assert validate_sheet_names(["Sheet1", "売上"]) == ["Sheet1", "売上"]


def test_validate_empty_dataframe() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        validate_dataframe_not_empty(pd.DataFrame())
    assert "データがありません" in exc_info.value.user_message


def test_validate_non_empty_dataframe() -> None:
    validate_dataframe_not_empty(pd.DataFrame({"a": [1]}))
