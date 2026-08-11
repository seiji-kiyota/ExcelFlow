"""Tests for excel_flow.file_loader."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest

from excel_flow.file_loader import (
    detect_file_format,
    get_excel_sheet_names,
    load_csv,
    load_excel,
    load_file,
)
from excel_flow.validators import ExcelFlowError

ROOT = Path(__file__).resolve().parents[1]
SAMPLE_XLSX = ROOT / "sample_data" / "sample_sales.xlsx"
SAMPLE_CSV = ROOT / "sample_data" / "sample_sales.csv"


def test_detect_xlsx_lowercase() -> None:
    assert detect_file_format("sales.xlsx") == "xlsx"


def test_detect_csv_lowercase() -> None:
    assert detect_file_format("data.csv") == "csv"


def test_detect_xlsx_uppercase() -> None:
    assert detect_file_format("SALES.XLSX") == "xlsx"


def test_detect_csv_uppercase() -> None:
    assert detect_file_format("DATA.CSV") == "csv"


def test_detect_unsupported_extension() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        detect_file_format("notes.txt")
    assert ".xlsx" in exc_info.value.user_message
    assert ".csv" in exc_info.value.user_message


def test_get_excel_sheet_names_sample() -> None:
    with SAMPLE_XLSX.open("rb") as file_obj:
        sheets = get_excel_sheet_names(file_obj)
    assert sheets
    assert "売上データ" in sheets


def test_load_excel_sample_sales() -> None:
    with SAMPLE_XLSX.open("rb") as file_obj:
        sheets = get_excel_sheet_names(file_obj)
        dataframe = load_excel(file_obj, sheets[0])

    assert len(dataframe) > 0
    assert len(dataframe.columns) > 0
    assert "日付" in dataframe.columns
    assert "売上" in dataframe.columns


def test_load_file_excel_via_helper() -> None:
    with SAMPLE_XLSX.open("rb") as file_obj:
        sheets = get_excel_sheet_names(file_obj)
        dataframe = load_file(file_obj, "sample_sales.xlsx", sheets[0])
    assert not dataframe.empty


def test_load_utf8_csv_sample() -> None:
    with SAMPLE_CSV.open("rb") as file_obj:
        dataframe = load_csv(file_obj)
    assert len(dataframe) > 0
    assert "部署" in dataframe.columns


def test_load_cp932_csv(tmp_path: Path) -> None:
    source = pd.read_csv(SAMPLE_CSV, encoding="utf-8-sig")
    cp932_path = tmp_path / "sample_cp932.csv"
    source.to_csv(cp932_path, index=False, encoding="cp932")

    with cp932_path.open("rb") as file_obj:
        dataframe = load_csv(file_obj)

    assert len(dataframe) == len(source)
    assert list(dataframe.columns) == list(source.columns)


def test_load_csv_via_load_file() -> None:
    with SAMPLE_CSV.open("rb") as file_obj:
        dataframe = load_file(file_obj, "sample_sales.csv")
    assert not dataframe.empty


def test_load_corrupt_excel_raises() -> None:
    broken = BytesIO(b"not-an-excel-file")
    with pytest.raises(ExcelFlowError) as exc_info:
        get_excel_sheet_names(broken)
    assert "Excel" in exc_info.value.user_message


def test_load_unreadable_csv_raises() -> None:
    broken = BytesIO(b"\xff\xfe\x00\x01not,csv,content\x80\x81")
    with pytest.raises(ExcelFlowError) as exc_info:
        load_csv(broken)
    assert "CSV" in exc_info.value.user_message
