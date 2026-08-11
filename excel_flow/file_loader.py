"""Load Excel and CSV files into pandas DataFrames."""

from __future__ import annotations

from typing import BinaryIO, Literal

import pandas as pd

from excel_flow.validators import (
    ExcelFlowError,
    validate_dataframe_not_empty,
    validate_sheet_names,
    validate_supported_extension,
)

FileFormat = Literal["xlsx", "csv"]
CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp932")


def detect_file_format(filename: str | None) -> FileFormat:
    """Detect whether a file is ``xlsx`` or ``csv`` from its name."""
    extension = validate_supported_extension(filename)
    if extension == ".xlsx":
        return "xlsx"
    return "csv"


def _ensure_seekable(file_obj: BinaryIO) -> None:
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)


def get_excel_sheet_names(file_obj: BinaryIO) -> list[str]:
    """Return sheet names from an Excel workbook (.xlsx)."""
    _ensure_seekable(file_obj)
    try:
        with pd.ExcelFile(file_obj, engine="openpyxl") as workbook:
            sheet_names = list(workbook.sheet_names)
    except ExcelFlowError:
        raise
    except Exception as exc:  # noqa: BLE001 - convert to user-facing error
        raise ExcelFlowError(
            "Excelファイルを読み込めませんでした。ファイルが破損していないか確認してください。",
            detail=str(exc),
        ) from exc
    finally:
        _ensure_seekable(file_obj)

    return validate_sheet_names(sheet_names)


def load_excel(file_obj: BinaryIO, sheet_name: str) -> pd.DataFrame:
    """Load one sheet from an Excel workbook into a DataFrame."""
    if not str(sheet_name).strip():
        raise ExcelFlowError("読み込むシートを選択してください。")

    _ensure_seekable(file_obj)
    try:
        dataframe = pd.read_excel(file_obj, sheet_name=sheet_name, engine="openpyxl")
    except ExcelFlowError:
        raise
    except Exception as exc:  # noqa: BLE001 - convert to user-facing error
        raise ExcelFlowError(
            "Excelファイルを読み込めませんでした。ファイルが破損していないか確認してください。",
            detail=str(exc),
        ) from exc
    finally:
        _ensure_seekable(file_obj)

    validate_dataframe_not_empty(dataframe)
    return dataframe


def load_csv(file_obj: BinaryIO) -> pd.DataFrame:
    """Load a CSV file, trying UTF-8 family encodings then CP932."""
    last_error: Exception | None = None

    for encoding in CSV_ENCODINGS:
        _ensure_seekable(file_obj)
        try:
            dataframe = pd.read_csv(file_obj, encoding=encoding)
            validate_dataframe_not_empty(dataframe)
            _ensure_seekable(file_obj)
            return dataframe
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
        except ExcelFlowError:
            raise
        except Exception as exc:  # noqa: BLE001 - convert to user-facing error
            last_error = exc
            # Encoding mismatch can also surface as a generic parser error.
            if encoding != CSV_ENCODINGS[-1]:
                continue
            raise ExcelFlowError(
                "CSVファイルを読み込めませんでした。文字コードまたはファイル内容を確認してください。",
                detail=str(exc),
            ) from exc

    raise ExcelFlowError(
        "CSVファイルを読み込めませんでした。文字コードまたはファイル内容を確認してください。",
        detail=str(last_error) if last_error else None,
    )


def load_file(
    file_obj: BinaryIO,
    filename: str,
    sheet_name: str | None = None,
) -> pd.DataFrame:
    """Load an uploaded Excel or CSV file into a DataFrame."""
    file_format = detect_file_format(filename)
    if file_format == "xlsx":
        if sheet_name is None or not str(sheet_name).strip():
            raise ExcelFlowError("読み込むシートを選択してください。")
        return load_excel(file_obj, sheet_name)
    return load_csv(file_obj)


def summarize_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    """Return a simple column-name / dtype summary table for UI display."""
    return pd.DataFrame(
        {
            "列名": dataframe.columns.astype(str),
            "データ型": [str(dtype) for dtype in dataframe.dtypes],
        }
    )
