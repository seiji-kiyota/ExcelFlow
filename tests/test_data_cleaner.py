"""Tests for excel_flow.data_cleaner."""

from __future__ import annotations

import pandas as pd
import pytest

from excel_flow.data_cleaner import (
    build_cleaning_summary,
    clean_dataframe,
    drop_blank_rows,
    drop_columns,
    drop_duplicate_rows,
    rename_columns,
    strip_string_columns,
    summarize_missing_values,
)
from excel_flow.validators import ExcelFlowError


def _sample_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "部署": ["営業1課", "営業2課", "営業1課"],
            "地域": [" 東京 ", "大阪", " 東京 "],
            "売上": [100, 200, 100],
            "備考": ["a", None, "a"],
        }
    )


def test_drop_one_column() -> None:
    original = _sample_frame()
    result = drop_columns(original, ["備考"])
    assert "備考" not in result.columns
    assert list(original.columns) == ["部署", "地域", "売上", "備考"]


def test_drop_multiple_columns() -> None:
    result = drop_columns(_sample_frame(), ["備考", "地域"])
    assert list(result.columns) == ["部署", "売上"]


def test_drop_columns_noop_when_unspecified() -> None:
    original = _sample_frame()
    result = drop_columns(original, [])
    assert list(result.columns) == list(original.columns)


def test_drop_all_columns_prevented() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        drop_columns(_sample_frame(), ["部署", "地域", "売上", "備考"])
    assert "すべての列を削除することはできません。" == exc_info.value.user_message


def test_rename_single_column() -> None:
    original = _sample_frame()
    result = rename_columns(original, {"部署": "部門"})
    assert "部門" in result.columns
    assert "部署" not in result.columns
    assert "部署" in original.columns


def test_rename_multiple_columns() -> None:
    result = rename_columns(_sample_frame(), {"部署": "部門", "地域": "エリア"})
    assert list(result.columns)[:2] == ["部門", "エリア"]


def test_rename_empty_name_error() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        rename_columns(_sample_frame(), {"部署": "  "})
    assert "空にすることはできません" in exc_info.value.user_message


def test_rename_duplicate_name_error() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        rename_columns(_sample_frame(), {"部署": "売上"})
    assert "重複" in exc_info.value.user_message


def test_drop_blank_rows_nan_only() -> None:
    frame = pd.DataFrame({"a": [1, None], "b": ["x", None]})
    cleaned, removed = drop_blank_rows(frame)
    assert removed == 1
    assert len(cleaned) == 1


def test_drop_blank_rows_empty_string_only() -> None:
    frame = pd.DataFrame({"a": ["x", ""], "b": ["y", ""]})
    cleaned, removed = drop_blank_rows(frame)
    assert removed == 1
    assert len(cleaned) == 1


def test_drop_blank_rows_whitespace_only() -> None:
    frame = pd.DataFrame({"a": ["x", "   "], "b": ["y", "  "]})
    cleaned, removed = drop_blank_rows(frame)
    assert removed == 1
    assert len(cleaned) == 1


def test_drop_blank_rows_keeps_normal_rows() -> None:
    frame = pd.DataFrame({"a": ["x", None], "b": ["y", "z"]})
    cleaned, removed = drop_blank_rows(frame)
    assert removed == 0
    assert len(cleaned) == 2


def test_drop_duplicate_rows() -> None:
    frame = pd.DataFrame({"a": [1, 1, 2], "b": ["x", "x", "y"]})
    cleaned, removed = drop_duplicate_rows(frame)
    assert removed == 1
    assert len(cleaned) == 2
    assert list(cleaned["a"]) == [1, 2]


def test_strip_string_columns_and_keep_numbers_nan() -> None:
    frame = pd.DataFrame(
        {
            "地域": [" 東京 ", None],
            "数量": [1, 2],
        }
    )
    original_numbers = frame["数量"].tolist()
    cleaned = strip_string_columns(frame)
    assert cleaned.loc[0, "地域"] == "東京"
    assert pd.isna(cleaned.loc[1, "地域"])
    assert cleaned["数量"].tolist() == original_numbers
    assert frame.loc[0, "地域"] == " 東京 "


def test_summarize_missing_values() -> None:
    frame = pd.DataFrame({"備考": ["a", None, None], "売上": [1, 2, 3]})
    summary = summarize_missing_values(frame)
    note_row = summary.loc[summary["列名"] == "備考"].iloc[0]
    assert note_row["欠損件数"] == 2
    assert note_row["欠損率"] == "66.7%"


def test_original_dataframe_preserved_by_clean_dataframe() -> None:
    original = _sample_frame()
    snapshot = original.copy()
    cleaned, summary = clean_dataframe(
        original,
        rename_mapping={"部署": "部門"},
        columns_to_drop=["備考"],
        strip_whitespace=True,
        remove_blank_rows=False,
        remove_duplicates=True,
    )
    pd.testing.assert_frame_equal(original, snapshot)
    assert "部門" in cleaned.columns
    assert "備考" not in cleaned.columns
    assert summary["removed_columns"] == 1


def test_cleaning_summary_counts() -> None:
    original = pd.DataFrame(
        {
            "a": [1, 1, None],
            "b": ["x", "x", "   "],
            "c": [10, 10, None],
        }
    )
    cleaned, summary = clean_dataframe(
        original,
        columns_to_drop=["c"],
        strip_whitespace=True,
        remove_blank_rows=True,
        remove_duplicates=True,
    )
    assert summary["original_rows"] == 3
    assert summary["original_columns"] == 3
    assert summary["cleaned_columns"] == 2
    assert summary["removed_columns"] == 1
    assert summary["blank_rows_removed"] == 1
    assert summary["duplicates_removed"] == 1
    assert summary["cleaned_rows"] == 1
    assert summary["removed_rows"] == 2
    assert len(cleaned) == 1


def test_clean_dataframe_requires_operations() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        clean_dataframe(_sample_frame())
    assert "整形条件が指定されていません。" == exc_info.value.user_message


def test_build_cleaning_summary_helper() -> None:
    original = pd.DataFrame({"a": [1, 2], "b": [3, 4]})
    cleaned = pd.DataFrame({"a": [1]})
    summary = build_cleaning_summary(
        original=original,
        cleaned=cleaned,
        blank_rows_removed=1,
        duplicates_removed=0,
    )
    assert summary["removed_rows"] == 1
    assert summary["removed_columns"] == 1
