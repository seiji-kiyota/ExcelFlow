"""Tests for excel_flow.aggregator."""

from __future__ import annotations

import pandas as pd
import pytest

from excel_flow.aggregator import (
    AGGREGATION_GROUP_KEY,
    AGGREGATION_METHOD_DEFAULT,
    AGGREGATION_METHOD_KEY,
    AGGREGATION_SORT_DEFAULT,
    AGGREGATION_SORT_KEY,
    AGGREGATION_VALUE_KEY,
    aggregate_data,
    describe_aggregation,
    get_numeric_columns,
    reset_aggregation_state,
    result_value_column_name,
    sort_aggregated,
)
from excel_flow.validators import ExcelFlowError


def _sample_sales() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "部署": ["営業1課", "営業1課", "営業2課", "営業2課", "営業3課"],
            "商品": ["ノートPC", "モニター", "ノートPC", "モニター", "ノートPC"],
            "地域": ["東京", "大阪", "東京", "大阪", "名古屋"],
            "数量": [2, 5, 1, 3, 4],
            "売上": [200, 100, 150, 90, 80],
            "備考": ["a", "b", "c", "d", "e"],
        }
    )


def test_get_numeric_columns_excludes_bool_and_object() -> None:
    frame = pd.DataFrame(
        {
            "数量": [1, 2],
            "フラグ": [True, False],
            "部署": ["A", "B"],
        }
    )
    assert get_numeric_columns(frame) == ["数量"]


def test_sum_one_group() -> None:
    original = _sample_sales()
    snapshot = original.copy()
    result = aggregate_data(original, ["部署"], "sum", "売上")
    pd.testing.assert_frame_equal(original, snapshot)
    assert list(result.columns) == ["部署", "売上_合計"]
    mapped = dict(zip(result["部署"], result["売上_合計"]))
    assert mapped["営業1課"] == 300
    assert mapped["営業2課"] == 240
    assert mapped["営業3課"] == 80


def test_sum_two_groups() -> None:
    result = aggregate_data(_sample_sales(), ["部署", "商品"], "sum", "売上")
    assert list(result.columns) == ["部署", "商品", "売上_合計"]
    assert len(result) == 5


def test_mean_aggregation() -> None:
    result = aggregate_data(_sample_sales(), ["地域"], "mean", "数量")
    assert "数量_平均" in result.columns
    tokyo = result.loc[result["地域"] == "東京", "数量_平均"].iloc[0]
    assert tokyo == pytest.approx(1.5)


def test_count_one_group() -> None:
    result = aggregate_data(_sample_sales(), ["部署"], "count")
    assert list(result.columns) == ["部署", "件数"]
    mapped = dict(zip(result["部署"], result["件数"]))
    assert mapped["営業1課"] == 2
    assert mapped["営業3課"] == 1


def test_count_two_groups() -> None:
    result = aggregate_data(_sample_sales(), ["部署", "商品"], "count")
    assert "件数" in result.columns
    assert len(result) == 5


def test_max_aggregation() -> None:
    result = aggregate_data(_sample_sales(), ["部署"], "max", "売上")
    mapped = dict(zip(result["部署"], result["売上_最大"]))
    assert mapped["営業1課"] == 200


def test_min_aggregation() -> None:
    result = aggregate_data(_sample_sales(), ["部署"], "min", "売上")
    mapped = dict(zip(result["部署"], result["売上_最小"]))
    assert mapped["営業1課"] == 100


def test_sort_descending_and_ascending() -> None:
    result = aggregate_data(_sample_sales(), ["部署"], "sum", "売上")
    desc = sort_aggregated(result, "売上_合計", ascending=False)
    asc = sort_aggregated(result, "売上_合計", ascending=True)
    assert list(desc["売上_合計"]) == sorted(result["売上_合計"], reverse=True)
    assert list(asc["売上_合計"]) == sorted(result["売上_合計"])


def test_sort_is_stable_for_ties() -> None:
    frame = pd.DataFrame({"部署": ["B", "A", "C"], "件数": [1, 1, 1]})
    sorted_frame = sort_aggregated(frame, "件数", ascending=False)
    assert list(sorted_frame["部署"]) == ["A", "B", "C"]


def test_missing_group_columns() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        aggregate_data(_sample_sales(), [], "sum", "売上")
    assert "グループ項目" in exc_info.value.user_message


def test_too_many_group_columns() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        aggregate_data(_sample_sales(), ["部署", "商品", "地域"], "count")
    assert "最大2列" in exc_info.value.user_message


def test_missing_column() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        aggregate_data(_sample_sales(), ["存在しない"], "count")
    assert "見つかりません" in exc_info.value.user_message


def test_non_numeric_sum_rejected() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        aggregate_data(_sample_sales(), ["部署"], "sum", "備考")
    assert "数値列のみ" in exc_info.value.user_message


def test_empty_dataframe_rejected() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        aggregate_data(pd.DataFrame(columns=["部署", "売上"]), ["部署"], "sum", "売上")
    assert "集計できるデータがありません" in exc_info.value.user_message


def test_nan_in_numeric_column_is_ignored_by_sum() -> None:
    frame = pd.DataFrame(
        {
            "部署": ["営業1課", "営業1課", "営業2課"],
            "売上": [100.0, None, 50.0],
        }
    )
    result = aggregate_data(frame, ["部署"], "sum", "売上")
    mapped = dict(zip(result["部署"], result["売上_合計"]))
    assert mapped["営業1課"] == 100.0
    assert mapped["営業2課"] == 50.0


def test_result_column_names_and_description() -> None:
    assert result_value_column_name("sum", "売上") == "売上_合計"
    assert result_value_column_name("count") == "件数"
    assert describe_aggregation(
        group_columns=["部署", "商品"],
        aggregation="count",
    ) == "部署 × 商品別 / 件数"


def test_reset_aggregation_state_restores_defaults() -> None:
    original = _sample_sales()
    cleaned = original.copy()
    state = {
        "original_df": original,
        "cleaned_df": cleaned,
        "aggregated_df": aggregate_data(original, ["部署"], "sum", "売上"),
        "aggregation_config": {"description": "部署別 / 売上 / 合計"},
        AGGREGATION_GROUP_KEY: ["部署", "商品"],
        AGGREGATION_METHOD_KEY: "件数",
        AGGREGATION_VALUE_KEY: "数量",
        AGGREGATION_SORT_KEY: "昇順",
    }

    reset_aggregation_state(state)

    assert "aggregated_df" not in state
    assert "aggregation_config" not in state
    assert state[AGGREGATION_GROUP_KEY] == []
    assert state[AGGREGATION_METHOD_KEY] == AGGREGATION_METHOD_DEFAULT
    assert state[AGGREGATION_SORT_KEY] == AGGREGATION_SORT_DEFAULT
    assert AGGREGATION_VALUE_KEY not in state
    assert state["original_df"] is original
    assert state["cleaned_df"] is cleaned


def test_reset_aggregation_state_is_idempotent() -> None:
    state = {
        AGGREGATION_GROUP_KEY: [],
        AGGREGATION_METHOD_KEY: AGGREGATION_METHOD_DEFAULT,
        AGGREGATION_SORT_KEY: AGGREGATION_SORT_DEFAULT,
    }
    reset_aggregation_state(state)
    reset_aggregation_state(state)
    assert state[AGGREGATION_GROUP_KEY] == []
    assert state[AGGREGATION_METHOD_KEY] == AGGREGATION_METHOD_DEFAULT
    assert state[AGGREGATION_SORT_KEY] == AGGREGATION_SORT_DEFAULT
