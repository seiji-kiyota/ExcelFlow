"""Tests for excel_flow.chart_builder."""

from __future__ import annotations

import pandas as pd
import pytest

from excel_flow.chart_builder import (
    CHART_TITLE_KEY,
    CHART_TYPE_KEY,
    CHART_X_KEY,
    CHART_Y_KEY,
    build_chart,
    build_default_chart_title,
    reset_chart_state,
)
from excel_flow.validators import ExcelFlowError


def _one_group() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "部署": ["営業2課", "営業1課", "営業3課"],
            "売上_合計": [300, 200, 100],
        }
    )


def _two_group() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "部署": ["営業1課", "営業1課", "営業2課"],
            "商品": ["ノートPC", "モニター", "ノートPC"],
            "件数": [2, 1, 3],
        }
    )


def test_bar_chart_one_group() -> None:
    original = _one_group()
    snapshot = original.copy()
    figure = build_chart(original, "bar", "部署", "売上_合計", title="部署別 売上 合計")
    pd.testing.assert_frame_equal(original, snapshot)
    assert figure.layout.title.text == "部署別 売上 合計"
    assert figure.data


def test_bar_chart_two_group_with_color() -> None:
    figure = build_chart(
        _two_group(),
        "bar",
        "部署",
        "件数",
        title="部署×商品別 件数",
        color_column="商品",
    )
    assert len(figure.data) >= 1


def test_line_chart_one_and_two_group() -> None:
    one = build_chart(_one_group(), "line", "部署", "売上_合計", title="line")
    two = build_chart(
        _two_group(),
        "line",
        "部署",
        "件数",
        title="line2",
        color_column="商品",
    )
    assert one.data
    assert two.data


def test_pie_chart_one_group() -> None:
    figure = build_chart(_one_group(), "pie", "部署", "売上_合計", title="円")
    assert figure.layout.title.text == "円"
    assert figure.data


def test_pie_chart_rejects_two_group_color() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        build_chart(
            _two_group(),
            "pie",
            "部署",
            "件数",
            color_column="商品",
        )
    assert "円グラフは1つのグループ項目" in exc_info.value.user_message


def test_empty_dataframe_rejected() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        build_chart(pd.DataFrame(columns=["部署", "件数"]), "bar", "部署", "件数")
    assert "グラフ化できる集計結果がありません" in exc_info.value.user_message


def test_missing_x_column() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        build_chart(_one_group(), "bar", "存在しない", "売上_合計")
    assert "見つかりません" in exc_info.value.user_message


def test_missing_y_column() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        build_chart(_one_group(), "bar", "部署", "存在しない")
    assert "見つかりません" in exc_info.value.user_message


def test_non_numeric_y_rejected() -> None:
    frame = pd.DataFrame({"部署": ["A"], "備考": ["x"]})
    with pytest.raises(ExcelFlowError) as exc_info:
        build_chart(frame, "bar", "部署", "備考")
    assert "数値列のみ" in exc_info.value.user_message


def test_unsupported_chart_type() -> None:
    with pytest.raises(ExcelFlowError) as exc_info:
        build_chart(_one_group(), "scatter", "部署", "売上_合計")
    assert "グラフ種類" in exc_info.value.user_message


def test_title_can_be_empty() -> None:
    figure = build_chart(_one_group(), "bar", "部署", "売上_合計", title="")
    assert figure.data


def test_input_order_preserved_in_bar_chart() -> None:
    frame = pd.DataFrame({"部署": ["C", "A", "B"], "件数": [3, 1, 2]})
    figure = build_chart(frame, "bar", "部署", "件数", title="order")
    assert list(figure.data[0].x) == ["C", "A", "B"]
    assert list(figure.data[0].y) == [3, 1, 2]


def test_default_chart_title() -> None:
    assert build_default_chart_title(["部署"], "sum", "売上") == "部署別 売上 合計"
    assert build_default_chart_title(["部署", "商品"], "count") == "部署×商品別 件数"


def test_reset_chart_state() -> None:
    state = {
        "aggregated_df": _one_group(),
        "chart_generated": True,
        "chart_config": {"chart_type": "bar"},
        CHART_TYPE_KEY: "円グラフ",
        CHART_X_KEY: "部署",
        CHART_Y_KEY: "売上_合計",
        CHART_TITLE_KEY: "title",
    }
    reset_chart_state(state)
    assert "chart_generated" not in state
    assert "chart_config" not in state
    assert CHART_TYPE_KEY not in state
    assert "aggregated_df" in state
