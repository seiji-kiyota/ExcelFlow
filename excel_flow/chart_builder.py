"""Build Plotly charts from aggregated results."""

from __future__ import annotations

from typing import Any, MutableMapping

import pandas as pd
import plotly.express as px
from plotly.graph_objs import Figure

from excel_flow.aggregator import get_numeric_columns
from excel_flow.validators import ExcelFlowError, validate_chart_input

CHART_TYPE_KEY = "chart_type_label"
CHART_X_KEY = "chart_x_column"
CHART_Y_KEY = "chart_y_column"
CHART_TITLE_KEY = "chart_title"
CHART_TYPE_DEFAULT = "棒グラフ"

CHART_TYPE_LABELS: dict[str, str] = {
    "bar": "棒グラフ",
    "line": "折れ線グラフ",
    "pie": "円グラフ",
}

CHART_TYPE_OPTIONS: dict[str, str] = {
    "棒グラフ": "bar",
    "折れ線グラフ": "line",
    "円グラフ": "pie",
}


def reset_chart_state(session_state: MutableMapping) -> None:
    """Clear Phase 5 chart settings/widgets without touching aggregation data.

    Call before Phase 5 widgets are instantiated (e.g. button ``on_click``).
    """
    for key in (
        "chart_config",
        "chart_generated",
        CHART_TYPE_KEY,
        CHART_X_KEY,
        CHART_Y_KEY,
        CHART_TITLE_KEY,
    ):
        session_state.pop(key, None)


def build_default_chart_title(
    group_columns: list[str],
    aggregation: str,
    value_column: str | None = None,
) -> str:
    """Build a user-friendly default chart title from aggregation settings."""
    from excel_flow.aggregator import AGGREGATION_LABELS

    group_label = "×".join(group_columns)
    method_label = AGGREGATION_LABELS.get(aggregation, aggregation)
    if aggregation == "count":
        return f"{group_label}別 件数"
    return f"{group_label}別 {value_column} {method_label}"


def build_chart(
    dataframe: pd.DataFrame,
    chart_type: str,
    x_column: str,
    y_column: str,
    title: str = "",
    color_column: str | None = None,
) -> Figure:
    """Build a Plotly figure from an aggregated DataFrame.

    The input DataFrame is not modified and row order is preserved.
    """
    numeric_columns = get_numeric_columns(dataframe)
    method, x_name, y_name, color_name = validate_chart_input(
        dataframe,
        chart_type=chart_type,
        x_column=x_column,
        y_column=y_column,
        color_column=color_column,
        numeric_columns=numeric_columns,
    )

    working = dataframe.copy()
    chart_title = "" if title is None else str(title)

    try:
        if method == "bar":
            figure = px.bar(
                working,
                x=x_name,
                y=y_name,
                color=color_name,
                title=chart_title or None,
            )
        elif method == "line":
            figure = px.line(
                working,
                x=x_name,
                y=y_name,
                color=color_name,
                markers=True,
                title=chart_title or None,
            )
            # Slightly larger markers for readability; keep line width/color unchanged.
            figure.update_traces(marker={"size": 14})
        else:
            figure = px.pie(
                working,
                names=x_name,
                values=y_name,
                title=chart_title or None,
            )
            # Enlarge in-slice percentage labels only.
            figure.update_traces(textfont_size=22)
    except ExcelFlowError:
        raise
    except Exception as exc:  # noqa: BLE001 - convert to user-facing error
        raise ExcelFlowError(
            "グラフを作成できませんでした。集計結果と設定を確認してください。",
            detail=str(exc),
        ) from exc

    figure.update_layout(
        margin={"l": 40, "r": 40, "t": 60, "b": 40},
        height=420,
        legend_title_text="",
    )
    return figure


def chart_settings_from_aggregation(
    aggregation_config: dict[str, Any],
) -> dict[str, Any]:
    """Derive sensible default chart settings from Phase 4 aggregation config."""
    group_columns = list(aggregation_config.get("group_columns") or [])
    if not group_columns:
        raise ExcelFlowError("グラフ化できるグループ項目がありません。")

    result_column = str(aggregation_config["result_column"])
    color_column = group_columns[1] if len(group_columns) >= 2 else None
    title = build_default_chart_title(
        group_columns=group_columns,
        aggregation=str(aggregation_config.get("aggregation", "")),
        value_column=aggregation_config.get("value_column"),
    )
    return {
        "chart_type": "bar",
        "x_column": group_columns[0],
        "y_column": result_column,
        "color_column": color_column,
        "title": title,
        "group_columns": group_columns,
    }
