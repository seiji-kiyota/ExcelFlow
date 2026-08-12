"""Aggregate data with group-by operations."""

from __future__ import annotations

from typing import Literal, MutableMapping

import pandas as pd

from excel_flow.validators import ExcelFlowError, validate_aggregation_input

AggregationMethod = Literal["sum", "mean", "count", "max", "min"]
SortOrder = Literal["asc", "desc"]

AGGREGATION_LABELS: dict[str, str] = {
    "sum": "合計",
    "mean": "平均",
    "count": "件数",
    "max": "最大",
    "min": "最小",
}

# Streamlit widget keys / defaults for Phase 4 UI reset.
AGGREGATION_GROUP_KEY = "aggregation_group_columns"
AGGREGATION_METHOD_KEY = "aggregation_method_label"
AGGREGATION_VALUE_KEY = "aggregation_value_column"
AGGREGATION_SORT_KEY = "aggregation_sort_label"
AGGREGATION_METHOD_DEFAULT = "合計"
AGGREGATION_SORT_DEFAULT = "降順"

_PANDAS_AGG_MAP: dict[str, str] = {
    "sum": "sum",
    "mean": "mean",
    "max": "max",
    "min": "min",
}


def reset_aggregation_state(session_state: MutableMapping) -> None:
    """Clear aggregation results and restore Phase 4 UI defaults.

    Keeps loaded / cleaned DataFrames untouched.

    Call this only before Phase 4 widgets are instantiated in the current
    script run (for example from a button ``on_click`` callback, or from
    Phase 2/3 handlers that run above the aggregation widgets).
    """
    session_state.pop("aggregated_df", None)
    session_state.pop("aggregation_config", None)
    session_state[AGGREGATION_GROUP_KEY] = []
    session_state[AGGREGATION_METHOD_KEY] = AGGREGATION_METHOD_DEFAULT
    session_state[AGGREGATION_SORT_KEY] = AGGREGATION_SORT_DEFAULT
    session_state.pop(AGGREGATION_VALUE_KEY, None)


def get_numeric_columns(dataframe: pd.DataFrame) -> list[str]:
    """Return numeric column names suitable for sum/mean/max/min.

    Boolean columns are excluded.
    """
    numeric_columns: list[str] = []
    for column in dataframe.columns:
        series = dataframe[column]
        if pd.api.types.is_bool_dtype(series):
            continue
        if pd.api.types.is_numeric_dtype(series):
            numeric_columns.append(str(column))
    return numeric_columns


def result_value_column_name(aggregation: str, value_column: str | None = None) -> str:
    """Build a Japanese-friendly result column name."""
    label = AGGREGATION_LABELS.get(aggregation, aggregation)
    if aggregation == "count":
        return "件数"
    return f"{value_column}_{label}"


def aggregate_data(
    dataframe: pd.DataFrame,
    group_columns: list[str],
    aggregation: str,
    value_column: str | None = None,
) -> pd.DataFrame:
    """Aggregate ``dataframe`` without modifying the original object.

    - ``count`` returns group row counts as ``件数``
    - other methods aggregate ``value_column`` and rename results like ``売上_合計``
    """
    numeric_columns = get_numeric_columns(dataframe)
    groups, method, column = validate_aggregation_input(
        dataframe,
        group_columns=group_columns,
        aggregation=aggregation,
        value_column=value_column,
        numeric_columns=numeric_columns,
    )

    working = dataframe.copy()

    if method == "count":
        result = (
            working.groupby(groups, dropna=False, as_index=False)
            .size()
            .rename(columns={"size": "件数"})
        )
    else:
        assert column is not None
        aggregated = working.groupby(groups, dropna=False, as_index=False)[column].agg(
            _PANDAS_AGG_MAP[method]
        )
        result_name = result_value_column_name(method, column)
        result = aggregated.rename(columns={column: result_name})

    if result.empty:
        raise ExcelFlowError("集計結果が空でした。条件を確認してください。")

    return result


def sort_aggregated(
    dataframe: pd.DataFrame,
    value_column: str,
    *,
    ascending: bool = False,
) -> pd.DataFrame:
    """Sort aggregated results by ``value_column`` with a stable order.

    Group columns are used as secondary keys so ties stay deterministic.
    """
    if value_column not in dataframe.columns:
        raise ExcelFlowError(
            "並べ替え対象の列が見つかりません。",
            detail=f"missing sort column: {value_column}",
        )

    result = dataframe.copy()
    secondary = [column for column in result.columns if column != value_column]
    by_columns = [value_column, *secondary]
    ascending_flags = [ascending] + [True] * len(secondary)
    return result.sort_values(
        by=by_columns,
        ascending=ascending_flags,
        kind="mergesort",
        na_position="last",
    ).reset_index(drop=True)


def describe_aggregation(
    *,
    group_columns: list[str],
    aggregation: str,
    value_column: str | None = None,
) -> str:
    """Return a short Japanese summary of aggregation settings."""
    group_label = " × ".join(group_columns)
    method_label = AGGREGATION_LABELS.get(aggregation, aggregation)
    if aggregation == "count":
        return f"{group_label}別 / 件数"
    return f"{group_label}別 / {value_column} / {method_label}"
