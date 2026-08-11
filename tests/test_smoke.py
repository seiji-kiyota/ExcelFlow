"""Phase 1 smoke tests for project structure and imports."""

from pathlib import Path

import excel_flow
import excel_flow.aggregator
import excel_flow.chart_builder
import excel_flow.data_cleaner
import excel_flow.excel_exporter
import excel_flow.file_loader
import excel_flow.validators

ROOT = Path(__file__).resolve().parents[1]


def test_excel_flow_package_importable() -> None:
    assert excel_flow.__version__


def test_modules_importable() -> None:
    assert excel_flow.file_loader.__doc__
    assert excel_flow.data_cleaner.__doc__
    assert excel_flow.aggregator.__doc__
    assert excel_flow.chart_builder.__doc__
    assert excel_flow.excel_exporter.__doc__
    assert excel_flow.validators.__doc__


def test_app_py_exists() -> None:
    assert (ROOT / "app.py").is_file()


def test_sample_sales_xlsx_exists() -> None:
    sample = ROOT / "sample_data" / "sample_sales.xlsx"
    assert sample.is_file()
    assert sample.stat().st_size > 0
