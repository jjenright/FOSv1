from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.pipeline import CurrentYearPipeline
from tests.test_current_extractor import build_current_sheet


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_runs_end_to_end_current_year_pipeline(tmp_path) -> None:
    source_workbook, _ = build_current_sheet()
    source = tmp_path / "Budget.xlsx"
    output = tmp_path / "Financial_Operating_System.xlsx"
    source_workbook.save(source)
    source_workbook.close()

    result = CurrentYearPipeline(PROJECT_ROOT).run(
        source,
        sheet_name="2025",
        output_path=output,
        fos_version="test",
    )

    assert output.is_file()
    assert result.validation.is_valid
    assert result.load_result.transaction_rows == 3
    assert result.load_result.income_rows == 1
    assert result.load_result.exception_rows == 1
    assert result.validation_summary_path.is_file()
    assert result.exceptions_path.is_file()

    workbook = load_workbook(output, data_only=False)
    try:
        assert workbook["Import_Log"]["E2"].value == "PASS"
        assert workbook["Dashboard"]["B9"].value == "=SUM(FactIncome!F:F)"
    finally:
        workbook.close()


def test_rejects_non_current_layout_before_extraction(tmp_path) -> None:
    source_workbook, worksheet = build_current_sheet()
    worksheet.title = "2010"
    source = tmp_path / "Budget.xlsx"
    source_workbook.save(source)
    source_workbook.close()

    with pytest.raises(ValueError, match="supports current-layout imports only"):
        CurrentYearPipeline(PROJECT_ROOT).run(source, sheet_name="2010")


def test_rejects_missing_workbook(tmp_path) -> None:
    with pytest.raises(FileNotFoundError, match="Workbook not found"):
        CurrentYearPipeline(PROJECT_ROOT).run(tmp_path / "missing.xlsx")
