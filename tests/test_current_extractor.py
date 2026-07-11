from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from src.extract import CurrentLayoutExtractor
from src.transform import CategoryRegistry


CONFIG = Path(__file__).resolve().parents[1] / "config" / "categories.yaml"


def build_current_sheet():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "2025"
    worksheet["A1"] = "Jan 1 - Jan 14"
    worksheet["B1"] = "PayAmount"
    worksheet["C1"] = "Bi-weekly"
    worksheet["D1"] = "BWAmount"
    worksheet["E1"] = "Monthlies"
    worksheet["F1"] = "MonAmount"

    worksheet["A2"] = "Prev balance"
    worksheet["B2"] = 1000
    worksheet["A3"] = "JE's Pay"
    worksheet["B3"] = 4000
    worksheet["C2"] = "Savings"
    worksheet["D2"] = 500
    worksheet["C3"] = "Costco"
    worksheet["D3"] = 200
    worksheet["C4"] = "Mystery merchant"
    worksheet["D4"] = 25
    worksheet["C5"] = "Fuel"
    worksheet["D5"] = 0
    worksheet["E2"] = "Mortgage (21st)"
    worksheet["F2"] = 2000
    return workbook, worksheet


def test_extracts_and_separates_current_layout_records() -> None:
    workbook, worksheet = build_current_sheet()
    extractor = CurrentLayoutExtractor(CategoryRegistry(CONFIG))

    result = extractor.extract_worksheet(worksheet)

    assert result.periods == ("Jan 1 - Jan 14",)
    assert len(result.records) == 4
    assert len(result.source_rows) == 5
    assert len(result.income) == 1
    assert len(result.transfers) == 1
    assert len(result.variable_expenses) == 1
    assert len(result.fixed_expenses) == 1
    assert result.income[0].amount == Decimal("4000")
    assert result.variable_expenses[0].category_id == "FOD001"
    assert result.fixed_expenses[0].source_cell == "E2"
    assert len(result.unknown_categories) == 1
    assert result.unknown_categories[0].original_name == "Mystery merchant"
    workbook.close()


def test_zero_values_are_skipped_by_default_and_optional() -> None:
    workbook, worksheet = build_current_sheet()
    extractor = CurrentLayoutExtractor(CategoryRegistry(CONFIG))

    default_result = extractor.extract_worksheet(worksheet)
    with_zero_result = extractor.extract_worksheet(worksheet, include_zero=True)

    assert all(transaction.amount != 0 for transaction in default_result.variable_expenses)
    assert any(transaction.amount == 0 for transaction in with_zero_result.variable_expenses)
    workbook.close()


def test_rejects_non_current_layout() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "2025"
    extractor = CurrentLayoutExtractor(CategoryRegistry(CONFIG))

    with pytest.raises(ValueError, match="does not match the current layout"):
        extractor.extract_worksheet(worksheet)
    workbook.close()


def test_rejects_sheet_name_without_year() -> None:
    workbook, worksheet = build_current_sheet()
    worksheet.title = "Current"
    extractor = CurrentLayoutExtractor(CategoryRegistry(CONFIG))

    with pytest.raises(ValueError, match="does not begin with a year"):
        extractor.extract_worksheet(worksheet)
    workbook.close()
