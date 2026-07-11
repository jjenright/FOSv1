from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from src.extract import LegacyLayoutExtractor
from src.transform import CategoryRegistry


CONFIG = Path(__file__).resolve().parents[1] / "config" / "categories.yaml"


def build_legacy_sheet():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "2010"
    worksheet["A1"] = "Pay Period"
    worksheet["B1"] = "Jan 1 - Jan 14"
    worksheet["A2"] = "Pay"
    worksheet["B2"] = 2000
    worksheet["A4"] = "Bi-Weekly Deductions"
    worksheet["B5"] = 100
    worksheet["C5"] = "Gas"
    worksheet["B6"] = 25
    worksheet["C6"] = "Mystery merchant"
    worksheet["A7"] = "Total Spent During Period"
    worksheet["B7"] = 125
    worksheet["C7"] = "$800 target"
    return workbook, worksheet


def test_extracts_legacy_income_expense_and_unknown() -> None:
    workbook, worksheet = build_legacy_sheet()
    result = LegacyLayoutExtractor(CategoryRegistry(CONFIG)).extract_worksheet(worksheet)

    assert result.periods == ("Jan 1 - Jan 14",)
    assert len(result.source_rows) == 3
    assert len(result.records) == 2
    assert len(result.income) == 1
    assert len(result.variable_expenses) == 1
    assert result.income[0].amount == Decimal("2000")
    assert result.variable_expenses[0].category_id == "TRA002"
    assert len(result.unknown_categories) == 1
    assert result.unknown_categories[0].source_cell == "C6"
    workbook.close()


def test_duplicate_legacy_period_labels_are_made_unique() -> None:
    workbook, worksheet = build_legacy_sheet()
    worksheet["A10"] = "Pay Period"
    worksheet["B10"] = "Jan 1 - Jan 14"
    worksheet["A11"] = "Pay"
    worksheet["B11"] = 1000

    result = LegacyLayoutExtractor(CategoryRegistry(CONFIG)).extract_worksheet(worksheet)

    assert result.periods == ("Jan 1 - Jan 14", "Jan 1 - Jan 14 [2]")
    workbook.close()
