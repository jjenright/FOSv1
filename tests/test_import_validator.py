from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook

from src.extract import (
    CurrentExtractionResult,
    CurrentLayoutExtractor,
    ExtractedRecord,
    SourceRow,
)
from src.models import Transaction
from src.transform import CategoryRegistry
from src.validate import ImportValidator, write_validation_report

CONFIG = Path(__file__).resolve().parents[1] / "config" / "categories.yaml"


def build_result_with_unknown() -> CurrentExtractionResult:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "2025"
    worksheet["A1"] = "Jan 1 - Jan 14"
    worksheet["B1"] = "PayAmount"
    worksheet["C1"] = "Bi-weekly"
    worksheet["D1"] = "BWAmount"
    worksheet["E1"] = "Monthlies"
    worksheet["F1"] = "MonAmount"
    worksheet["A2"] = "JE's Pay"
    worksheet["B2"] = 4000
    worksheet["C2"] = "Costco"
    worksheet["D2"] = 200
    worksheet["C3"] = "Mystery merchant"
    worksheet["D3"] = 25
    worksheet["E2"] = "Mortgage (21st)"
    worksheet["F2"] = 2000
    result = CurrentLayoutExtractor(CategoryRegistry(CONFIG)).extract_worksheet(worksheet)
    workbook.close()
    return result


def test_validates_reconciliation_and_groups_unknowns() -> None:
    registry = CategoryRegistry(CONFIG)
    report = ImportValidator(registry).validate_current(build_result_with_unknown())

    assert report.is_valid
    assert not report.errors
    assert len(report.warnings) == 1
    assert report.metrics["source_row_count"] == 4
    assert report.metrics["normalized_record_count"] == 3
    assert report.metrics["unknown_record_count"] == 1
    assert report.metrics["reconciliation_difference"] == "0"
    assert len(report.exceptions) == 1
    assert report.exceptions[0].sample_label == "Mystery merchant"
    assert report.exceptions[0].total_amount == Decimal("25")


def test_writes_json_summary_and_csv_exceptions(tmp_path) -> None:
    registry = CategoryRegistry(CONFIG)
    report = ImportValidator(registry).validate_current(build_result_with_unknown())

    summary_path, exceptions_path = write_validation_report(report, tmp_path)

    assert summary_path.is_file()
    assert exceptions_path.is_file()
    assert '"is_valid": true' in summary_path.read_text(encoding="utf-8")
    assert "Mystery merchant" in exceptions_path.read_text(encoding="utf-8")


def test_detects_duplicate_source_rows() -> None:
    registry = CategoryRegistry(CONFIG)
    transaction = Transaction(
        year=2025,
        period="P1",
        category_id="FOD001",
        description="Groceries",
        amount=Decimal("10"),
        source_sheet="2025",
        source_cell="C2",
    )
    result = CurrentExtractionResult(
        records=(
            ExtractedRecord(section="variable_expenses", transaction=transaction),
            ExtractedRecord(section="variable_expenses", transaction=transaction),
        ),
        unknown_categories=(),
        source_rows=(
            SourceRow("P1", "Costco", Decimal("10"), "2025", "C2"),
            SourceRow("P1", "Costco", Decimal("10"), "2025", "C2"),
        ),
        periods=("P1",),
    )

    report = ImportValidator(registry).validate_current(result)

    assert not report.is_valid
    assert "DUPLICATE_SOURCE_ROWS" in {issue.code for issue in report.errors}


def test_detects_row_and_amount_reconciliation_failure() -> None:
    registry = CategoryRegistry(CONFIG)
    result = CurrentExtractionResult(
        records=(),
        unknown_categories=(),
        source_rows=(SourceRow("P1", "Costco", Decimal("10"), "2025", "C2"),),
        periods=("P1",),
    )

    report = ImportValidator(registry).validate_current(result)
    error_codes = {issue.code for issue in report.errors}

    assert "ROW_RECONCILIATION_FAILED" in error_codes
    assert "AMOUNT_RECONCILIATION_FAILED" in error_codes
