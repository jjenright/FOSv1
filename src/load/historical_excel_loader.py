"""Excel loader for validated workbook-wide historical imports."""

from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from src.extract import HistoricalExtractionResult
from src.load.excel_loader import (
    BLACK,
    CURRENCY_FORMAT,
    DARK_BLUE,
    DATE_TIME_FORMAT,
    INTEGER_FORMAT,
    LIGHT_GREEN,
    LIGHT_YELLOW,
    WHITE,
    ExcelFOSLoader,
    LoadResult,
)
from src.transform import CategoryRegistry
from src.validate import HistoricalValidationReport


class HistoricalExcelFOSLoader(ExcelFOSLoader):
    """Write all official annual worksheets into one FOS workbook."""

    REQUIRED_SHEETS = (
        "Dashboard",
        "Import_Log",
        "DimYear",
        "DimCategory",
        "FactTransactions",
        "FactIncome",
        "Exceptions",
        "Validation",
    )

    def __init__(self, category_registry: CategoryRegistry) -> None:
        super().__init__(category_registry)

    def load_historical(
        self,
        extraction: HistoricalExtractionResult,
        validation: HistoricalValidationReport,
        output_path: str | Path,
        *,
        source_workbook: str | Path,
        fos_version: str,
        imported_at: datetime | None = None,
    ) -> LoadResult:
        if not validation.is_valid:
            messages = "; ".join(issue.message for issue in validation.errors)
            raise ValueError(f"Cannot load invalid historical import: {messages}")

        destination = Path(output_path)
        destination.parent.mkdir(parents=True, exist_ok=True)
        imported_at = imported_at or datetime.now(timezone.utc).replace(tzinfo=None)

        workbook = Workbook()
        workbook.remove(workbook.active)
        for sheet_name in self.REQUIRED_SHEETS:
            workbook.create_sheet(sheet_name)

        self._write_categories(workbook["DimCategory"])
        self._write_transactions(workbook["FactTransactions"], extraction)
        self._write_income(workbook["FactIncome"], extraction)
        self._write_exceptions(workbook["Exceptions"], validation.aggregate)
        self._write_validation(workbook["Validation"], validation.aggregate)
        self._write_dim_year(workbook["DimYear"], extraction, validation)
        self._write_import_log_historical(
            workbook["Import_Log"],
            extraction=extraction,
            validation=validation,
            source_workbook=Path(source_workbook),
            fos_version=fos_version,
            imported_at=imported_at,
            output_path=destination,
        )
        self._write_dashboard_historical(
            workbook["Dashboard"],
            extraction=extraction,
            validation=validation,
            source_workbook=Path(source_workbook),
            fos_version=fos_version,
            imported_at=imported_at,
        )

        workbook.save(destination)
        workbook.close()

        return LoadResult(
            output_path=destination,
            category_rows=self.category_registry.category_count(),
            transaction_rows=(
                len(extraction.transfers)
                + len(extraction.variable_expenses)
                + len(extraction.fixed_expenses)
            ),
            income_rows=len(extraction.income),
            exception_rows=len(validation.exceptions),
        )

    def _write_dim_year(
        self,
        worksheet: Any,
        extraction: HistoricalExtractionResult,
        validation: HistoricalValidationReport,
    ) -> None:
        reports = {item.sheet_name: item for item in validation.sheets}
        headers = [
            "Year",
            "SourceSheet",
            "Layout",
            "PayPeriods",
            "SourceRows",
            "NormalizedRecords",
            "UnknownRows",
            "Income ($)",
            "Transfers ($)",
            "VariableExpenses ($)",
            "FixedExpenses ($)",
            "UnknownAmount ($)",
            "Status",
        ]
        rows: list[list[Any]] = []
        for sheet in extraction.sheets:
            result = sheet.result
            report = reports[sheet.sheet_name].report
            rows.append(
                [
                    sheet.year,
                    sheet.sheet_name,
                    sheet.layout,
                    len(result.periods),
                    len(result.source_rows),
                    len(result.records),
                    len(result.unknown_categories),
                    float(sum((item.amount for item in result.income), Decimal("0"))),
                    float(sum((item.amount for item in result.transfers), Decimal("0"))),
                    float(
                        sum(
                            (item.amount for item in result.variable_expenses),
                            Decimal("0"),
                        )
                    ),
                    float(
                        sum(
                            (item.amount for item in result.fixed_expenses),
                            Decimal("0"),
                        )
                    ),
                    float(
                        sum(
                            (item.amount for item in result.unknown_categories),
                            Decimal("0"),
                        )
                    ),
                    "PASS" if report.is_valid else "FAIL",
                ]
            )
        self._write_table(worksheet, headers, rows, "DimYearTable")
        worksheet.freeze_panes = "A2"
        for row in range(2, worksheet.max_row + 1):
            for column in range(8, 13):
                worksheet.cell(row, column).number_format = CURRENCY_FORMAT
            worksheet.cell(row, 13).fill = PatternFill(
                "solid", fgColor=LIGHT_GREEN if worksheet.cell(row, 13).value == "PASS" else LIGHT_YELLOW
            )
        self._set_widths(
            worksheet,
            [10, 16, 24, 12, 12, 19, 13, 16, 16, 22, 20, 18, 12],
        )

    def _write_import_log_historical(
        self,
        worksheet: Any,
        *,
        extraction: HistoricalExtractionResult,
        validation: HistoricalValidationReport,
        source_workbook: Path,
        fos_version: str,
        imported_at: datetime,
        output_path: Path,
    ) -> None:
        report_by_sheet = {item.sheet_name: item for item in validation.sheets}
        headers = [
            "ImportDate",
            "SourceWorkbook",
            "SourceSheet",
            "Layout",
            "FOSVersion",
            "Status",
            "Periods",
            "SourceRows",
            "NormalizedRecords",
            "UnknownRows",
            "Warnings",
            "Errors",
            "ReconciliationDifference",
            "OutputWorkbook",
        ]
        rows: list[list[Any]] = []
        for sheet in extraction.sheets:
            report = report_by_sheet[sheet.sheet_name].report
            rows.append(
                [
                    imported_at,
                    source_workbook.name,
                    sheet.sheet_name,
                    sheet.layout,
                    fos_version,
                    "PASS" if report.is_valid else "FAIL",
                    len(sheet.result.periods),
                    len(sheet.result.source_rows),
                    len(sheet.result.records),
                    len(sheet.result.unknown_categories),
                    len(report.warnings),
                    len(report.errors),
                    float(Decimal(str(report.metrics["reconciliation_difference"]))),
                    output_path.name,
                ]
            )
        self._write_table(worksheet, headers, rows, "ImportLogTable")
        worksheet.freeze_panes = "A2"
        for row in range(2, worksheet.max_row + 1):
            worksheet.cell(row, 1).number_format = DATE_TIME_FORMAT
            worksheet.cell(row, 13).number_format = CURRENCY_FORMAT
            worksheet.cell(row, 6).fill = PatternFill(
                "solid", fgColor=LIGHT_GREEN if worksheet.cell(row, 6).value == "PASS" else LIGHT_YELLOW
            )
        self._set_widths(
            worksheet,
            [20, 34, 16, 24, 14, 12, 10, 12, 19, 13, 10, 10, 24, 34],
        )

    def _write_dashboard_historical(
        self,
        worksheet: Any,
        *,
        extraction: HistoricalExtractionResult,
        validation: HistoricalValidationReport,
        source_workbook: Path,
        fos_version: str,
        imported_at: datetime,
    ) -> None:
        worksheet.sheet_view.showGridLines = False
        worksheet.merge_cells("A1:H2")
        worksheet["A1"] = "Family Financial Operating System — Historical Data Load"
        worksheet["A1"].font = Font(size=20, bold=True, color=WHITE)
        worksheet["A1"].fill = PatternFill("solid", fgColor=DARK_BLUE)
        worksheet["A1"].alignment = Alignment(horizontal="left", vertical="center")

        worksheet["A4"] = "Version"
        worksheet["B4"] = fos_version
        worksheet["D4"] = "Source"
        worksheet["E4"] = source_workbook.name
        worksheet["G4"] = "Years"
        worksheet["H4"] = len(extraction.sheets)
        worksheet["A5"] = "Imported"
        worksheet["B5"] = imported_at
        worksheet["B5"].number_format = DATE_TIME_FORMAT
        worksheet["D5"] = "Range"
        worksheet["E5"] = (
            f"{extraction.sheets[0].year}–{extraction.sheets[-1].year}"
            if extraction.sheets
            else ""
        )
        for cell in ("A4", "D4", "G4", "A5", "D5"):
            worksheet[cell].font = Font(bold=True, color=DARK_BLUE)

        worksheet["A8"] = "Metric"
        worksheet["B8"] = "Value"
        metrics = [
            ("Income", "=SUM(FactIncome!F:F)", CURRENCY_FORMAT),
            (
                "Transfers",
                '=SUMIF(FactTransactions!D:D,"transfers",FactTransactions!G:G)',
                CURRENCY_FORMAT,
            ),
            (
                "Variable / irregular expenses",
                '=SUMIF(FactTransactions!D:D,"variable_expenses",FactTransactions!G:G)',
                CURRENCY_FORMAT,
            ),
            (
                "Fixed expenses",
                '=SUMIF(FactTransactions!D:D,"fixed_expenses",FactTransactions!G:G)',
                CURRENCY_FORMAT,
            ),
            ("Unmapped amount", "=SUM(Exceptions!D:D)", CURRENCY_FORMAT),
            (
                "Normalized records",
                "=COUNTA(FactTransactions!A:A)-1+COUNTA(FactIncome!A:A)-1",
                INTEGER_FORMAT,
            ),
            (
                "Exceptions requiring review",
                "=COUNTA(Exceptions!A:A)-1",
                INTEGER_FORMAT,
            ),
            ("Years imported", "=COUNTA(DimYear!A:A)-1", INTEGER_FORMAT),
        ]
        for row_number, (label, formula, number_format) in enumerate(metrics, start=9):
            worksheet.cell(row_number, 1, label)
            worksheet.cell(row_number, 2, formula)
            worksheet.cell(row_number, 2).number_format = number_format
            worksheet.cell(row_number, 2).font = Font(color=BLACK)
        self._style_header(worksheet, 8, 2)

        worksheet["D8"] = "Import Status"
        worksheet["E8"] = "PASS" if validation.is_valid else "FAIL"
        worksheet["D8"].font = Font(bold=True, color=WHITE)
        worksheet["D8"].fill = PatternFill("solid", fgColor=DARK_BLUE)
        worksheet["E8"].font = Font(bold=True, color=DARK_BLUE)
        worksheet["E8"].fill = PatternFill(
            "solid", fgColor=LIGHT_GREEN if validation.is_valid else LIGHT_YELLOW
        )
        worksheet["D10"] = (
            "Historical import is reconciled. Review the Exceptions sheet before "
            "using unmapped amounts in analysis."
        )
        worksheet.merge_cells("D10:H12")
        worksheet["D10"].alignment = Alignment(wrap_text=True, vertical="top")
        worksheet["D10"].fill = PatternFill("solid", fgColor=LIGHT_YELLOW)

        self._set_widths(worksheet, [34, 20, 4, 22, 30, 4, 14, 22])
