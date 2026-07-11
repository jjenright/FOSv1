"""Excel loader for validated workbook-wide historical imports."""

from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from src.extract import HistoricalExtractionResult
from src.kpi import AnnualKPI, CurrentSnapshot
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
        "Annual_KPIs",
        "Current_Snapshot",
        "KPI_Definitions",
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
        annual_kpis: tuple[AnnualKPI, ...] = (),
        current_snapshot: CurrentSnapshot | None = None,
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
        self._write_annual_kpis(workbook["Annual_KPIs"], annual_kpis)
        self._write_current_snapshot(workbook["Current_Snapshot"], current_snapshot)
        self._write_kpi_definitions(workbook["KPI_Definitions"])
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
            annual_kpis=annual_kpis,
            current_snapshot=current_snapshot,
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

    def _write_annual_kpis(
        self,
        worksheet: Any,
        annual_kpis: tuple[AnnualKPI, ...],
    ) -> None:
        headers = [
            "Year",
            "Coverage",
            "PayPeriods",
            "TrueIncome ($)",
            "CashFlowAdjustments ($)",
            "FixedExpenses ($)",
            "VariableIrregularExpenses ($)",
            "KnownOperatingExpenses ($)",
            "CoreSpending ($)",
            "LifestyleSpending ($)",
            "WealthBuilding ($)",
            "TargetedDebtReduction ($)",
            "ExcludedTransfers ($)",
            "UnknownAmount ($)",
            "FixedCostRatio",
            "KnownExpenseRatio",
            "WealthBuildingRate",
            "SavingsVelocity",
            "FinancialFlexibility",
            "HousingRatio",
            "TransportationRatio",
            "FoodRatio",
            "LifestyleRatio",
            "DataCoverageRatio",
            "ComparisonEligible",
        ]
        rows: list[list[Any]] = []
        for item in annual_kpis:
            rows.append(
                [
                    item.year,
                    item.coverage_status,
                    item.pay_periods,
                    float(item.true_income),
                    float(item.cash_flow_adjustments),
                    float(item.fixed_expenses),
                    float(item.variable_irregular_expenses),
                    float(item.known_operating_expenses),
                    float(item.core_spending),
                    float(item.lifestyle_spending),
                    float(item.wealth_building),
                    float(item.targeted_debt_reduction),
                    float(item.excluded_transfers),
                    float(item.unknown_amount),
                    float(item.fixed_cost_ratio) if item.fixed_cost_ratio is not None else None,
                    float(item.known_expense_ratio) if item.known_expense_ratio is not None else None,
                    float(item.wealth_building_rate) if item.wealth_building_rate is not None else None,
                    float(item.savings_velocity) if item.savings_velocity is not None else None,
                    float(item.financial_flexibility) if item.financial_flexibility is not None else None,
                    float(item.housing_ratio) if item.housing_ratio is not None else None,
                    float(item.transportation_ratio) if item.transportation_ratio is not None else None,
                    float(item.food_ratio) if item.food_ratio is not None else None,
                    float(item.lifestyle_ratio) if item.lifestyle_ratio is not None else None,
                    float(item.data_coverage_ratio) if item.data_coverage_ratio is not None else None,
                    item.comparison_eligible,
                ]
            )
        self._write_table(worksheet, headers, rows, "AnnualKPIsTable")
        worksheet.freeze_panes = "A2"
        for row in range(2, worksheet.max_row + 1):
            for column in range(4, 15):
                worksheet.cell(row, column).number_format = CURRENCY_FORMAT
            for column in range(15, 25):
                worksheet.cell(row, column).number_format = '0.0%;[Red](0.0%);-'
            if worksheet.cell(row, 2).value == "Partial":
                worksheet.cell(row, 2).fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
            if not worksheet.cell(row, 25).value:
                worksheet.cell(row, 25).fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
        self._set_widths(
            worksheet,
            [10, 12, 12, 16, 22, 18, 28, 26, 18, 20, 20, 25, 20, 18,
             17, 19, 20, 18, 20, 15, 21, 14, 17, 18, 20],
        )
        if len(rows) >= 2:
            chart = LineChart()
            chart.title = "Annual Income, Known Expenses and Wealth Building"
            chart.y_axis.title = "Amount ($)"
            chart.x_axis.title = "Year"
            chart.style = 10
            data = Reference(worksheet, min_col=4, max_col=11, min_row=1, max_row=worksheet.max_row)
            categories = Reference(worksheet, min_col=1, min_row=2, max_row=worksheet.max_row)
            # Keep only TrueIncome, KnownOperatingExpenses and WealthBuilding series.
            chart.add_data(Reference(worksheet, min_col=4, min_row=1, max_row=worksheet.max_row), titles_from_data=True)
            chart.add_data(Reference(worksheet, min_col=8, min_row=1, max_row=worksheet.max_row), titles_from_data=True)
            chart.add_data(Reference(worksheet, min_col=11, min_row=1, max_row=worksheet.max_row), titles_from_data=True)
            chart.set_categories(categories)
            chart.height = 8
            chart.width = 15
            worksheet.add_chart(chart, "AA2")

    def _write_current_snapshot(
        self,
        worksheet: Any,
        snapshot: CurrentSnapshot | None,
    ) -> None:
        worksheet.sheet_view.showGridLines = False
        worksheet.merge_cells("A1:D1")
        worksheet["A1"] = "Current Financial Snapshot"
        worksheet["A1"].font = Font(bold=True, color=WHITE, size=14)
        worksheet["A1"].fill = PatternFill("solid", fgColor=DARK_BLUE)
        if snapshot is None:
            worksheet["A3"] = "Current snapshot unavailable."
            return
        rows = [
            ("Source sheet", snapshot.source_sheet, "Source workbook"),
            ("Latest complete cash-flow year", snapshot.latest_complete_year, "Annual_KPIs"),
            ("Total assets", float(snapshot.total_assets), "A & L"),
            ("Total liabilities", float(snapshot.total_liabilities), "A & L"),
            ("Net worth", float(snapshot.net_worth), "A & L"),
            ("Financial assets", float(snapshot.financial_assets), "A & L"),
            ("Savings cash", float(snapshot.savings_cash), "A & L"),
            ("Home equity", float(snapshot.home_equity), "A & L"),
            ("Mortgage balance", float(snapshot.mortgage_balance), "A & L"),
            ("Line of credit balance", float(snapshot.line_of_credit_balance), "A & L"),
            ("Vehicle loan balance", float(snapshot.vehicle_loan_balance), "A & L"),
            (
                "Emergency fund months",
                float(snapshot.emergency_fund_months) if snapshot.emergency_fund_months is not None else None,
                "Savings ÷ monthly core spending",
            ),
            (
                "Unsecured debt ratio",
                float(snapshot.unsecured_debt_ratio) if snapshot.unsecured_debt_ratio is not None else None,
                "LOC ÷ latest complete-year income",
            ),
            (
                "Financial Pressure Index",
                float(snapshot.fpi_score) if snapshot.fpi_score is not None else None,
                "Provisional FPI v1",
            ),
            ("FPI band", snapshot.fpi_band, "Low / Moderate / High / Very High"),
        ]
        worksheet.append(["Metric", "Value", "Source / Definition", "Notes"])
        for label, value, source in rows:
            worksheet.append([label, value, source, ""])
        self._style_header(worksheet, 2, 4)
        for row in range(3, 18):
            label = worksheet.cell(row, 1).value
            if label in {
                "Total assets", "Total liabilities", "Net worth", "Financial assets",
                "Savings cash", "Home equity", "Mortgage balance", "Line of credit balance",
                "Vehicle loan balance",
            }:
                worksheet.cell(row, 2).number_format = CURRENCY_FORMAT
            elif label == "Unsecured debt ratio":
                worksheet.cell(row, 2).number_format = '0.0%;[Red](0.0%);-'
            elif label in {"Emergency fund months", "Financial Pressure Index"}:
                worksheet.cell(row, 2).number_format = '0.0'
        worksheet["D16"] = (
            "FPI is a directional household metric, not an industry-standard score. "
            "It uses the latest complete-year cash flow plus current savings and LOC balances."
        )
        worksheet["D16"].alignment = Alignment(wrap_text=True, vertical="top")
        worksheet["D16"].fill = PatternFill("solid", fgColor=LIGHT_YELLOW)
        self._set_widths(worksheet, [34, 20, 38, 56])
        worksheet.freeze_panes = "A3"

    def _write_kpi_definitions(self, worksheet: Any) -> None:
        headers = ["KPI", "Definition", "Calculation", "Important limitation"]
        rows = [
            ["True Income", "Net household income recorded in the annual sheets.", "INC001 + INC002 + INC003", "Excludes cash-flow adjustment rows."],
            ["Known Operating Expenses", "Mapped fixed, variable, irregular and capital outflows.", "Fixed expenses + variable/irregular expenses", "Unmapped amounts are shown separately."],
            ["Fixed Cost Ratio", "Share of true income committed to mapped fixed expenses.", "Fixed expenses ÷ true income", "Based on recorded workbook classifications."],
            ["Wealth-Building Rate", "Share of income directed to savings, TFSA, RRSP/pension and RESP.", "Wealth-building transfers ÷ true income", "Does not include investment market growth."],
            ["Savings Velocity", "Wealth building plus targeted LOC/student-loan reduction.", "(Wealth building + targeted debt reduction) ÷ true income", "Does not estimate mortgage or vehicle-loan principal."],
            ["Financial Flexibility", "Income remaining after fixed expenses and targeted debt reduction.", "(True income - fixed expenses - targeted debt reduction) ÷ true income", "Variable living costs must still be paid from this amount."],
            ["Data Coverage Ratio", "Share of mapped outflows versus mapped plus unmapped outflows.", "Mapped outflows ÷ (mapped outflows + unknown amount)", "Measures classification completeness, not financial health."],
            ["Comparison Eligible", "Year suitable for direct trend comparison.", "Complete year and data coverage ≥ 85%", "Partial or lower-coverage years remain visible but are flagged."],
            ["FPI v1", "Directional Financial Pressure Index from 0 to 100.", "30% fixed costs + 25% emergency reserve + 30% LOC ratio + 15% known cash margin", "Provisional, household-specific and not an industry-standard score."],
        ]
        self._write_table(worksheet, headers, rows, "KPIDefinitionsTable")
        worksheet.freeze_panes = "A2"
        self._set_widths(worksheet, [26, 55, 58, 62])
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    def _write_dashboard_historical(
        self,
        worksheet: Any,
        *,
        extraction: HistoricalExtractionResult,
        validation: HistoricalValidationReport,
        source_workbook: Path,
        fos_version: str,
        imported_at: datetime,
        annual_kpis: tuple[AnnualKPI, ...],
        current_snapshot: CurrentSnapshot | None,
    ) -> None:
        worksheet.sheet_view.showGridLines = False
        worksheet.merge_cells("A1:H2")
        worksheet["A1"] = "Family Financial Operating System — KPI Engine"
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

        complete = [item for item in annual_kpis if item.coverage_status == "Complete"]
        latest = max(complete, key=lambda item: item.year) if complete else None

        worksheet["A8"] = "Current Position"
        worksheet["B8"] = "Value"
        self._style_header(worksheet, 8, 2)
        current_metrics = [
            ("Net worth", current_snapshot.net_worth if current_snapshot else None, CURRENCY_FORMAT),
            ("Financial assets", current_snapshot.financial_assets if current_snapshot else None, CURRENCY_FORMAT),
            ("Savings cash", current_snapshot.savings_cash if current_snapshot else None, CURRENCY_FORMAT),
            ("Line of credit", current_snapshot.line_of_credit_balance if current_snapshot else None, CURRENCY_FORMAT),
            ("Emergency fund months", current_snapshot.emergency_fund_months if current_snapshot else None, "0.0"),
            ("Financial Pressure Index", current_snapshot.fpi_score if current_snapshot else None, "0.0"),
            ("FPI band", current_snapshot.fpi_band if current_snapshot else "Unavailable", "General"),
        ]
        for row, (label, value, fmt) in enumerate(current_metrics, start=9):
            worksheet.cell(row, 1, label)
            worksheet.cell(row, 2, float(value) if isinstance(value, Decimal) else value)
            worksheet.cell(row, 2).number_format = fmt

        worksheet["D8"] = f"Latest Complete Year ({latest.year if latest else 'N/A'})"
        worksheet["E8"] = "Value"
        self._style_header(worksheet, 8, 5)
        latest_metrics = [
            ("True income", latest.true_income if latest else None, CURRENCY_FORMAT),
            ("Known operating expenses", latest.known_operating_expenses if latest else None, CURRENCY_FORMAT),
            ("Fixed cost ratio", latest.fixed_cost_ratio if latest else None, '0.0%;[Red](0.0%);-'),
            ("Wealth-building rate", latest.wealth_building_rate if latest else None, '0.0%;[Red](0.0%);-'),
            ("Savings velocity", latest.savings_velocity if latest else None, '0.0%;[Red](0.0%);-'),
            ("Financial flexibility", latest.financial_flexibility if latest else None, '0.0%;[Red](0.0%);-'),
            ("Data coverage", latest.data_coverage_ratio if latest else None, '0.0%;[Red](0.0%);-'),
        ]
        for row, (label, value, fmt) in enumerate(latest_metrics, start=9):
            worksheet.cell(row, 4, label)
            worksheet.cell(row, 5, float(value) if isinstance(value, Decimal) else value)
            worksheet.cell(row, 5).number_format = fmt

        worksheet["G8"] = "Status"
        worksheet["H8"] = "PASS" if validation.is_valid else "FAIL"
        worksheet["G8"].font = Font(bold=True, color=WHITE)
        worksheet["G8"].fill = PatternFill("solid", fgColor=DARK_BLUE)
        worksheet["H8"].font = Font(bold=True, color=DARK_BLUE)
        worksheet["H8"].fill = PatternFill(
            "solid", fgColor=LIGHT_GREEN if validation.is_valid else LIGHT_YELLOW
        )
        worksheet["G10"] = (
            "Use Annual_KPIs for trend analysis. Years marked Partial or not comparison-eligible "
            "should not be interpreted as full-year changes. FPI v1 is provisional."
        )
        worksheet.merge_cells("G10:H15")
        worksheet["G10"].alignment = Alignment(wrap_text=True, vertical="top")
        worksheet["G10"].fill = PatternFill("solid", fgColor=LIGHT_YELLOW)

        self._set_widths(worksheet, [34, 20, 4, 34, 20, 4, 25, 25])
        for row in range(9, 16):
            for col in (1, 2, 4, 5):
                worksheet.cell(row, col).border = Border(
                    bottom=Side(style="thin", color="D9E1F2")
                )

