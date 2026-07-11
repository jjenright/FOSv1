"""Verify the production Family Financial Operating System release."""

from __future__ import annotations

import argparse
import re
import sys
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.extract import HistoricalWorkbookExtractor, LayoutDetector  # noqa: E402
from src.historical_pipeline import HistoricalPipeline  # noqa: E402
from src.insights import InsightsEngine  # noqa: E402
from src.kpi import KPIEngine  # noqa: E402
from src.load import HistoricalExcelFOSLoader  # noqa: E402
from src.transform import CategoryRegistry  # noqa: E402
from src.validate import HistoricalImportValidator, ImportValidator  # noqa: E402
from src.version import __version__  # noqa: E402

ZERO = Decimal("0")
ANNUAL_PATTERN = re.compile(r"20\d{2}( \(old\))?")


def workbook_sheet_names(workbook_path: Path) -> list[str]:
    """Read worksheet names without evaluating workbook formulas."""
    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    try:
        with ZipFile(workbook_path) as archive:
            workbook_xml = archive.read("xl/workbook.xml")
    except (FileNotFoundError, BadZipFile, KeyError) as exc:
        raise ValueError(f"Unable to read workbook structure: {workbook_path}") from exc

    root = ElementTree.fromstring(workbook_xml)
    sheets = root.find("main:sheets", namespace)
    if sheets is None:
        return []
    return [sheet.attrib["name"] for sheet in sheets]


def verify_current_snapshot_consistency(snapshot) -> None:
    """Validate snapshot relationships without hardcoding mutable balances."""
    expected_net_worth = snapshot.total_assets - snapshot.total_liabilities
    if snapshot.net_worth != expected_net_worth:
        raise ValueError(
            "Current net worth calculation mismatch: "
            f"expected {expected_net_worth}, found {snapshot.net_worth}."
        )
    if snapshot.fpi_score is not None:
        if snapshot.fpi_score < ZERO or snapshot.fpi_score > Decimal("100"):
            raise ValueError("Current FPI score is outside the valid 0–100 range.")
        expected_band = KPIEngine._fpi_band(snapshot.fpi_score)
        if snapshot.fpi_band != expected_band:
            raise ValueError(
                "Current FPI band mismatch: "
                f"expected {expected_band}, found {snapshot.fpi_band}."
            )


def verify_current_snapshot_sheet(worksheet, snapshot) -> None:
    """Confirm the generated Current_Snapshot sheet matches the calculated snapshot."""
    cell_net_worth = Decimal(str(worksheet["B7"].value))
    if cell_net_worth != snapshot.net_worth:
        raise ValueError(
            "Current net worth KPI output mismatch: "
            f"expected {snapshot.net_worth}, found {cell_net_worth}."
        )

    cell_fpi = worksheet["B16"].value
    if snapshot.fpi_score is None:
        if cell_fpi is not None:
            raise ValueError("Current FPI KPI output should be blank.")
    elif Decimal(str(cell_fpi)) != snapshot.fpi_score:
        raise ValueError(
            "Current FPI KPI output mismatch: "
            f"expected {snapshot.fpi_score}, found {cell_fpi}."
        )

    if worksheet["B17"].value != snapshot.fpi_band:
        raise ValueError(
            "Current FPI band output mismatch: "
            f"expected {snapshot.fpi_band}, found {worksheet['B17'].value}."
        )


def verify_configuration(workbook_path: Path, detector: LayoutDetector) -> None:
    """Ensure every annual source worksheet has an explicit layout rule."""
    source_sheets = workbook_sheet_names(workbook_path)
    annual_sheets = {name for name in source_sheets if ANNUAL_PATTERN.fullmatch(name)}
    configured = set(detector.configured_sheets())
    missing = sorted(annual_sheets - configured)
    if missing:
        raise ValueError("Unconfigured annual worksheets: " + ", ".join(missing))


def verify_dictionary_contract(registry: CategoryRegistry) -> None:
    """Protect the approved high-value normalization rules."""
    required = {
        "Costco": "FOD001",
        "Mortgage (21st)": "HOU001",
        "Fuel": "TRA002",
        "Visa payment": "TRF005",
        "Questrade TFSA": "TRF002",
        "Canadian Tire": "HOU012",
        "Walmart": "HOU012",
        "Amazon": "HOU012",
    }
    for label, expected_id in required.items():
        category = registry.lookup(label)
        if category is None or category.category_id != expected_id:
            actual = None if category is None else category.category_id
            raise ValueError(
                f"Category mapping mismatch for '{label}': expected {expected_id}, found {actual}."
            )


def verify_extraction_relationships(extraction, validation) -> dict[str, Decimal | int]:
    """Validate source-to-normalized reconciliation using live workbook values."""
    source_rows = len(extraction.source_rows)
    normalized_records = len(extraction.records)
    unknown_rows = len(extraction.unknown_categories)
    transaction_rows = (
        len(extraction.transfers)
        + len(extraction.variable_expenses)
        + len(extraction.fixed_expenses)
    )
    categorized_records = len(extraction.income) + transaction_rows

    if not extraction.sheets:
        raise ValueError("No official annual worksheets were imported.")
    if source_rows == 0:
        raise ValueError("The historical import did not find any financial rows.")
    if source_rows != normalized_records + unknown_rows:
        raise ValueError(
            "Source-row reconciliation mismatch: "
            f"{source_rows} source rows != {normalized_records} normalized + {unknown_rows} unknown."
        )
    if normalized_records != categorized_records:
        raise ValueError(
            "Normalized-record relationship mismatch: "
            f"{normalized_records} records != {categorized_records} categorized records."
        )
    if not validation.is_valid:
        messages = "; ".join(issue.message for issue in validation.errors)
        raise ValueError(f"Historical validation failed: {messages}")

    source_total = sum((item.amount for item in extraction.source_rows), ZERO)
    normalized_total = sum((item.transaction.amount for item in extraction.records), ZERO)
    unknown_total = sum((item.amount for item in extraction.unknown_categories), ZERO)
    reconciliation_difference = Decimal(
        str(validation.metrics["reconciliation_difference"])
    )
    if source_total != normalized_total + unknown_total:
        raise ValueError(
            "Dollar reconciliation mismatch: "
            f"{source_total} source != {normalized_total} normalized + {unknown_total} unknown."
        )
    if reconciliation_difference != ZERO:
        raise ValueError(
            f"Historical source reconciliation difference is {reconciliation_difference}."
        )

    imported_years = [sheet.year for sheet in extraction.sheets]
    if len(imported_years) != len(set(imported_years)):
        raise ValueError("Duplicate official years were imported.")
    archived_imports = [sheet.sheet_name for sheet in extraction.sheets if "(old)" in sheet.sheet_name]
    if archived_imports:
        raise ValueError("Archived worksheets were imported: " + ", ".join(archived_imports))

    return {
        "sheets": len(extraction.sheets),
        "periods": sum(len(sheet.result.periods) for sheet in extraction.sheets),
        "source_rows": source_rows,
        "records": normalized_records,
        "unknown_rows": unknown_rows,
        "unknown_labels": len(validation.exceptions),
        "income_rows": len(extraction.income),
        "transaction_rows": transaction_rows,
        "source_total": source_total,
    }


def verify_generated_workbook(
    output_path: Path,
    *,
    extraction,
    validation,
    annual_kpis,
    snapshot,
    insight_report,
) -> None:
    """Validate the generated XLSX structure and release outputs."""
    from openpyxl import load_workbook

    try:
        with ZipFile(output_path) as archive:
            corrupt_member = archive.testzip()
            if corrupt_member is not None:
                raise ValueError(f"Generated XLSX contains a corrupt member: {corrupt_member}")
    except BadZipFile as exc:
        raise ValueError("Generated FOS workbook is not a valid XLSX archive.") from exc

    workbook = load_workbook(output_path, data_only=False, read_only=False)
    try:
        required = list(HistoricalExcelFOSLoader.REQUIRED_SHEETS)
        if workbook.sheetnames != required:
            raise ValueError(f"Generated worksheet order mismatch: {workbook.sheetnames}")

        expected_transactions = (
            len(extraction.transfers)
            + len(extraction.variable_expenses)
            + len(extraction.fixed_expenses)
        )
        row_checks = {
            "DimYear": len(extraction.sheets) + 1,
            "FactTransactions": expected_transactions + 1,
            "FactIncome": len(extraction.income) + 1,
            "Exceptions": len(validation.exceptions) + 1,
            "Annual_KPIs": len(annual_kpis) + 1,
            "Insights": len(insight_report.insights) + 1,
            "Action_Plan": len(insight_report.actions) + 1,
        }
        for sheet_name, expected_rows in row_checks.items():
            actual_rows = workbook[sheet_name].max_row
            if actual_rows != expected_rows:
                raise ValueError(
                    f"{sheet_name} row count mismatch: expected {expected_rows}, found {actual_rows}."
                )

        transactions = workbook["FactTransactions"]
        if transactions.auto_filter.ref is not None:
            raise ValueError("FactTransactions contains an overlapping worksheet AutoFilter.")
        if "FactTransactionsTable" not in transactions.tables:
            raise ValueError("FactTransactionsTable is missing.")
        if "FactIncomeTable" not in workbook["FactIncome"].tables:
            raise ValueError("FactIncomeTable is missing.")

        verify_current_snapshot_sheet(workbook["Current_Snapshot"], snapshot)
        dashboard = workbook["Dashboard"]
        if dashboard["A1"].value != "Family Financial Operating System — Executive Dashboard":
            raise ValueError("Executive dashboard title is missing.")
        if dashboard["A8"].value != "Net Worth" or dashboard["A14"].value != "True Income":
            raise ValueError("Executive dashboard KPI cards are incomplete.")
        if dashboard["A61"].value != "Executive Takeaways":
            raise ValueError("Executive dashboard takeaways are missing.")
        if len(dashboard._charts) != 4:
            raise ValueError("Executive dashboard must contain four charts.")
        if not dashboard.column_dimensions["P"].hidden or not dashboard.column_dimensions["AB"].hidden:
            raise ValueError("Executive dashboard helper columns must remain hidden.")
        if dashboard.auto_filter.ref is not None:
            raise ValueError("Executive dashboard contains an unexpected AutoFilter.")
    finally:
        workbook.close()


def verify_private_workbook(workbook_path: Path) -> None:
    detector = LayoutDetector(PROJECT_ROOT / "config" / "layouts.yaml")
    registry = CategoryRegistry(PROJECT_ROOT / "config" / "categories.yaml")

    verify_configuration(workbook_path, detector)
    verify_dictionary_contract(registry)

    extraction = HistoricalWorkbookExtractor(detector, registry).extract(workbook_path)
    validation = HistoricalImportValidator(ImportValidator(registry)).validate(extraction)
    metrics = verify_extraction_relationships(extraction, validation)

    kpi_engine = KPIEngine(registry)
    annual_kpis = kpi_engine.calculate_annual(extraction)
    if len(annual_kpis) != len(extraction.sheets):
        raise ValueError("Annual KPI count does not match imported annual worksheets.")
    complete = [item for item in annual_kpis if item.coverage_status == "Complete"]
    if not complete:
        raise ValueError("No complete year is available for current financial KPIs.")

    snapshot = kpi_engine.calculate_current_snapshot(workbook_path, annual_kpis)
    verify_current_snapshot_consistency(snapshot)
    insight_report = InsightsEngine(registry).analyze(extraction, annual_kpis, snapshot)
    if insight_report.latest_year != snapshot.latest_complete_year:
        raise ValueError("Insight latest year does not match the current snapshot year.")
    if len(insight_report.insights) != 6 or len(insight_report.actions) != 6:
        raise ValueError("The production insight layer must generate six insights and six actions.")
    expected_target = (
        next(item for item in annual_kpis if item.year == insight_report.latest_year).core_spending
        / Decimal("12")
        * insight_report.emergency_target_months
    )
    expected_gap = max(ZERO, expected_target - snapshot.savings_cash)
    if insight_report.emergency_target_amount != expected_target:
        raise ValueError("Emergency-reserve target calculation mismatch.")
    if insight_report.emergency_fund_gap != expected_gap:
        raise ValueError("Emergency-reserve funding gap mismatch.")

    with TemporaryDirectory() as temporary_dir:
        output = Path(temporary_dir) / "Financial_Operating_System.xlsx"
        pipeline = HistoricalPipeline(PROJECT_ROOT).run(
            workbook_path,
            output_path=output,
            fos_version=__version__,
        )
        if not pipeline.validation_summary_path.is_file():
            raise ValueError("Historical validation summary was not created.")
        if not pipeline.exceptions_path.is_file():
            raise ValueError("Historical exceptions report was not created.")
        if pipeline.insight_report is None:
            raise ValueError("Historical insight report was not created.")
        verify_generated_workbook(
            output,
            extraction=extraction,
            validation=validation,
            annual_kpis=annual_kpis,
            snapshot=snapshot,
            insight_report=insight_report,
        )

    print(f"- Official worksheets imported: {metrics['sheets']}")
    print(f"- Pay periods extracted: {metrics['periods']}")
    print(f"- Source rows reconciled: {metrics['source_rows']}")
    print(f"- Normalized records: {metrics['records']}")
    print(f"- Unmapped rows: {metrics['unknown_rows']}")
    print(f"- Unmapped labels: {metrics['unknown_labels']}")
    print(f"- Source total: ${metrics['source_total']:,.2f}")
    print("- Reconciliation difference: $0.00")
    print(f"- Latest complete year: {snapshot.latest_complete_year}")
    print(f"- Current net worth: ${snapshot.net_worth:,.2f}")
    if snapshot.fpi_score is not None:
        print(f"- Provisional FPI: {snapshot.fpi_score} ({snapshot.fpi_band})")
    print("- Dashboard, tables, insights and action plan: OK")
    print("- XLSX archive integrity: OK")


def main() -> int:
    parser = argparse.ArgumentParser(description=f"Verify FOS v{__version__}")
    parser.add_argument("--workbook", type=Path, help="Optional private budget workbook path.")
    args = parser.parse_args()

    detector = LayoutDetector(PROJECT_ROOT / "config" / "layouts.yaml")
    registry = CategoryRegistry(PROJECT_ROOT / "config" / "categories.yaml")

    print(f"FOS v{__version__} verification")
    print("- Version source: OK")
    print(f"- Configured categories: {registry.category_count()}")
    print(f"- Configured aliases: {registry.alias_count()}")
    verify_dictionary_contract(registry)
    print("- Approved category mappings: OK")
    print(f"- Legacy layout detection: {detector.detect('2010')}")
    print(f"- Current layout detection: {detector.detect('2025')}")

    if args.workbook:
        verify_private_workbook(args.workbook)

    print("Verification PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
