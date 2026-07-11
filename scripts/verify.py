"""Verify the FOS v0.2.0-alpha.5 components."""

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

from src.extract import CurrentLayoutExtractor, LayoutDetector  # noqa: E402
from src.models import Category, ImportResult, ImportSession, Transaction  # noqa: E402,F401
from src.transform import CategoryRegistry  # noqa: E402
from src.validate import ImportValidator, write_validation_report  # noqa: E402


def workbook_sheet_names(workbook_path: Path) -> list[str]:
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


def verify_required_workbook_aliases(workbook_path: Path, registry: CategoryRegistry) -> None:
    from openpyxl import load_workbook

    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    try:
        annual_pattern = re.compile(r"20\d{2}( \(old\))?")
        labels: set[str] = set()
        for sheet_name in workbook.sheetnames:
            if not annual_pattern.fullmatch(sheet_name):
                continue
            worksheet = workbook[sheet_name]
            for row in worksheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        labels.add(" ".join(cell.value.split()))
    finally:
        workbook.close()

    required = {
        "Costco": "FOD001",
        "Mortgage (21st)": "HOU001",
        "Fuel": "TRA002",
        "Visa payment": "TRF005",
        "Questrade TFSA": "TRF002",
        "Canadian Tire": "HOU012",
    }
    for label, expected_id in required.items():
        if label not in labels:
            raise ValueError(f"Required workbook label not found: {label}")
        actual_id = registry.lookup(label).category_id
        if actual_id != expected_id:
            raise ValueError(
                f"Workbook label '{label}' mapped to '{actual_id}', expected '{expected_id}'."
            )


def verify_2025_extraction_and_validation(
    workbook_path: Path, registry: CategoryRegistry
) -> None:
    extractor = CurrentLayoutExtractor(registry)
    result = extractor.extract(workbook_path, "2025")
    report = ImportValidator(registry).validate_current(result)

    expected = {
        "periods": 26,
        "source_rows": 497,
        "records": 463,
        "unknowns": 34,
        "unknown_labels": 28,
        "income_count": 85,
        "transfer_count": 86,
        "variable_count": 157,
        "fixed_count": 135,
        "income_total": Decimal("215878.10"),
        "transfer_total": Decimal("119717.02"),
        "variable_total": Decimal("28492.48"),
        "fixed_total": Decimal("59879.19"),
        "unknown_total": Decimal("8112.69"),
        "source_total": Decimal("432079.48"),
    }
    actual = {
        "periods": len(result.periods),
        "source_rows": len(result.source_rows),
        "records": len(result.records),
        "unknowns": len(result.unknown_categories),
        "unknown_labels": len(report.exceptions),
        "income_count": len(result.income),
        "transfer_count": len(result.transfers),
        "variable_count": len(result.variable_expenses),
        "fixed_count": len(result.fixed_expenses),
        "income_total": sum((item.amount for item in result.income), Decimal("0")),
        "transfer_total": sum((item.amount for item in result.transfers), Decimal("0")),
        "variable_total": sum(
            (item.amount for item in result.variable_expenses), Decimal("0")
        ),
        "fixed_total": sum((item.amount for item in result.fixed_expenses), Decimal("0")),
        "unknown_total": sum(
            (item.amount for item in result.unknown_categories), Decimal("0")
        ),
        "source_total": sum((item.amount for item in result.source_rows), Decimal("0")),
    }
    if actual != expected:
        differences = [
            f"{key}: expected {expected[key]}, found {actual[key]}"
            for key in expected
            if actual[key] != expected[key]
        ]
        raise ValueError("2025 extraction mismatch: " + "; ".join(differences))

    if not report.is_valid:
        raise ValueError(
            "2025 validation failed: "
            + "; ".join(issue.message for issue in report.errors)
        )
    if Decimal(str(report.metrics["reconciliation_difference"])) != 0:
        raise ValueError("2025 source reconciliation did not equal zero.")

    with TemporaryDirectory() as temporary_dir:
        summary_path, exceptions_path = write_validation_report(report, temporary_dir)
        if not summary_path.is_file() or not exceptions_path.is_file():
            raise ValueError("Validation report files were not created.")

    print(f"- 2025 pay periods extracted: {actual['periods']}")
    print(f"- 2025 source financial rows: {actual['source_rows']}")
    print(f"- 2025 normalized records: {actual['records']}")
    print(f"- 2025 unmapped records: {actual['unknowns']}")
    print(f"- 2025 unmapped labels: {actual['unknown_labels']}")
    print(f"- 2025 source total: ${actual['source_total']:,.2f}")
    print(f"- 2025 reconciliation difference: {report.metrics['reconciliation_difference']}")
    print("- 2025 validation and exceptions reports: OK")


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify FOS v0.2.0-alpha.5")
    parser.add_argument("--workbook", type=Path, help="Optional private Budget workbook path.")
    args = parser.parse_args()

    detector = LayoutDetector(PROJECT_ROOT / "config" / "layouts.yaml")
    registry = CategoryRegistry(PROJECT_ROOT / "config" / "categories.yaml")

    print("FOS v0.2.0-alpha.5 verification")
    print("- Core models: OK")
    print(f"- Configured categories: {registry.category_count()}")
    print(f"- Configured aliases: {registry.alias_count()}")
    print(f"- Costco mapping: {registry.lookup('Costco').display_name}")
    print(f"- Canadian Tire mapping: {registry.lookup('Canadian Tire').display_name}")
    print(f"- Mortgage (21st) mapping: {registry.lookup('Mortgage (21st)').display_name}")
    print(f"- 2010 layout: {detector.detect('2010')}")
    print(f"- 2017 (old) layout: {detector.detect('2017 (old)')}")
    print(f"- 2025 layout: {detector.detect('2025')}")

    if args.workbook:
        source_sheets = workbook_sheet_names(args.workbook)
        configured = set(detector.configured_sheets())
        annual_pattern = re.compile(r"20\d{2}( \(old\))?")
        annual_sheets = {name for name in source_sheets if annual_pattern.fullmatch(name)}
        missing = sorted(annual_sheets - configured)
        extra = sorted(configured - annual_sheets)
        print(f"- Workbook annual sheets found: {len(annual_sheets)}")
        if missing:
            print(f"- Unconfigured annual sheets: {', '.join(missing)}")
            return 1
        if extra:
            print(f"- Configured sheets absent from workbook: {', '.join(extra)}")
            return 1
        print("- Workbook/layout configuration match: OK")
        verify_required_workbook_aliases(args.workbook, registry)
        print("- Workbook/category dictionary checks: OK")
        verify_2025_extraction_and_validation(args.workbook, registry)

    print("Verification PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
