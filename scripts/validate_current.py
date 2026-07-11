"""Validate one current-layout worksheet and write exceptions reports."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.extract import CurrentLayoutExtractor  # noqa: E402
from src.transform import CategoryRegistry  # noqa: E402
from src.validate import ImportValidator, write_validation_report  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a current-layout FOS worksheet.")
    parser.add_argument("workbook", type=Path, help="Path to the private source workbook.")
    parser.add_argument("--sheet", default="2025", help="Annual worksheet to validate.")
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "output",
        help="Directory for validation_summary.json and exceptions.csv.",
    )
    args = parser.parse_args()

    registry = CategoryRegistry(PROJECT_ROOT / "config" / "categories.yaml")
    result = CurrentLayoutExtractor(registry).extract(args.workbook, args.sheet)
    report = ImportValidator(registry).validate_current(result)
    summary_path, exceptions_path = write_validation_report(report, args.output)

    print(f"Worksheet: {args.sheet}")
    print(f"Valid: {report.is_valid}")
    print(f"Errors: {len(report.errors)}")
    print(f"Warnings: {len(report.warnings)}")
    print(f"Normalized records: {report.metrics['normalized_record_count']}")
    print(f"Unmapped records: {report.metrics['unknown_record_count']}")
    print(f"Reconciliation difference: {report.metrics['reconciliation_difference']}")
    print(f"Summary: {summary_path}")
    print(f"Exceptions: {exceptions_path}")
    return 0 if report.is_valid else 1


if __name__ == "__main__":
    raise SystemExit(main())
