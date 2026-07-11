"""Command-line entry point for the FOS v0.2.0 data engine."""

from __future__ import annotations

import argparse
from pathlib import Path

from src.pipeline import CurrentYearPipeline


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Extract, validate, and load a current-layout budget worksheet."
    )
    parser.add_argument("workbook", type=Path, help="Private source budget workbook.")
    parser.add_argument("--sheet", default="2025", help="Annual worksheet to import.")
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "output" / "Financial_Operating_System.xlsx",
        help="Destination FOS workbook path.",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()

    try:
        result = CurrentYearPipeline(PROJECT_ROOT).run(
            args.workbook,
            sheet_name=args.sheet,
            output_path=args.output,
            fos_version="0.2.0",
        )
    except (FileNotFoundError, ValueError) as exc:
        print(f"ERROR: {exc}")
        return 1

    load = result.load_result
    validation = result.validation
    print("FOS update completed successfully.")
    print(f"Workbook: {load.output_path}")
    print(f"Categories: {load.category_rows}")
    print(f"Transactions: {load.transaction_rows}")
    print(f"Income rows: {load.income_rows}")
    print(f"Exceptions: {load.exception_rows}")
    print(f"Warnings: {len(validation.warnings)}")
    print(f"Validation summary: {result.validation_summary_path}")
    print(f"Exceptions report: {result.exceptions_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
