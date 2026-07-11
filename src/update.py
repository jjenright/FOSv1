"""Command-line entry point for the production FOS update workflow."""

from __future__ import annotations

import argparse
from pathlib import Path

from src.historical_pipeline import HistoricalPipeline
from src.pipeline import CurrentYearPipeline
from src.version import __version__


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Extract, validate, analyze, and load a private budget workbook "
            "into the Family Financial Operating System."
        )
    )
    parser.add_argument("workbook", type=Path, help="Private source budget workbook.")
    parser.add_argument(
        "--sheet",
        help="Import one current-layout worksheet instead of the full official history.",
    )
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
        if args.sheet:
            result = CurrentYearPipeline(PROJECT_ROOT).run(
                args.workbook,
                sheet_name=args.sheet,
                output_path=args.output,
                fos_version=__version__,
            )
            validation = result.validation
        else:
            result = HistoricalPipeline(PROJECT_ROOT).run(
                args.workbook,
                output_path=args.output,
                fos_version=__version__,
            )
            validation = result.validation
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"ERROR: {exc}")
        return 1

    load = result.load_result
    print(f"FOS v{__version__} update completed successfully.")
    print(f"Workbook: {load.output_path}")
    print(f"Categories: {load.category_rows}")
    print(f"Transactions: {load.transaction_rows}")
    print(f"Income rows: {load.income_rows}")
    print(f"Exceptions: {load.exception_rows}")
    print(f"Warnings: {len(validation.warnings)}")
    print(f"Validation summary: {result.validation_summary_path}")
    print(f"Exceptions report: {result.exceptions_path}")
    if getattr(result, "insight_report", None) is not None:
        report = result.insight_report
        print(f"Latest complete year: {report.latest_year}")
        print(f"Insights generated: {len(report.insights)}")
        print(f"Actions generated: {len(report.actions)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
