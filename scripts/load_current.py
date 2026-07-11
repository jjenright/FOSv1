"""Extract, validate, and load one current-layout year into an FOS workbook."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.extract import CurrentLayoutExtractor  # noqa: E402
from src.load import ExcelFOSLoader  # noqa: E402
from src.transform import CategoryRegistry  # noqa: E402
from src.validate import ImportValidator, write_validation_report  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(
        description='Load a validated current-layout worksheet into an FOS workbook.'
    )
    parser.add_argument('workbook', type=Path, help='Private source budget workbook.')
    parser.add_argument('--sheet', default='2025', help='Annual worksheet to import.')
    parser.add_argument(
        '--output',
        type=Path,
        default=PROJECT_ROOT / 'output' / 'Financial_Operating_System.xlsx',
        help='Destination FOS workbook path.',
    )
    args = parser.parse_args()

    registry = CategoryRegistry(PROJECT_ROOT / 'config' / 'categories.yaml')
    extraction = CurrentLayoutExtractor(registry).extract(args.workbook, args.sheet)
    validation = ImportValidator(registry).validate_current(extraction)
    write_validation_report(validation, args.output.parent)

    if not validation.is_valid:
        for issue in validation.errors:
            print(f'ERROR {issue.code}: {issue.message}')
        return 1

    result = ExcelFOSLoader(registry).load_current(
        extraction,
        validation,
        args.output,
        source_workbook=args.workbook,
        source_sheet=args.sheet,
        fos_version='0.2.0-alpha.6',
    )
    print(f'Created: {result.output_path}')
    print(f'Categories: {result.category_rows}')
    print(f'Transactions: {result.transaction_rows}')
    print(f'Income rows: {result.income_rows}')
    print(f'Exceptions: {result.exception_rows}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
