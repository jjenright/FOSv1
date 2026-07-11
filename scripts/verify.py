"""Verify the FOS v0.2.0-alpha.3 components."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.extract import LayoutDetector  # noqa: E402
from src.models import Category, ImportResult, ImportSession, Transaction  # noqa: E402,F401
from src.transform import CategoryRegistry  # noqa: E402


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
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover - dependency check
        raise RuntimeError("openpyxl is required for workbook category verification.") from exc

    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
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


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify FOS v0.2.0-alpha.3")
    parser.add_argument("--workbook", type=Path, help="Optional private Budget workbook path.")
    args = parser.parse_args()

    detector = LayoutDetector(PROJECT_ROOT / "config" / "layouts.yaml")
    registry = CategoryRegistry(PROJECT_ROOT / "config" / "categories.yaml")

    print("FOS v0.2.0-alpha.3 verification")
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

    print("Verification PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
