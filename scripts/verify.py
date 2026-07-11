"""Verify the rebuilt Commit 1 and Commit 2 components.

Optionally pass the source workbook path to compare its annual worksheet names
with ``config/layouts.yaml`` without reading financial cell data.
"""

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


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify FOS v0.2.0-alpha.2")
    parser.add_argument(
        "--workbook",
        type=Path,
        help="Optional path to the private Budget workbook.",
    )
    args = parser.parse_args()

    detector = LayoutDetector(PROJECT_ROOT / "config" / "layouts.yaml")

    print("FOS v0.2.0-alpha.2 verification")
    print("- Core models: OK")
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

    print("Verification PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
