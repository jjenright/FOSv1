from pathlib import Path

import pytest

from src.extract import LayoutDetector


CONFIG = Path(__file__).resolve().parents[1] / "config" / "layouts.yaml"


def test_detects_configured_layouts() -> None:
    detector = LayoutDetector(CONFIG)
    assert detector.detect("2010") == "legacy"
    assert detector.detect("2017 (old)") == "transitional"
    assert detector.detect("2025") == "current"


def test_rejects_non_year_support_sheet() -> None:
    detector = LayoutDetector(CONFIG)
    with pytest.raises(ValueError, match="No layout configured"):
        detector.detect("A & L")


def test_configured_sheet_count() -> None:
    detector = LayoutDetector(CONFIG)
    assert len(detector.configured_sheets()) == 19
