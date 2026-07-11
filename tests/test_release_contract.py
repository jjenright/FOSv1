from pathlib import Path

from scripts.release_check import static_release_checks
from src.update import build_parser
from src.version import __version__, get_version

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_runtime_version_matches_version_file() -> None:
    assert __version__ == "1.0.0"
    assert get_version() == (PROJECT_ROOT / "VERSION").read_text(encoding="utf-8").strip()


def test_static_release_contract_passes() -> None:
    static_release_checks()


def test_update_cli_supports_private_output_override() -> None:
    args = build_parser().parse_args(
        ["Budget.xlsx", "--output", "private/FOS.xlsx"]
    )
    assert args.workbook == Path("Budget.xlsx")
    assert args.output == Path("private/FOS.xlsx")
