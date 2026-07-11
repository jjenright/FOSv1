"""Static and optional private-workbook release checks for FOS."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from scripts.verify import verify_private_workbook  # noqa: E402
from src.version import __version__  # noqa: E402

SEMVER = re.compile(r"^\d+\.\d+\.\d+$")
REQUIRED_FILES = (
    "README.md",
    "INSTALL.md",
    "VERIFY.md",
    "CHANGELOG.md",
    "VERSION",
    "LICENSE",
    "config/categories.yaml",
    "config/layouts.yaml",
    "src/update.py",
    "scripts/verify.py",
    "run_fos.ps1",
    "docs/user_guide.md",
    "docs/architecture.md",
    "docs/privacy_and_backup.md",
    "docs/maintenance.md",
)
PRIVACY_PATTERNS = (
    "private_data/",
    "workbook/",
    "output/*",
    "sample_data/*",
    "*.xlsx",
    "*.xlsm",
    "*.xls",
)


def static_release_checks() -> None:
    if not SEMVER.fullmatch(__version__):
        raise ValueError(f"VERSION is not semantic versioning: {__version__}")
    missing = [name for name in REQUIRED_FILES if not (PROJECT_ROOT / name).is_file()]
    if missing:
        raise ValueError("Required release files are missing: " + ", ".join(missing))

    gitignore = (PROJECT_ROOT / ".gitignore").read_text(encoding="utf-8")
    missing_privacy = [pattern for pattern in PRIVACY_PATTERNS if pattern not in gitignore]
    if missing_privacy:
        raise ValueError(
            "Privacy exclusions are missing from .gitignore: "
            + ", ".join(missing_privacy)
        )

    for document in ("README.md", "INSTALL.md", "VERIFY.md"):
        text = (PROJECT_ROOT / document).read_text(encoding="utf-8")
        if __version__ not in text:
            raise ValueError(f"{document} does not reference release {__version__}.")


def main() -> int:
    parser = argparse.ArgumentParser(description=f"Check FOS v{__version__} release readiness")
    parser.add_argument("--workbook", type=Path, help="Optional private workbook for end-to-end checks.")
    args = parser.parse_args()

    try:
        static_release_checks()
        print(f"FOS v{__version__} static release checks: PASS")
        if args.workbook:
            verify_private_workbook(args.workbook)
            print("Private-workbook release checks: PASS")
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"ERROR: {exc}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
