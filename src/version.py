"""Single source of truth for the installed FOS release version."""

from __future__ import annotations

from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
VERSION_FILE = PROJECT_ROOT / "VERSION"


def get_version() -> str:
    """Return the semantic version stored in the repository VERSION file."""
    try:
        version = VERSION_FILE.read_text(encoding="utf-8").strip()
    except OSError as exc:
        raise RuntimeError(f"Unable to read FOS version file: {VERSION_FILE}") from exc
    if not version:
        raise RuntimeError(f"FOS version file is empty: {VERSION_FILE}")
    return version


__version__ = get_version()
