"""Workbook-layout detection driven by ``config/layouts.yaml``."""

from pathlib import Path
from typing import Any

import yaml


class LayoutDetector:
    """Resolve a worksheet name to a configured workbook layout."""

    def __init__(self, config_file: str | Path) -> None:
        self.config_file = Path(config_file)
        if not self.config_file.is_file():
            raise FileNotFoundError(f"Layout configuration not found: {self.config_file}")

        with self.config_file.open("r", encoding="utf-8") as stream:
            loaded: Any = yaml.safe_load(stream)

        if not isinstance(loaded, dict) or not isinstance(loaded.get("layouts"), dict):
            raise ValueError("Layout configuration must contain a 'layouts' mapping.")

        self.layouts: dict[str, dict[str, Any]] = loaded["layouts"]
        self._sheet_to_layout = self._build_sheet_index()

    def _build_sheet_index(self) -> dict[str, str]:
        index: dict[str, str] = {}
        for layout_name, layout_config in self.layouts.items():
            sheets = layout_config.get("sheets", [])
            if not isinstance(sheets, list):
                raise ValueError(f"Layout '{layout_name}' must define a list of sheets.")

            for sheet_name in sheets:
                normalized = str(sheet_name).strip()
                if normalized in index:
                    raise ValueError(
                        f"Worksheet '{normalized}' is assigned to more than one layout."
                    )
                index[normalized] = layout_name
        return index

    def detect(self, sheet_name: str) -> str:
        """Return the configured layout name for ``sheet_name``."""

        normalized = sheet_name.strip()
        try:
            return self._sheet_to_layout[normalized]
        except KeyError as exc:
            raise ValueError(f"No layout configured for worksheet '{sheet_name}'.") from exc

    def configured_sheets(self) -> tuple[str, ...]:
        """Return all configured worksheet names in deterministic order."""

        return tuple(self._sheet_to_layout)
