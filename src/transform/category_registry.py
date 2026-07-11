"""Load and resolve the production category dictionary."""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any

import yaml

from src.models import Category

_ORDINAL_IN_PARENS = re.compile(r"\s*\(\s*\d+(?:st|nd|rd|th)\s*\)\s*", re.IGNORECASE)
_TRAILING_ORDINAL = re.compile(r"\s*[-–—]\s*\d+(?:st|nd|rd|th)\s*$", re.IGNORECASE)


def normalize_category_label(label: str) -> str:
    """Return a stable comparison key for a workbook category label."""

    normalized = unicodedata.normalize("NFKC", label).strip()
    normalized = normalized.replace("’", "'").replace("‘", "'")
    normalized = normalized.replace("–", "-").replace("—", "-")
    normalized = _ORDINAL_IN_PARENS.sub(" ", normalized)
    normalized = _TRAILING_ORDINAL.sub("", normalized)
    normalized = " ".join(normalized.split())
    return normalized.casefold()


class CategoryRegistry:
    """Read ``categories.yaml`` and resolve workbook aliases to categories."""

    def __init__(self, config_file: str | Path) -> None:
        self.config_file = Path(config_file)
        if not self.config_file.is_file():
            raise FileNotFoundError(f"Category configuration not found: {self.config_file}")

        with self.config_file.open("r", encoding="utf-8") as stream:
            loaded: Any = yaml.safe_load(stream)

        if not isinstance(loaded, dict) or not isinstance(loaded.get("categories"), list):
            raise ValueError("Category configuration must contain a 'categories' list.")

        self.version = str(loaded.get("version", "unknown"))
        self._categories_by_id: dict[str, dict[str, Any]] = {}
        self._aliases: dict[str, str] = {}
        self._load(loaded["categories"])

    def _load(self, entries: list[dict[str, Any]]) -> None:
        for entry in entries:
            category_id = str(entry.get("category_id", "")).strip()
            if not category_id:
                raise ValueError("Every category must define a category_id.")
            if category_id in self._categories_by_id:
                raise ValueError(f"Duplicate category_id: {category_id}")

            aliases = entry.get("aliases", [])
            if not isinstance(aliases, list) or not aliases:
                raise ValueError(f"Category '{category_id}' must define at least one alias.")

            self._categories_by_id[category_id] = entry
            for alias in aliases:
                key = normalize_category_label(str(alias))
                previous = self._aliases.get(key)
                if previous and previous != category_id:
                    raise ValueError(
                        f"Alias '{alias}' maps to both '{previous}' and '{category_id}'."
                    )
                self._aliases[key] = category_id

    def find(self, original_name: str) -> Category | None:
        """Return a matching category or ``None`` when the alias is unknown."""

        category_id = self._aliases.get(normalize_category_label(original_name))
        if category_id is None:
            return None
        return self._to_model(category_id, original_name)

    def lookup(self, original_name: str) -> Category:
        """Return a matching category or raise ``KeyError``."""

        category = self.find(original_name)
        if category is None:
            raise KeyError(f"Unknown financial category: {original_name}")
        return category

    def _to_model(self, category_id: str, original_name: str) -> Category:
        entry = self._categories_by_id[category_id]
        return Category(
            category_id=category_id,
            original_name=original_name,
            display_name=str(entry["display_name"]),
            category_type=str(entry["category_type"]),
            master_category=str(entry["master_category"]),
            financial_purpose=str(entry["financial_purpose"]),
            weekly_budget=bool(entry["weekly_budget"]),
            controllable=str(entry["controllable"]),
            active=bool(entry.get("active", True)),
        )

    def category_count(self) -> int:
        return len(self._categories_by_id)

    def alias_count(self) -> int:
        return len(self._aliases)

    def category_ids(self) -> tuple[str, ...]:
        return tuple(self._categories_by_id)

    def entries(self) -> tuple[dict[str, Any], ...]:
        """Return configured category entries in source order."""

        return tuple(dict(entry) for entry in self._categories_by_id.values())
