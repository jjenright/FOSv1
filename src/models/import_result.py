"""Import-result model used by extractors, validators, and loaders."""

from dataclasses import dataclass, field


@dataclass(slots=True)
class ImportResult:
    """Summary of records, warnings, and errors produced by an import."""

    records_imported: int = 0
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def success(self) -> bool:
        """Return True when the import completed without errors."""

        return not self.errors
