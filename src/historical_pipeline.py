"""End-to-end historical workbook import pipeline."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.extract import HistoricalWorkbookExtractor, LayoutDetector
from src.load import HistoricalExcelFOSLoader, LoadResult
from src.transform import CategoryRegistry
from src.validate import (
    HistoricalImportValidator,
    HistoricalValidationReport,
    ImportValidator,
    write_historical_validation_report,
)


@dataclass(frozen=True, slots=True)
class HistoricalPipelineResult:
    """Summary returned after a successful historical FOS update."""

    load_result: LoadResult
    validation: HistoricalValidationReport
    validation_summary_path: Path
    exceptions_path: Path


class HistoricalPipeline:
    """Extract, validate, and load all official annual worksheets."""

    def __init__(self, project_root: str | Path) -> None:
        self.project_root = Path(project_root)
        self.detector = LayoutDetector(self.project_root / "config" / "layouts.yaml")
        self.registry = CategoryRegistry(self.project_root / "config" / "categories.yaml")

    def run(
        self,
        workbook_path: str | Path,
        *,
        output_path: str | Path | None = None,
        fos_version: str = "0.3.0",
        sheets: tuple[str, ...] | list[str] | None = None,
    ) -> HistoricalPipelineResult:
        source = Path(workbook_path)
        if not source.is_file():
            raise FileNotFoundError(f"Workbook not found: {source}")

        destination = Path(output_path) if output_path else (
            self.project_root / "output" / "Financial_Operating_System.xlsx"
        )
        extraction = HistoricalWorkbookExtractor(
            self.detector, self.registry
        ).extract(source, sheets=sheets)
        validation = HistoricalImportValidator(
            ImportValidator(self.registry)
        ).validate(extraction)
        summary_path, exceptions_path = write_historical_validation_report(
            validation, destination.parent
        )

        if not validation.is_valid:
            messages = "; ".join(issue.message for issue in validation.errors)
            raise ValueError(f"Historical import validation failed: {messages}")

        load_result = HistoricalExcelFOSLoader(self.registry).load_historical(
            extraction,
            validation,
            destination,
            source_workbook=source,
            fos_version=fos_version,
        )
        return HistoricalPipelineResult(
            load_result=load_result,
            validation=validation,
            validation_summary_path=summary_path,
            exceptions_path=exceptions_path,
        )
