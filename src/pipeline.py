"""End-to-end FOS data-engine pipeline for one current-layout worksheet."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.extract import CurrentLayoutExtractor, LayoutDetector
from src.load import ExcelFOSLoader, LoadResult
from src.transform import CategoryRegistry
from src.validate import ImportValidator, ValidationReport, write_validation_report


@dataclass(frozen=True, slots=True)
class PipelineResult:
    """Summary returned after a successful FOS update."""

    load_result: LoadResult
    validation: ValidationReport
    validation_summary_path: Path
    exceptions_path: Path


class CurrentYearPipeline:
    """Extract, validate, and load one configured current-layout worksheet."""

    def __init__(self, project_root: str | Path) -> None:
        self.project_root = Path(project_root)
        self.detector = LayoutDetector(self.project_root / "config" / "layouts.yaml")
        self.registry = CategoryRegistry(self.project_root / "config" / "categories.yaml")

    def run(
        self,
        workbook_path: str | Path,
        *,
        sheet_name: str = "2025",
        output_path: str | Path | None = None,
        fos_version: str = "0.2.0",
    ) -> PipelineResult:
        """Run the complete current-layout import pipeline."""

        source = Path(workbook_path)
        if not source.is_file():
            raise FileNotFoundError(f"Workbook not found: {source}")

        detected_layout = self.detector.detect(sheet_name)
        if detected_layout != "current":
            raise ValueError(
                f"Worksheet '{sheet_name}' uses layout '{detected_layout}', "
                "but v0.2.0 supports current-layout imports only."
            )

        destination = Path(output_path) if output_path else (
            self.project_root / "output" / "Financial_Operating_System.xlsx"
        )

        extraction = CurrentLayoutExtractor(self.registry).extract(source, sheet_name)
        validation = ImportValidator(self.registry).validate_current(extraction)
        summary_path, exceptions_path = write_validation_report(
            validation, destination.parent
        )

        if not validation.is_valid:
            messages = "; ".join(issue.message for issue in validation.errors)
            raise ValueError(f"Import validation failed: {messages}")

        load_result = ExcelFOSLoader(self.registry).load_current(
            extraction,
            validation,
            destination,
            source_workbook=source,
            source_sheet=sheet_name,
            fos_version=fos_version,
        )

        return PipelineResult(
            load_result=load_result,
            validation=validation,
            validation_summary_path=summary_path,
            exceptions_path=exceptions_path,
        )
