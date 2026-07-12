"""End-to-end historical workbook import pipeline."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from src.extract import HistoricalWorkbookExtractor, LayoutDetector
from src.insights import InsightReport, InsightsEngine
from src.decision import DecisionIntelligenceEngine, DecisionReport
from src.kpi import KPIEngine
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
    insight_report: InsightReport | None = None
    decision_report: DecisionReport | None = None


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
        fos_version: str = "1.1.0",
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

        kpi_engine = KPIEngine(self.registry)
        annual_kpis = kpi_engine.calculate_annual(extraction)
        current_snapshot = (
            kpi_engine.calculate_current_snapshot(source, annual_kpis)
            if any(item.coverage_status == "Complete" for item in annual_kpis)
            else None
        )
        insight_report = (
            InsightsEngine(self.registry).analyze(extraction, annual_kpis, current_snapshot)
            if current_snapshot is not None else None
        )
        decision_report = (
            DecisionIntelligenceEngine(self.registry, self.project_root / "config" / "decision_intelligence.yaml").analyze(extraction, annual_kpis, current_snapshot)
            if current_snapshot is not None else None
        )

        load_result = HistoricalExcelFOSLoader(self.registry).load_historical(
            extraction,
            validation,
            destination,
            source_workbook=source,
            fos_version=fos_version,
            annual_kpis=annual_kpis,
            current_snapshot=current_snapshot,
            insight_report=insight_report,
            decision_report=decision_report,
        )
        return HistoricalPipelineResult(
            load_result=load_result,
            validation=validation,
            validation_summary_path=summary_path,
            exceptions_path=exceptions_path,
            insight_report=insight_report,
            decision_report=decision_report,
        )
