"""Validation for workbook-wide historical imports."""

from __future__ import annotations

import csv
import json
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from src.extract import HistoricalExtractionResult
from src.transform import normalize_category_label
from src.validate.import_validator import (
    ExceptionSummary,
    ImportValidator,
    ValidationIssue,
    ValidationReport,
)


@dataclass(frozen=True, slots=True)
class SheetValidation:
    """Validation report attached to one imported worksheet."""

    sheet_name: str
    layout: str
    report: ValidationReport


@dataclass(frozen=True, slots=True)
class HistoricalValidationReport:
    """Aggregate and per-sheet validation for a historical import."""

    sheets: tuple[SheetValidation, ...]
    aggregate: ValidationReport

    @property
    def is_valid(self) -> bool:
        return self.aggregate.is_valid

    @property
    def errors(self):
        return self.aggregate.errors

    @property
    def warnings(self):
        return self.aggregate.warnings

    @property
    def exceptions(self):
        return self.aggregate.exceptions

    @property
    def metrics(self):
        return self.aggregate.metrics


class HistoricalImportValidator:
    """Validate every annual sheet and create one aggregate report."""

    def __init__(self, validator: ImportValidator) -> None:
        self.validator = validator

    def validate(
        self, extraction: HistoricalExtractionResult
    ) -> HistoricalValidationReport:
        sheet_reports: list[SheetValidation] = []
        aggregate_issues: list[ValidationIssue] = []

        for sheet in extraction.sheets:
            report = self.validator.validate_current(sheet.result)
            sheet_reports.append(
                SheetValidation(
                    sheet_name=sheet.sheet_name,
                    layout=sheet.layout,
                    report=report,
                )
            )
            for issue in report.issues:
                aggregate_issues.append(
                    ValidationIssue(
                        code=issue.code,
                        severity=issue.severity,
                        message=f"[{sheet.sheet_name}] {issue.message}",
                        source_refs=issue.source_refs,
                    )
                )

        exceptions = self._summarize_unknowns(extraction)
        source_total = sum(
            (row.amount for row in extraction.source_rows), Decimal("0")
        )
        normalized_total = sum(
            (record.transaction.amount for record in extraction.records),
            Decimal("0"),
        )
        unknown_total = sum(
            (unknown.amount for unknown in extraction.unknown_categories),
            Decimal("0"),
        )
        difference = source_total - normalized_total - unknown_total

        metrics: dict[str, int | str] = {
            "sheet_count": len(extraction.sheets),
            "excluded_sheet_count": len(extraction.excluded_sheets),
            "period_count": sum(
                len(sheet.result.periods) for sheet in extraction.sheets
            ),
            "source_row_count": len(extraction.source_rows),
            "normalized_record_count": len(extraction.records),
            "unknown_record_count": len(extraction.unknown_categories),
            "unknown_label_count": len(exceptions),
            "source_total": str(source_total),
            "normalized_total": str(normalized_total),
            "unknown_total": str(unknown_total),
            "reconciliation_difference": str(difference),
            "error_count": sum(
                1 for issue in aggregate_issues if issue.severity == "error"
            ),
            "warning_count": sum(
                1 for issue in aggregate_issues if issue.severity == "warning"
            ),
        }

        aggregate = ValidationReport(
            issues=tuple(aggregate_issues),
            exceptions=exceptions,
            metrics=metrics,
        )
        return HistoricalValidationReport(
            sheets=tuple(sheet_reports),
            aggregate=aggregate,
        )

    @staticmethod
    def _summarize_unknowns(
        extraction: HistoricalExtractionResult,
    ) -> tuple[ExceptionSummary, ...]:
        groups: dict[str, dict[str, object]] = defaultdict(
            lambda: {
                "sample_label": "",
                "occurrences": 0,
                "total_amount": Decimal("0"),
                "source_refs": set(),
            }
        )
        for unknown in extraction.unknown_categories:
            normalized = normalize_category_label(unknown.original_name)
            group = groups[normalized]
            if not group["sample_label"]:
                group["sample_label"] = unknown.original_name
            group["occurrences"] = int(group["occurrences"]) + 1
            group["total_amount"] = (
                Decimal(str(group["total_amount"])) + unknown.amount
            )
            refs = group["source_refs"]
            assert isinstance(refs, set)
            refs.add(f"{unknown.source_sheet}!{unknown.source_cell}")

        summaries = [
            ExceptionSummary(
                normalized_label=normalized,
                sample_label=str(group["sample_label"]),
                occurrences=int(group["occurrences"]),
                total_amount=Decimal(str(group["total_amount"])),
                source_refs=tuple(sorted(str(ref) for ref in group["source_refs"])),
            )
            for normalized, group in groups.items()
        ]
        return tuple(
            sorted(
                summaries,
                key=lambda item: (-abs(item.total_amount), item.normalized_label),
            )
        )


def write_historical_validation_report(
    report: HistoricalValidationReport, output_dir: str | Path
) -> tuple[Path, Path]:
    """Write aggregate/per-sheet JSON and a workbook-wide exceptions CSV."""

    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    summary_path = destination / "historical_validation_summary.json"
    exceptions_path = destination / "historical_exceptions.csv"

    payload = {
        "is_valid": report.is_valid,
        "metrics": report.metrics,
        "sheets": [
            {
                "sheet_name": sheet.sheet_name,
                "layout": sheet.layout,
                "is_valid": sheet.report.is_valid,
                "metrics": sheet.report.metrics,
                "issues": [
                    {
                        "code": issue.code,
                        "severity": issue.severity,
                        "message": issue.message,
                        "source_refs": list(issue.source_refs),
                    }
                    for issue in sheet.report.issues
                ],
            }
            for sheet in report.sheets
        ],
    }
    summary_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    with exceptions_path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(
            [
                "normalized_label",
                "sample_label",
                "occurrences",
                "total_amount",
                "source_refs",
            ]
        )
        for exception in report.exceptions:
            writer.writerow(
                [
                    exception.normalized_label,
                    exception.sample_label,
                    exception.occurrences,
                    str(exception.total_amount),
                    "; ".join(exception.source_refs),
                ]
            )

    return summary_path, exceptions_path
