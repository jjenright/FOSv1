"""Validation and exceptions reporting for normalized workbook imports."""

from __future__ import annotations

import csv
import json
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from src.extract import CurrentExtractionResult
from src.transform import CategoryRegistry, normalize_category_label

SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"


@dataclass(frozen=True, slots=True)
class ValidationIssue:
    """One validation finding."""

    code: str
    severity: str
    message: str
    source_refs: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class ExceptionSummary:
    """Grouped summary for one unmapped workbook label."""

    normalized_label: str
    sample_label: str
    occurrences: int
    total_amount: Decimal
    source_refs: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class ValidationReport:
    """Validation result for one current-layout extraction."""

    issues: tuple[ValidationIssue, ...]
    exceptions: tuple[ExceptionSummary, ...]
    metrics: dict[str, int | str]

    @property
    def errors(self) -> tuple[ValidationIssue, ...]:
        return tuple(issue for issue in self.issues if issue.severity == SEVERITY_ERROR)

    @property
    def warnings(self) -> tuple[ValidationIssue, ...]:
        return tuple(issue for issue in self.issues if issue.severity == SEVERITY_WARNING)

    @property
    def is_valid(self) -> bool:
        return not self.errors


class ImportValidator:
    """Validate current-layout extraction results before loading them."""

    def __init__(self, category_registry: CategoryRegistry) -> None:
        self.category_registry = category_registry

    def validate_current(self, result: CurrentExtractionResult) -> ValidationReport:
        issues: list[ValidationIssue] = []
        exceptions = self._summarize_unknowns(result)

        if not result.periods:
            issues.append(
                ValidationIssue(
                    code="NO_PERIODS",
                    severity=SEVERITY_ERROR,
                    message="No pay periods were extracted.",
                )
            )

        duplicate_periods = self._duplicates(result.periods)
        if duplicate_periods:
            issues.append(
                ValidationIssue(
                    code="DUPLICATE_PERIODS",
                    severity=SEVERITY_ERROR,
                    message="Duplicate pay-period labels were extracted: "
                    + ", ".join(duplicate_periods),
                )
            )

        classified_refs = [
            f"{record.transaction.source_sheet}!{record.transaction.source_cell}"
            for record in result.records
        ] + [
            f"{unknown.source_sheet}!{unknown.source_cell}"
            for unknown in result.unknown_categories
        ]
        duplicate_refs = self._duplicates(classified_refs)
        if duplicate_refs:
            issues.append(
                ValidationIssue(
                    code="DUPLICATE_SOURCE_ROWS",
                    severity=SEVERITY_ERROR,
                    message="A source row was classified more than once.",
                    source_refs=duplicate_refs,
                )
            )

        configured_ids = set(self.category_registry.category_ids())
        missing_ids = sorted(
            {
                record.transaction.category_id
                for record in result.records
                if record.transaction.category_id not in configured_ids
            }
        )
        if missing_ids:
            issues.append(
                ValidationIssue(
                    code="UNKNOWN_CATEGORY_IDS",
                    severity=SEVERITY_ERROR,
                    message="Extracted records reference category IDs absent from the dictionary: "
                    + ", ".join(missing_ids),
                )
            )

        source_count = len(result.source_rows)
        classified_count = len(result.records) + len(result.unknown_categories)
        if source_count != classified_count:
            issues.append(
                ValidationIssue(
                    code="ROW_RECONCILIATION_FAILED",
                    severity=SEVERITY_ERROR,
                    message=(
                        f"Source rows ({source_count}) do not reconcile to normalized and "
                        f"unmapped rows ({classified_count})."
                    ),
                )
            )

        source_total = sum((row.amount for row in result.source_rows), Decimal("0"))
        classified_total = sum(
            (record.transaction.amount for record in result.records), Decimal("0")
        ) + sum(
            (unknown.amount for unknown in result.unknown_categories), Decimal("0")
        )
        difference = source_total - classified_total
        if difference != 0:
            issues.append(
                ValidationIssue(
                    code="AMOUNT_RECONCILIATION_FAILED",
                    severity=SEVERITY_ERROR,
                    message=(
                        f"Source total ({source_total}) does not reconcile to normalized and "
                        f"unmapped total ({classified_total}); difference {difference}."
                    ),
                )
            )

        if exceptions:
            issues.append(
                ValidationIssue(
                    code="UNMAPPED_CATEGORIES",
                    severity=SEVERITY_WARNING,
                    message=(
                        f"{len(result.unknown_categories)} source rows across "
                        f"{len(exceptions)} category labels require mapping review."
                    ),
                    source_refs=tuple(
                        reference
                        for exception in exceptions
                        for reference in exception.source_refs
                    ),
                )
            )

        metrics: dict[str, int | str] = {
            "period_count": len(result.periods),
            "source_row_count": source_count,
            "normalized_record_count": len(result.records),
            "unknown_record_count": len(result.unknown_categories),
            "unknown_label_count": len(exceptions),
            "source_total": str(source_total),
            "normalized_total": str(
                sum((record.transaction.amount for record in result.records), Decimal("0"))
            ),
            "unknown_total": str(
                sum((unknown.amount for unknown in result.unknown_categories), Decimal("0"))
            ),
            "reconciliation_difference": str(difference),
            "error_count": sum(1 for issue in issues if issue.severity == SEVERITY_ERROR),
            "warning_count": sum(
                1 for issue in issues if issue.severity == SEVERITY_WARNING
            ),
        }

        return ValidationReport(
            issues=tuple(issues),
            exceptions=exceptions,
            metrics=metrics,
        )

    @staticmethod
    def _duplicates(values: Iterable[str]) -> tuple[str, ...]:
        seen: set[str] = set()
        duplicates: set[str] = set()
        for value in values:
            if value in seen:
                duplicates.add(value)
            seen.add(value)
        return tuple(sorted(duplicates))

    @staticmethod
    def _summarize_unknowns(
        result: CurrentExtractionResult,
    ) -> tuple[ExceptionSummary, ...]:
        groups: dict[str, dict[str, object]] = defaultdict(
            lambda: {
                "sample_label": "",
                "occurrences": 0,
                "total_amount": Decimal("0"),
                "source_refs": set(),
            }
        )
        for unknown in result.unknown_categories:
            normalized = normalize_category_label(unknown.original_name)
            group = groups[normalized]
            if not group["sample_label"]:
                group["sample_label"] = unknown.original_name
            group["occurrences"] = int(group["occurrences"]) + 1
            group["total_amount"] = Decimal(str(group["total_amount"])) + unknown.amount
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
                key=lambda item: (-item.total_amount, item.normalized_label),
            )
        )


def write_validation_report(report: ValidationReport, output_dir: str | Path) -> tuple[Path, Path]:
    """Write a JSON validation summary and CSV exceptions report."""

    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    summary_path = destination / "validation_summary.json"
    exceptions_path = destination / "exceptions.csv"

    summary = {
        "is_valid": report.is_valid,
        "metrics": report.metrics,
        "issues": [
            {
                "code": issue.code,
                "severity": issue.severity,
                "message": issue.message,
                "source_refs": list(issue.source_refs),
            }
            for issue in report.issues
        ],
    }
    summary_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")

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
