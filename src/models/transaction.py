"""Transaction domain model for normalized workbook entries."""

from dataclasses import dataclass
from datetime import date
from decimal import Decimal


@dataclass(frozen=True, slots=True)
class Transaction:
    """A normalized financial event extracted from a source workbook."""

    year: int
    period: str
    category_id: str
    description: str
    amount: Decimal
    source_sheet: str
    source_cell: str
    transaction_date: date | None = None
