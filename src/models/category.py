"""Category domain model for normalized financial records."""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class Category:
    """A normalized financial category used by the FOS data engine."""

    category_id: str
    original_name: str
    display_name: str
    category_type: str
    master_category: str
    financial_purpose: str
    weekly_budget: bool
    controllable: str
    active: bool = True
