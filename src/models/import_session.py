"""Import-session metadata model."""

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True, slots=True)
class ImportSession:
    """Metadata describing one execution of the FOS import process."""

    workbook_name: str
    started_at: datetime
    workbook_version: str = "unknown"
