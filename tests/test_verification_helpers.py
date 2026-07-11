from decimal import Decimal
from types import SimpleNamespace

from openpyxl import Workbook

from scripts.verify import (
    verify_current_snapshot_consistency,
    verify_current_snapshot_sheet,
)


def make_snapshot(vehicle_loan: str = "54000"):
    return SimpleNamespace(
        total_assets=Decimal("1198646"),
        total_liabilities=Decimal("571500"),
        net_worth=Decimal("627146"),
        fpi_score=Decimal("37.6"),
        fpi_band="Moderate",
        vehicle_loan_balance=Decimal(vehicle_loan),
    )


def test_snapshot_consistency_accepts_updated_liability_balance() -> None:
    verify_current_snapshot_consistency(make_snapshot())


def test_snapshot_sheet_uses_dynamic_snapshot_values() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet["B7"] = 627146
    worksheet["B16"] = 37.6
    worksheet["B17"] = "Moderate"

    verify_current_snapshot_sheet(worksheet, make_snapshot())
