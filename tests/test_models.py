from datetime import datetime
from decimal import Decimal

from src.models import Category, ImportResult, ImportSession, Transaction


def test_category_model() -> None:
    category = Category(
        category_id="F001",
        original_name="Costco",
        display_name="Groceries",
        category_type="Expense",
        master_category="Food",
        financial_purpose="Essential Living",
        weekly_budget=True,
        controllable="Yes",
    )
    assert category.category_id == "F001"
    assert category.active is True


def test_transaction_model() -> None:
    transaction = Transaction(
        year=2025,
        period="P01",
        category_id="F001",
        description="Costco",
        amount=Decimal("123.45"),
        source_sheet="2025",
        source_cell="B20",
    )
    assert transaction.amount == Decimal("123.45")


def test_import_models() -> None:
    session = ImportSession("Budget.xlsx", datetime(2026, 7, 11, 9, 0, 0))
    result = ImportResult(records_imported=10)
    assert session.workbook_name == "Budget.xlsx"
    assert result.success is True

    result.errors.append("example")
    assert result.success is False
