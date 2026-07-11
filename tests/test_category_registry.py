from pathlib import Path

import pytest

from src.transform import CategoryRegistry, normalize_category_label


CONFIG = Path(__file__).resolve().parents[1] / "config" / "categories.yaml"


def test_dictionary_loads_and_has_expected_size() -> None:
    registry = CategoryRegistry(CONFIG)
    assert registry.version == "0.4.0"
    assert registry.category_count() >= 50
    assert registry.alias_count() >= 150


def test_user_approved_merchant_mappings() -> None:
    registry = CategoryRegistry(CONFIG)
    assert registry.lookup("Costco").category_id == "FOD001"
    assert registry.lookup("Canadian Tire").category_id == "HOU012"
    assert registry.lookup("Walmart").category_id == "HOU012"
    assert registry.lookup("Amazon").category_id == "HOU012"


def test_due_dates_are_normalized() -> None:
    registry = CategoryRegistry(CONFIG)
    assert registry.lookup("Mortgage (21st)").category_id == "HOU001"
    assert registry.lookup("Home insurance (17th)").category_id == "HOU004"
    assert registry.lookup("Enbridge (30th)").category_id == "HOU006"
    assert registry.lookup("Bank fee (8th)").category_id == "FIN003"


def test_core_workbook_aliases() -> None:
    registry = CategoryRegistry(CONFIG)
    assert registry.lookup("JE's Pay").category_id == "INC001"
    assert registry.lookup("Truck (2nd Friday)").category_id == "TRA001"
    assert registry.lookup("Visa payment").category_id == "TRF005"
    assert registry.lookup("Questrade TFSA").category_id == "TRF002"
    assert registry.lookup("Therapist L").category_id == "HLT004"


def test_unknown_category_is_not_silently_assigned() -> None:
    registry = CategoryRegistry(CONFIG)
    assert registry.find("Completely new merchant") is None
    with pytest.raises(KeyError, match="Unknown financial category"):
        registry.lookup("Completely new merchant")


def test_label_normalization() -> None:
    assert normalize_category_label("  Mortgage (21st) ") == "mortgage"
    assert normalize_category_label("RBC – CAN Additional Payment") == (
        "rbc - can additional payment"
    )
