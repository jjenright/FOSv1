from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

from src.insights import InsightsEngine
from src.kpi import AnnualKPI, CurrentSnapshot
from src.transform import CategoryRegistry

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def annual(
    year: int,
    *,
    income: str,
    fixed: str,
    variable: str,
    core: str,
    wealth: str,
) -> AnnualKPI:
    true_income = Decimal(income)
    fixed_expenses = Decimal(fixed)
    variable_expenses = Decimal(variable)
    wealth_building = Decimal(wealth)
    known = fixed_expenses + variable_expenses
    return AnnualKPI(
        year=year,
        coverage_status="Complete",
        pay_periods=26,
        true_income=true_income,
        cash_flow_adjustments=Decimal("0"),
        fixed_expenses=fixed_expenses,
        variable_irregular_expenses=variable_expenses,
        known_operating_expenses=known,
        core_spending=Decimal(core),
        lifestyle_spending=Decimal("3000"),
        wealth_building=wealth_building,
        targeted_debt_reduction=Decimal("0"),
        excluded_transfers=Decimal("0"),
        unknown_amount=Decimal("1000"),
        fixed_cost_ratio=fixed_expenses / true_income,
        known_expense_ratio=known / true_income,
        wealth_building_rate=wealth_building / true_income,
        savings_velocity=wealth_building / true_income,
        financial_flexibility=(true_income - fixed_expenses) / true_income,
        housing_ratio=Decimal("0.20"),
        transportation_ratio=Decimal("0.08"),
        food_ratio=Decimal("0.04"),
        lifestyle_ratio=Decimal("0.02"),
        data_coverage_ratio=Decimal("0.96"),
        comparison_eligible=True,
    )


def snapshot() -> CurrentSnapshot:
    return CurrentSnapshot(
        source_sheet="A & L",
        latest_complete_year=2025,
        total_assets=Decimal("1000000"),
        total_liabilities=Decimal("500000"),
        net_worth=Decimal("500000"),
        financial_assets=Decimal("125000"),
        savings_cash=Decimal("6000"),
        home_equity=Decimal("450000"),
        mortgage_balance=Decimal("450000"),
        line_of_credit_balance=Decimal("17500"),
        vehicle_loan_balance=Decimal("32500"),
        emergency_fund_months=Decimal("0.9"),
        unsecured_debt_ratio=Decimal("0.08"),
        fpi_score=Decimal("38.0"),
        fpi_band="Moderate",
    )


def test_insight_report_uses_latest_year_and_three_year_benchmark() -> None:
    registry = CategoryRegistry(PROJECT_ROOT / "config" / "categories.yaml")
    engine = InsightsEngine(registry)
    history = (
        annual(2022, income="200000", fixed="35000", variable="10000", core="42000", wealth="18000"),
        annual(2023, income="220000", fixed="40000", variable="12000", core="48000", wealth="20000"),
        annual(2024, income="230000", fixed="50000", variable="30000", core="70000", wealth="24000"),
        annual(2025, income="215000", fixed="60000", variable="28000", core="82000", wealth="23000"),
    )

    report = engine.analyze(SimpleNamespace(sheets=()), history, snapshot())

    assert report.latest_year == 2025
    assert report.prior_year == 2024
    assert report.benchmark_years == (2022, 2023, 2024)
    assert report.emergency_target_amount == Decimal("20500")
    assert report.emergency_fund_gap == Decimal("14500")
    assert report.actions[0].area == "Debt"
    assert report.actions[0].target_value == Decimal("0")
    assert report.actions[1].area == "Liquidity"
    assert report.insights[0].theme == "Liquidity"


def test_spending_evolution_signals_income_decline_and_fixed_cost_increase() -> None:
    registry = CategoryRegistry(PROJECT_ROOT / "config" / "categories.yaml")
    engine = InsightsEngine(registry)
    history = (
        annual(2024, income="230000", fixed="50000", variable="30000", core="70000", wealth="24000"),
        annual(2025, income="215000", fixed="60000", variable="28000", core="82000", wealth="23000"),
    )

    report = engine.analyze(SimpleNamespace(sheets=()), history, snapshot())
    by_id = {item.metric_id: item for item in report.spending_evolution}

    assert by_id["income"].direction == "Decrease"
    assert by_id["income"].signal == "Watch"
    assert by_id["fixed_expenses"].direction == "Increase"
    assert by_id["fixed_expenses"].signal == "Watch"
