# Changelog

## v0.5.0

- Rebuilt the first worksheet as an executive financial dashboard.
- Added current-position and latest complete-year KPI cards.
- Added validation, data-coverage, unmapped-amount, and trend-eligibility status.
- Added annual cash-flow, ratio-trend, balance-sheet, and spending-mix charts.
- Excluded partial or lower-coverage years from trend charts where comparison-eligible years exist.
- Added hidden chart-helper data, print layout, and dashboard regression tests.
- Preserved all v0.4.1 KPI, classification, and reconciliation logic.

## v0.4.1

- Removed hardcoded current net-worth and FPI values from private-workbook verification.
- Verification now reconciles net worth to current `A & L` assets and liabilities.
- Generated `Current_Snapshot` values are checked against the calculated snapshot dynamically.
- Added regression tests for updated liability balances.

## v0.4.0

- Added annual KPI calculations across the normalized 2008–2026 history.
- Added current balance-sheet snapshot from the `A & L` worksheet.
- Added Net Worth, Fixed Cost Ratio, Wealth-Building Rate, Savings Velocity,
  Financial Flexibility, category ratios, data coverage, and comparison flags.
- Added provisional Financial Pressure Index (FPI v1) using existing workbook data.
- Added `Annual_KPIs`, `Current_Snapshot`, and `KPI_Definitions` output sheets.
- Updated the FOS dashboard with current position and latest complete-year metrics.
- Added KPI tests and end-to-end verification against the source workbook.

## v0.3.0

- Added full historical workbook import for 2008–2026.
- Excluded archived `2017 (old)` from official history.
- Reconciled all extracted source rows to normalized records or exceptions.
