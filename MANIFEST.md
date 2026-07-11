# v0.5.0 Manifest

## Purpose

Provide the first executive dashboard for the Family Financial Operating System.

## Modified

- `src/load/historical_excel_loader.py` — executive dashboard cards, status strip,
  hidden chart datasets, four charts, print layout, and dashboard notes.
- `tests/test_historical_pipeline.py` — dashboard structure, chart, hidden-column,
  and AutoFilter regression checks.
- `src/update.py`, `src/pipeline.py`, `src/historical_pipeline.py`, and
  `scripts/verify.py` — release version 0.5.0.
- `docs/dashboard_guide.md` — dashboard interpretation guide.
- Release documentation and version files.

## Dashboard content

- Current position: net worth, liquid savings, total debt, emergency fund, and FPI.
- Latest complete year: income, known expenses, wealth building, savings velocity,
  and financial flexibility.
- Status strip: validation, coverage, unmapped amount, and comparison eligibility.
- Charts: annual cash flow, key ratios, current balance sheet, and spending mix.

## Calculation integrity

No KPI formula, category classification, import, or reconciliation logic changed.
