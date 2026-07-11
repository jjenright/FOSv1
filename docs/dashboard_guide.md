# Executive Dashboard Guide

The Dashboard is the primary FOS view. It summarizes calculated outputs without
introducing a second calculation layer.

## Current Position

- **Net Worth:** assets minus liabilities from `A & L`.
- **Liquid Savings:** the `Savings` balance in `A & L`.
- **Total Debt:** all liabilities recorded in `A & L`.
- **Emergency Fund:** liquid savings divided by monthly core spending.
- **FPI:** provisional Financial Pressure Index from 0 to 100.

## Latest Complete-Year Performance

The dashboard selects the latest year with at least 24 recorded pay periods. It
shows true income, known operating expenses, wealth building, savings velocity,
and financial flexibility.

## Trend Charts

Trend charts use years marked `ComparisonEligible` in `Annual_KPIs`. When no year
meets that condition, the dashboard falls back to complete years and then to all
available years. Partial years remain visible in `Annual_KPIs` for reference.

## Data-quality strip

Review data coverage and unmapped amounts before treating a year as fully
classified. Unmapped source rows remain preserved in `Exceptions`.
