# Family Financial Operating System (FOS)

FOS transforms the private historical budget workbook into normalized financial
records, annual KPIs, a current financial snapshot, and an Excel reporting workbook.

## Current release

**v0.5.0 — Executive Dashboard**

The full update produces:

- Normalized historical transactions and income
- Import validation and exceptions
- Annual KPI history and current balance-sheet snapshot
- Provisional Financial Pressure Index (FPI v1)
- Executive dashboard with current-position and latest-year KPI cards
- Comparison-eligible cash-flow and ratio trends
- Current balance-sheet and latest-year spending-mix charts

## Run

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

The private outputs are created under `output\` and should not be committed to GitHub.
