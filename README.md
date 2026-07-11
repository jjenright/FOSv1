# Family Financial Operating System (FOS)

FOS transforms the private historical budget workbook into normalized financial
records, annual KPIs, a current financial snapshot, and an Excel reporting workbook.

## Current release

**v0.4.1 — KPI Engine**

The full update now produces:

- Normalized historical transactions and income
- Import validation and exceptions
- Annual KPI history
- Current assets, liabilities, and net worth
- Provisional Financial Pressure Index (FPI v1)
- KPI summary dashboard

## Run

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

The private outputs are created under `output\` and should not be committed to GitHub.
