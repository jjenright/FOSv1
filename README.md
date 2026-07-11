# Family Financial Operating System (FOS)

FOS transforms the private historical budget workbook into normalized financial
records, annual KPIs, a current balance-sheet snapshot, an executive dashboard,
and a deterministic financial insight layer.

## Current release

**v0.6.0 — Financial Insights and Action Plan**

The full update produces:

- Normalized historical transactions and income
- Import validation and traceable exceptions
- Annual KPI history and current financial snapshot
- Provisional Financial Pressure Index (FPI v1)
- Executive dashboard with current position and historical trends
- Spending evolution versus the prior year and recent benchmark
- Six evidence-backed financial insights
- Six prioritized, measurable operating actions
- Documented FOS thresholds and limitations

## Run

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

Private outputs are created under `output\` and must not be committed to GitHub.
