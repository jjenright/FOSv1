# Family Financial Operating System (FOS)

**v1.1.0 — Decision Intelligence**

FOS imports the private family budget workbook, reconciles every dollar, calculates KPIs, and generates a local Excel operating system.

v1.1.0 adds:

- Direct discretionary-spending opportunities with 25%, 50%, and 100% reduction scenarios
- Merchant intelligence and review suggestions without silently changing validated mappings
- LOC payoff scenarios using configurable rate and weekly-payment assumptions
- Twelve-month current-plan, focused-reduction, and maximum-reduction forecasts
- Financial DNA and expanded category-history analysis
- A filterable Spending Explorer with year, pay-period, category, and purpose drill-down
- Decision Centre and direct-choice cards on the Dashboard

Run from PowerShell:

```powershell
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Decision assumptions are stored in `config\decision_intelligence.yaml`. Private outputs are written to `output\` and excluded from Git.
