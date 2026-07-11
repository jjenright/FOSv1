# FOS User Guide

## Routine use

1. Update the private budget workbook as usual.
2. Close the workbook in Excel.
3. Run:

```powershell
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx" -SkipTests
```

4. Open `output\Financial_Operating_System.xlsx`.
5. Review the dashboard, insights, action plan and exceptions.

## Workbook sheets

- **Dashboard** — current position, latest complete-year performance, data
  quality, trends and executive takeaways.
- **Insights** — ranked observations with evidence, implications and actions.
- **Action_Plan** — measurable priorities, targets, gaps and status.
- **Spending_Evolution** — latest year versus prior year and recent benchmark.
- **Annual_KPIs** — one row per imported year with all calculated measures.
- **Current_Snapshot** — assets, liabilities, net worth, liquidity and FPI.
- **DimYear / DimCategory** — analytical dimensions.
- **FactTransactions / FactIncome** — normalized records with source-cell
  traceability.
- **Exceptions** — unmapped labels retained for review.
- **Validation / Import_Log** — reconciliation and run details.
- **KPI_Definitions / Insight_Definitions** — calculation and threshold notes.

## Interpreting exceptions

Exceptions are not import failures. They are source labels that were preserved
but not assigned to a normalized category. Review recurring or material labels
first, then add approved mappings to `config/categories.yaml` and rerun FOS.

## Important limitations

- Partial years are flagged and excluded from comparable trend analysis where
  appropriate.
- Results depend on the completeness and classification of the source workbook.
- The dashboard is for household planning and does not replace tax, legal,
  investment or credit advice.
