# Verify FOS v1.0.1

Run from the repository root:

```powershell
.\.venv\Scripts\Activate.ps1
py -m pytest
py scripts\release_check.py
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

Unlike pre-1.0 verification, the production verifier does **not** hardcode a
specific year, row count, account balance or dollar total. It verifies the live
source workbook through these relationships:

- Every annual worksheet has a configured layout
- Archived `(old)` worksheets are excluded
- Source rows equal normalized records plus explicit exceptions
- Source dollars equal normalized dollars plus exception dollars
- Reconciliation difference is exactly zero
- KPI years match imported years
- Net worth equals assets minus liabilities
- FPI score and band are internally consistent
- Six insights and six actions are generated
- Generated table row counts match the live extraction
- Dashboard cards, four charts and takeaways are present
- Structured tables do not overlap worksheet AutoFilters
- The generated XLSX archive is structurally readable

A passing run ends with:

```text
Verification PASSED
```
