# Verify v0.5.0

Run from the repository root:

```powershell
.\.venv\Scripts\Activate.ps1
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Expected result:

- All tests pass.
- 18 official annual worksheets import.
- Historical reconciliation difference is zero.
- Current net worth reconciles to the current `A & L` balances.
- The executive dashboard contains KPI cards and four charts.
- Hidden chart-helper columns remain hidden.
- Excel tables remain structurally valid and no overlapping AutoFilter is created.
- Verification prints `Verification PASSED`.
