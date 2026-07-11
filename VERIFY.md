# Verify v0.4.1

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
- Current net worth reconciles to the current values in the source `A & L` sheet.
- FPI and its band are validated dynamically.
- Verification prints `Verification PASSED`.
