# Verify v0.3.0

Run from the repository root:

```powershell
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Expected private-workbook verification includes:

- 29 tests passed
- 18 official annual worksheets imported
- 412 pay periods extracted
- 5,655 normalized records
- 775 unmapped records retained for review
- $0.00 reconciliation difference
- `2017 (old)` excluded
- historical FOS workbook created successfully

Generate the production output with:

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```
