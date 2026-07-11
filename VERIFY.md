# Verify v0.6.0

Run from the repository root:

```powershell
.\.venv\Scripts\Activate.ps1
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

Expected release checks include:

- 37 automated tests passing
- 18 official annual worksheets imported
- 6,430 source rows reconciled
- 5,655 normalized records
- $0.00 reconciliation difference
- 2025 selected as the latest complete year
- 2022–2024 selected as the recent benchmark
- Six financial insights and six action-plan items
- Four executive-dashboard charts
- Valid structured tables with no overlapping worksheet AutoFilter
