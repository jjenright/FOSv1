# Verify FOS v1.1.0

```powershell
.\.venv\Scripts\Activate.ps1
py -m pytest
py scripts\release_check.py
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

Verification confirms zero-dollar reconciliation, dynamic KPIs, decision opportunities, LOC acceleration, three twelve-month forecast scenarios, Financial DNA output, filterable explorer tables, workbook integrity, and privacy exclusions.
