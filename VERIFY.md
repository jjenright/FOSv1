# Verify FOS v1.1.1

```powershell
.\.venv\Scripts\Activate.ps1
py -m pytest
py scripts\release_check.py
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

Verification confirms zero-dollar reconciliation, dynamic KPIs, decision opportunities, LOC acceleration, three twelve-month forecast scenarios, Financial DNA output, filterable explorer tables, workbook integrity, and privacy exclusions.


A successful private-workbook verification reports `visa_imported_record_count` and `visa_imported_total` when a detailed Visa sheet is present. Duplicate transaction IDs cause verification to fail rather than silently removing spending.
