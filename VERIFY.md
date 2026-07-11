# Verify v0.2.1

Run from the repository root:

```powershell
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

The generated `Financial_Operating_System.xlsx` should open without Excel repairing or removing `FactTransactionsTable`.
