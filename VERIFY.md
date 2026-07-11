# Verify v0.2.0

## Repository verification

```powershell
py -m pytest
py scripts\verify.py
```

## Private-workbook verification

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Expected private-workbook results include:

- 19 configured annual worksheets;
- 26 pay periods in the 2025 worksheet;
- 497 source financial rows;
- 463 normalized records;
- 34 unmapped records across 28 labels;
- zero reconciliation difference;
- a successfully generated FOS workbook; and
- `Verification PASSED`.

## Run the production update

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

Expected output files:

```text
output\Financial_Operating_System.xlsx
output\validation_summary.json
output\exceptions.csv
```
