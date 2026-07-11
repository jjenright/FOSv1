# Verify v0.2.0-alpha.6

From the repository root with the virtual environment activated:

```powershell
py -m pytest
py scripts\verify.py
```

Expected result: **22 tests pass** and the script ends with
`Verification PASSED`.

Verify the complete pipeline against the private historical workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Expected private-workbook checks:

- 19 annual worksheets;
- 26 pay periods in 2025;
- 497 source financial rows;
- 463 normalized records;
- 34 unmapped records across 28 labels;
- source total of `$432,079.48`;
- reconciliation difference of `0.00`;
- 378 non-income transaction rows loaded;
- 85 income rows loaded; and
- a valid seven-sheet FOS Excel workbook.

Create the workbook yourself:

```powershell
py scripts\load_current.py "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

Generated files:

- `output\Financial_Operating_System.xlsx`
- `output\validation_summary.json`
- `output\exceptions.csv`
