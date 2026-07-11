# Verify v0.2.0-alpha.5

From the repository root with the virtual environment activated:

```powershell
py -m pytest
py scripts\verify.py
```

Expected result: **20 tests pass** and the script ends with
`Verification PASSED`.

To validate against the private historical workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Expected private-workbook checks:

- 19 annual worksheets;
- 26 pay periods in 2025;
- 497 source financial rows;
- 463 normalized records;
- 34 unmapped records across 28 labels;
- source total of `$432,079.48`; and
- reconciliation difference of `0.00`.

To create detailed reports:

```powershell
py scripts\validate_current.py "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

The command writes:

- `output\validation_summary.json`
- `output\exceptions.csv`
