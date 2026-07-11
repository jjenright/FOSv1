# Verify v0.2.0-alpha.4

From the repository root with the virtual environment activated:

```powershell
py -m pytest
py scripts\verify.py
```

Expected result: 16 tests pass and the script ends with `Verification PASSED`.

To test the extractor against the private historical workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

The workbook check validates the 19 annual worksheets and confirms that the 2025
extractor finds 26 pay periods, 463 normalized records, and 34 unmapped one-off
records while reconciling the section totals used by this release.
