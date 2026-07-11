# Verify v0.2.0-alpha.3

From the repository root with the virtual environment activated:

```powershell
py -m pytest
py scripts\verify.py
```

Expected result: all tests pass and the script ends with `Verification PASSED`.

To confirm the dictionary against the private historical workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

The workbook verification checks the annual sheet configuration and several critical
category mappings used by the current layout.
