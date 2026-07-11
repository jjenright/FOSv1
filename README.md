# Family Financial Operating System (FOS)

Current development release: **v0.2.0-alpha.6**

This release loads a validated `2025` current-layout import into a structured
FOS Excel workbook. It creates:

- an import-status dashboard;
- the production category dimension;
- normalized transaction and income fact tables;
- grouped category exceptions;
- validation details; and
- an import audit log with source-workbook traceability.

Unmapped source rows remain visible in the `Exceptions` worksheet and are not
silently assigned to a category.

## Windows verification

From the repository root with the virtual environment activated:

```powershell
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py
```

Verify the complete extract/validate/load pipeline against the private workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Create the private FOS workbook:

```powershell
py scripts\load_current.py "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

The generated workbook and validation reports are written to `output\` and
must not be committed to GitHub.
