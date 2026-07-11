# Family Financial Operating System (FOS)

Current release: **v0.2.1 — Excel table hotfix**

This release completes the first end-to-end FOS pipeline for a current-layout
annual worksheet. One command now:

1. reads the private budget workbook;
2. detects the configured worksheet layout;
3. extracts and normalizes income, transfers, variable expenses, and fixed expenses;
4. validates row counts, source amounts, category IDs, and source-cell uniqueness;
5. writes validation and exceptions reports; and
6. creates the structured FOS Excel workbook.

Unmapped source rows remain visible in the `Exceptions` worksheet and are not
silently assigned to a category.

## Windows setup and verification

From the repository root with the virtual environment activated:

```powershell
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py
```

Verify the complete pipeline against the private workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

## Run the FOS update

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

Equivalent convenience command:

```powershell
py scripts\update_fos.py "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

The generated workbook and reports are written to `output\` and must not be
committed to GitHub.

## v0.2.1 boundary

This release supports the configured **current** worksheet layout. Historical
legacy and transitional extraction is the next release milestone.
