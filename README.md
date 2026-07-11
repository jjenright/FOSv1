# Family Financial Operating System (FOS)

Current development release: **v0.2.0-alpha.4**

This release adds the first production workbook extractor. It reads the `2025`
current-layout worksheet, detects all 26 pay-period blocks, maps known rows through
the production category dictionary, and separates the normalized records into:

- income;
- transfers and debt payments;
- variable and irregular expenses; and
- fixed expenses.

Every record retains its source worksheet and cell. Non-zero rows that are not in
the category dictionary are returned as explicit unknown-category records instead
of being silently assigned.

## Windows verification

From the repository root with the virtual environment activated:

```powershell
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py
```

Optional verification against the private workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

The private workbook must not be committed to GitHub.
