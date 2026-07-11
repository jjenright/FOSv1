# Family Financial Operating System (FOS)

Current development release: **v0.2.0-alpha.5**

This release validates the normalized `2025` current-layout import before it can
be loaded into the FOS. It adds:

- source-row and dollar reconciliation;
- duplicate pay-period and source-cell checks;
- category-ID integrity checks;
- grouped exceptions for unmapped workbook labels;
- JSON validation summaries and CSV exceptions reports; and
- automated tests against synthetic data and the private workbook baseline.

The validator treats unmapped categories as warnings, not silent guesses. An
import is valid when it has no structural or reconciliation errors.

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

To write the detailed validation files for the `2025` worksheet:

```powershell
py scripts\validate_current.py "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

This writes `output\validation_summary.json` and `output\exceptions.csv`. The
private workbook and generated outputs must not be committed to GitHub.
