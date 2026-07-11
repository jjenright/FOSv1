# Family Financial Operating System (FOS)

Current release: **v0.3.0 — Historical Import**

This release imports the official annual budget history into one normalized FOS
workbook. It supports the legacy and current worksheet patterns found in the
source workbook, validates every annual sheet, preserves source-cell
traceability, and excludes the archived `2017 (old)` sheet to prevent duplicate
2017 data.

## Windows setup and verification

From the repository root with the virtual environment activated:

```powershell
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py
```

Verify against the private budget workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

## Generate the full historical FOS workbook

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

The command writes these private outputs under `output\`:

- `Financial_Operating_System.xlsx`
- `historical_validation_summary.json`
- `historical_exceptions.csv`

To retain the previous one-sheet workflow, provide `--sheet`, for example:

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```

Source workbooks and generated outputs must not be committed to GitHub.

## v0.3.0 boundary

This release completes the historical data import. Unmapped one-off labels are
preserved in the Exceptions outputs for later review; they are never silently
guessed. KPI calculations and financial insights are the next release milestone.
