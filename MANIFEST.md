# v0.4.1 Manifest

## Purpose

Verification hotfix for mutable balances in the private `A & L` worksheet.

## Modified

- `scripts/verify.py` — removes hardcoded net-worth and FPI expectations.
- `tests/test_verification_helpers.py` — covers the corrected $54,000 vehicle-loan case.
- `src/update.py`, `src/pipeline.py`, and `src/historical_pipeline.py` — use release version 0.4.1.
- Release documentation and version files.

## Behaviour

- Net worth is validated as current assets minus current liabilities.
- Generated `Current_Snapshot` cells are checked against the calculated snapshot.
- FPI is checked for a valid 0–100 range and the correct score band.
- No KPI formulas or financial classifications changed.
