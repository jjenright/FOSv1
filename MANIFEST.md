# v0.2.1 Manifest

## Modified

- `src/load/excel_loader.py` — removes overlapping worksheet AutoFilter.
- `tests/test_excel_loader.py` — regression coverage for table/filter validity.
- Runtime version and release documentation files.

## Behaviour

The `FactTransactions` Excel table retains its own AutoFilter; no second worksheet-level AutoFilter is written.
