# v1.1.1 Manifest

- `src/extract/visa.py` — detailed Visa table detection and extraction
- `src/extract/historical.py` — merge Visa detail into the matching annual year
- `src/validate/historical_validator.py` — Visa import and duplicate metrics
- `src/load/excel_loader.py` — transaction date and merchant traceability
- `src/load/decision_excel_writer.py` — detailed Spending Explorer rows
- `scripts/verify.py` — single-pass private workbook verification
- `tests/test_visa_import.py` — Visa integration regression tests
- `docs/release_notes_v1.1.1.md`

Generated spreadsheets and reports remain private under `output\`.
