# v0.3.0 Manifest

## Added

- `src/extract/legacy.py` — legacy pay-period worksheet extraction.
- `src/extract/historical.py` — workbook-wide extraction and archive exclusion.
- `src/validate/historical_validator.py` — aggregate and per-sheet validation.
- `src/load/historical_excel_loader.py` — full-history Excel FOS output.
- `src/historical_pipeline.py` — end-to-end historical pipeline.
- Historical extractor and pipeline tests.

## Modified

- `src/extract/current.py` — skips balance/total control rows and handles duplicate period labels.
- `config/categories.yaml` — maps legacy `Visa` payment labels.
- `src/update.py` — imports the full history by default; `--sheet` remains available.
- `scripts/verify.py` — verifies the full workbook import and output tables.
- Release documentation and version files.

## Historical import rules

- Imports 2008–2014, 2016, 2017, and 2018–2026.
- Excludes `2017 (old)` to prevent double counting.
- Preserves unknown labels in Exceptions instead of guessing.
- Reconciles every extracted source row to either a normalized record or an exception.
