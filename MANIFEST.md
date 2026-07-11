# v1.0.1 Manifest

## Production engine

- `src/update.py` — complete update command
- `src/version.py` — centralized release version
- `src/historical_pipeline.py` — full import, KPI, insight and workbook pipeline
- `scripts/verify.py` — dynamic private-workbook verification
- `scripts/release_check.py` — static and optional end-to-end release checks
- `run_fos.ps1` — Windows one-command workflow

## Configuration

- `config/categories.yaml` — 64-category production dictionary and aliases
- `config/layouts.yaml` — annual worksheet layout assignments

## Documentation

- `README.md`, `INSTALL.md`, `VERIFY.md`, `CHANGELOG.md`
- `docs/user_guide.md`
- `docs/architecture.md`
- `docs/privacy_and_backup.md`
- `docs/maintenance.md`
- `docs/release_notes_v1.0.1.md`
- KPI, dashboard, insight, data dictionary and financial DNA guides

## Private generated outputs

The update command creates under `output\`:

- `Financial_Operating_System.xlsx`
- `historical_validation_summary.json`
- `historical_exceptions.csv`

The output folder contents and source workbooks are excluded from Git.
