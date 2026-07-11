# v0.6.0 Manifest

## Core additions

- `src/insights/engine.py` — spending evolution, insights, and action plan
- `src/insights/__init__.py` — public insight interfaces
- `tests/test_insights_engine.py` — insight-engine unit tests
- `docs/insights_guide.md` — workbook interpretation and operating rules

## Updated integration

- `src/historical_pipeline.py` — calculates the insight report
- `src/load/historical_excel_loader.py` — writes insight sheets and dashboard takeaways
- `src/update.py` — release version 0.6.0
- `scripts/verify.py` — full insight-layer verification

## Generated private outputs

The update command creates these files under `output\`:

- `Financial_Operating_System.xlsx`
- `historical_validation_summary.json`
- `historical_exceptions.csv`

Generated outputs and source budget workbooks are intentionally excluded from Git.
