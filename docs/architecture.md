# FOS Architecture

FOS is a local, deterministic Python pipeline. The private budget workbook is
the source of truth; the generated FOS workbook is an analytical output and can
be recreated at any time.

## Processing flow

1. **Layout detection** — `config/layouts.yaml` selects the correct extractor
   for each official annual worksheet.
2. **Extraction** — current and legacy extractors read source rows while
   retaining worksheet and cell references.
3. **Normalization** — `config/categories.yaml` maps source labels to the
   approved category dictionary. Unknown labels remain explicit exceptions.
4. **Validation** — row counts and dollar totals reconcile source records,
   normalized records and exceptions.
5. **KPI calculation** — annual performance and the current balance-sheet
   snapshot are calculated from normalized records and `A & L`.
6. **Insight generation** — documented operating rules produce spending
   evolution, six insights and six actions.
7. **Workbook loading** — the Excel loader creates the dashboard, fact tables,
   dimensions, definitions, validation and action sheets.

## Main modules

- `src/extract/` — layout detection and workbook extraction
- `src/transform/` — category registry and normalization
- `src/validate/` — reconciliation and exceptions
- `src/kpi/` — annual and current financial KPIs
- `src/insights/` — deterministic decision layer
- `src/load/` — Excel output generation
- `src/historical_pipeline.py` — full production orchestration
- `src/update.py` — command-line entry point

## Design boundaries

- The source workbook is never modified.
- Unknown categories are never silently guessed.
- Generated outputs are disposable and reproducible.
- FOS calculations are deterministic and do not require an AI service.
- The provisional FPI is an internal operating indicator, not a credit score or
  regulated financial measure.
