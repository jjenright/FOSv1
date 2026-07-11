# Changelog

## v0.2.1

- Fixed an invalid overlapping worksheet AutoFilter on `FactTransactions`.
- Preserved the table-owned AutoFilter and structured Excel table.
- Added a regression test to prevent Excel repair warnings.

## v0.2.0

- Added the integrated `CurrentYearPipeline`.
- Replaced the placeholder `src.update` entry point with a production command.
- Added one-command extraction, validation, reporting, and Excel loading.
- Added end-to-end integration tests.
- Added full private-workbook pipeline verification.
- Marked the current-layout Data Engine complete.

## v0.2.0-alpha.6

- Added the Excel FOS loader for validated current-layout imports.
- Added Dashboard, Import_Log, DimCategory, FactTransactions, FactIncome, Exceptions, and Validation worksheets.
- Added Excel tables, financial number formatting, source-cell traceability, and a first import summary chart.
- Added a command-line load utility and automated loader tests.

## v0.2.0-alpha.5

- Added source-row and amount reconciliation.
- Added grouped exceptions and validation reports.

## v0.2.0-alpha.4

- Added the production current-layout extractor.

## v0.2.0-alpha.3

- Added the production category dictionary and registry.

## v0.2.0-alpha.2

- Added runnable core models and layout detection.
