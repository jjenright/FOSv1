# Changelog

## v0.2.0-alpha.5

- Added source-row and amount reconciliation for current-layout imports.
- Added duplicate period, duplicate source-cell, and category-ID integrity checks.
- Added grouped unmapped-category exceptions with counts, totals, and source references.
- Added JSON validation summaries and CSV exceptions reports.
- Added a command-line validation script and automated validator tests.
- Verified the 2025 import reconciles 497 source rows and $432,079.48 with zero difference.

## v0.2.0-alpha.4

- Added the production current-layout extractor for the 2025 annual worksheet.
- Detects all pay-period blocks using the worksheet header structure.
- Separates income, transfers, variable/irregular expenses, and fixed expenses.
- Preserves source worksheet and cell traceability on every normalized record.
- Reports unknown non-zero categories explicitly rather than guessing.
- Added automated extractor tests and private-workbook reconciliation checks.

## v0.2.0-alpha.3

- Added the production `categories.yaml` dictionary based on the historical workbook.
- Preserved the approved Costco, Canadian Tire, Walmart, and Amazon mappings.
- Added category-label normalization for ordinal due dates and punctuation variants.
- Added `CategoryRegistry` with explicit unknown-category handling.
- Added automated category dictionary tests.
- Extended verification to check critical mappings against the private workbook.

## v0.2.0-alpha.2

- Rebuilt Commit 1 as runnable core domain models.
- Rebuilt Commit 2 as configuration-driven layout detection.
- Preserved the existing `src/` repository structure.
- Added automated tests and a verification script.
- Added Windows installation and verification instructions.
- Added safeguards to keep private financial workbooks out of GitHub.
