# Changelog

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
