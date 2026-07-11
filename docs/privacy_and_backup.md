# Privacy and Backup

FOS is designed to process the family workbook locally. The repository contains
code and configuration, not private financial records.

## Files that must remain private

- The source budget workbook
- Generated FOS workbooks
- Validation JSON files
- Exception CSV files
- Any exported screenshots containing account balances

The default `output\`, `private_data\`, `workbook\`, and `sample_data\`
contents are excluded by `.gitignore`.

## Recommended backup practice

- Keep the source workbook in a backed-up private folder such as OneDrive with
  appropriate account security.
- Retain periodic dated copies of the source workbook before major structural
  changes.
- Treat the generated FOS workbook as reproducible output rather than the only
  copy of financial history.
- Before every Git commit, run `git status` and confirm no `.xlsx`, `.csv`, or
  validation JSON containing private information is staged.
