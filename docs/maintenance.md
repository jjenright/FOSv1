# FOS Maintenance

## Adding a category alias

Edit `config/categories.yaml`, add the approved source label to the appropriate
category alias list, then run:

```powershell
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Unknown labels should remain exceptions until their financial meaning is clear.

## Adding a new annual worksheet

1. Create the annual sheet in the source workbook using an existing supported
   layout.
2. Add the worksheet name to `config/layouts.yaml`.
3. Run verification. The production verifier dynamically accepts new totals and
   row counts when reconciliation remains valid.

## Updating balances

Update the `A & L` worksheet directly. Net worth and FPI verification are
calculated from the live balances and do not rely on hardcoded amounts.

## Release maintenance

- Update the root `VERSION` file.
- Update `CHANGELOG.md`, `README.md`, `INSTALL.md` and `VERIFY.md`.
- Run `py scripts\release_check.py --workbook "..."`.
- Run the complete test suite before committing.
