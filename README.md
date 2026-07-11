# Family Financial Operating System (FOS)

Current development release: **v0.2.0-alpha.3**

This release contains the first production category dictionary generated from the
historical household budget workbook. It preserves the approved merchant mappings:

- Costco → Groceries
- Canadian Tire → Household Supplies
- Walmart → Household Supplies
- Amazon → Household Supplies

It also adds a category registry that normalizes recurring due-date labels such as
`Mortgage (21st)` and reports unknown categories rather than silently guessing.

## Windows verification

From the repository root with the virtual environment activated:

```powershell
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py
```

Optional verification against the private workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

The private workbook must not be committed to GitHub.
