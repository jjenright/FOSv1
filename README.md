# Family Financial Operating System (FOS)

Current development release: **v0.2.0-alpha.2**

This release rebuilds the first two development commits as a runnable package:

1. Core financial domain models.
2. Configuration-driven worksheet layout detection.

The private household budget workbook is intentionally excluded from GitHub.

## Windows setup

Open a terminal in the repository root and run:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py
```

To verify the layout configuration against your private workbook:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Expected result: `Verification PASSED`.

See `INSTALL.md` and `VERIFY.md` for more detail.
