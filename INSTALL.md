# Install v0.6.0

1. Extract the release ZIP.
2. Copy its contents into the local `FOSv1` repository, replacing existing files.
3. Open PowerShell in the repository root.
4. Activate the virtual environment and install dependencies:

```powershell
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

5. Run tests and verification:

```powershell
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

6. Generate the private FOS workbook:

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

The workbook and reports are written to `output\`. They contain private financial
data and are excluded by `.gitignore`.
