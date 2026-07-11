# Install v0.3.0

1. Extract this package.
2. Copy all extracted contents into the local `FOSv1` repository and replace existing files.
3. Open a terminal in the repository root.
4. Activate the virtual environment and run the checks.

```powershell
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Generate the historical FOS workbook:

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```
