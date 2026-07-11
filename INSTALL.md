# Install v0.2.1

1. Extract this package.
2. Copy its contents into the local `FOSv1` repository and replace existing files.
3. Activate the virtual environment.
4. Install dependencies and run the tests.

```powershell
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py
```

Regenerate the private workbook after installing the hotfix:

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx" --sheet 2025
```
