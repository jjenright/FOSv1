# Install v0.5.0

1. Close any open FOS workbook in Excel.
2. Extract this ZIP.
3. Copy all extracted files into the local `FOSv1` repository and replace existing files.
4. Keep the private budget workbook outside the GitHub repository.
5. From the repository root in PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m pytest
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Generate the current FOS workbook:

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

Open `output\Financial_Operating_System.xlsx` and use the first worksheet,
`Dashboard`, as the executive view. Outputs under `output\` are excluded from GitHub.
