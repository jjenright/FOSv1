# Install FOS v1.1.1

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
py -m pytest
Unblock-File -Path .\run_fos.ps1
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx"
```

For routine updates after tests have passed:

```powershell
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx" -SkipTests
```

Keep source and generated spreadsheets private.


## Visa sheet requirements

Copy the cleaned `Import_2025` sheet into the private source workbook. Keep the columns `TransactionID`, `TransactionDate`, `Amount ($)`, `TransactionType`, and `CategoryID`. The sheet name is flexible because FOS detects the header structure.
