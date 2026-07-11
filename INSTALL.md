# Install FOS v1.0.0

## First-time setup

1. Install Python 3 and Git for Windows.
2. Clone or open the local `FOSv1` repository.
3. Open PowerShell in the repository root.
4. Create the local environment and install dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

5. Run the automated tests:

```powershell
py -m pytest
```

## Generate the FOS workbook

The simplest production command is:

```powershell
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx"
```

The launcher verifies the source workbook and regenerates the complete FOS
workbook. For a faster routine update after the release has already been tested:

```powershell
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx" -SkipTests
```

Direct Python usage:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

The private outputs are written to `output\`:

- `Financial_Operating_System.xlsx`
- `historical_validation_summary.json`
- `historical_exceptions.csv`

Do not move the private workbook or generated output files into Git-tracked
folders. The standard locations are already protected by `.gitignore`.
