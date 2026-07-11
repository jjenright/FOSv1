# Install v0.2.0-alpha.6 on Windows

1. Extract this package.
2. Copy all files and folders into the root of your cloned `FOSv1` repository.
3. Allow Windows to overwrite files when prompted.
4. Open the repository folder in VS Code.
5. Open **Terminal → New Terminal**.
6. Activate the existing environment and install the requirements:

```powershell
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

If the virtual environment does not exist:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
```

Run the automated tests and package verification:

```powershell
py -m pytest
py scripts\verify.py
```

Keep the personal budget workbook outside the repository or in the ignored
`workbook/` folder. Generated workbooks and reports in `output/` are ignored.
