# Install v0.2.0-alpha.5 on Windows

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

If the virtual environment does not exist yet, create it first:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
```

If PowerShell blocks activation, run this once in the same terminal, then
activate again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Do not commit the personal budget workbook. Keep it outside the repository or in
the ignored `workbook/` folder. Generated reports in `output/` are also ignored.
