# Install v0.2.0-alpha.2 on Windows

1. Extract this package.
2. Copy all files and folders into the root of your cloned `FOSv1` repository.
3. Allow Windows to overwrite files when prompted.
4. Open the repository folder in VS Code.
5. Open **Terminal → New Terminal**.
6. Run:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
```

If PowerShell blocks activation, run this once in the same terminal, then activate again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Do not copy the personal budget workbook into the repository. Keep it elsewhere or in the ignored `workbook/` folder.

## Commit 3

Copy this package over the repository root, allow file replacement, then run the commands in `VERIFY.md`.
