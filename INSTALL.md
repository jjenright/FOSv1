# Install v0.2.0

1. Extract the release ZIP.
2. Copy all extracted files into the local `FOSv1` repository.
3. Allow Windows to replace existing project files.
4. Keep the private budget workbook outside the repository.
5. From the repository root, activate the existing virtual environment:

```powershell
.\.venv\Scripts\Activate.ps1
```

6. Install or refresh dependencies:

```powershell
py -m pip install -r requirements.txt
```

7. Run the automated tests:

```powershell
py -m pytest
```

8. Run project verification:

```powershell
py scripts\verify.py
```
