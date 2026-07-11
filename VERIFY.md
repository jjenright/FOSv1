# Verify v0.2.0-alpha.2

From the repository root with the virtual environment activated:

```powershell
py -m pytest
py scripts\verify.py
```

Expected test result:

```text
6 passed
```

Expected verification result:

```text
FOS v0.2.0-alpha.2 verification
- Core models: OK
- 2010 layout: legacy
- 2017 (old) layout: transitional
- 2025 layout: current
Verification PASSED
```

Optional workbook verification:

```powershell
py scripts\verify.py --workbook "C:\path\to\Budget-Jason-original.xlsx"
```
