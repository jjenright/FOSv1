# Family Financial Operating System (FOS)

**v1.1.1 — Visa Transaction Integration**

FOS imports the private family budget workbook, reconciles every dollar, calculates KPIs, and generates a local Excel operating system.

v1.1.1 keeps the v1.1.0 Decision Intelligence tools and adds:

- Automatic detection of the copied detailed Visa transaction sheet
- Import of categorized purchases, refunds, interest, and fees
- Exclusion of credit-card payments as transfers
- Transaction ID duplicate protection
- Transaction dates and merchant descriptions in FactTransactions and Spending Explorer
- Visa import counts and dollar totals in validation output

Run from PowerShell:

```powershell
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx"
```

Decision assumptions are stored in `config\decision_intelligence.yaml`. Private outputs are written to `output\` and excluded from Git.


## Detailed Visa transactions

When the private source workbook contains a copied `Import_2025`-style sheet, FOS detects it from the headers and merges the categorized rows into the matching annual year. Credit-card payments are excluded as transfers. Transaction IDs must be unique. The annual sheet should not separately contain the same individual purchases.
