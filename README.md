# Family Financial Operating System (FOS)

FOS converts the private family budget workbook into a repeatable financial
operating system. It imports the historical annual worksheets, normalizes the
transactions, validates every dollar, calculates KPIs, builds an executive
Excel dashboard, and generates a deterministic action plan.

## Production release

**v1.0.1 — Family Financial Operating System**

The generated workbook includes:

- An executive dashboard with current position and historical trends
- Full 2008–present historical import using configured annual layouts
- Source-sheet and source-cell traceability
- Annual income, spending, savings, debt and flexibility KPIs
- Current assets, liabilities, net worth and liquidity snapshot
- Provisional Financial Pressure Index (FPI v1)
- Spending evolution versus the prior year and recent benchmark
- Six evidence-backed insights and six measurable actions
- Validation, exceptions and import logs

FOS does not upload financial data or require a subscription. Processing is
performed locally on the computer running the command.

## Daily command

From PowerShell in the repository root:

```powershell
.\run_fos.ps1 -Workbook "C:\path\to\Budget-Jason-original.xlsx"
```

The standard Python command remains available:

```powershell
py -m src.update "C:\path\to\Budget-Jason-original.xlsx"
```

Private outputs are written to `output\` and are excluded from GitHub.

## Documentation

- [Installation](INSTALL.md)
- [Verification](VERIFY.md)
- [User guide](docs/user_guide.md)
- [Architecture](docs/architecture.md)
- [KPI definitions](docs/kpi_definitions.md)
- [Insight operating rules](docs/insights_guide.md)
- [Maintenance](docs/maintenance.md)
- [Privacy and backups](docs/privacy_and_backup.md)
