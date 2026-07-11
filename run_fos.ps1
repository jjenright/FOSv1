param(
    [Parameter(Mandatory = $true)]
    [string]$Workbook,

    [string]$Output = "",

    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

if (-not (Test-Path $Workbook -PathType Leaf)) {
    throw "Source workbook not found: $Workbook"
}

$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $Python -PathType Leaf)) {
    Write-Host "Creating the local Python environment..."
    py -m venv .venv
}

& $Python -c "import openpyxl, yaml, pytest" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing FOS dependencies..."
    & $Python -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

if (-not $SkipTests) {
    & $Python -m pytest
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

& $Python scripts\verify.py --workbook $Workbook
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

if ($Output) {
    & $Python -m src.update $Workbook --output $Output
} else {
    & $Python -m src.update $Workbook
}
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "FOS update finished. Private outputs are in the output folder."
