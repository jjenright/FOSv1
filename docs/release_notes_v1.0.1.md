# FOS v1.0.1 Release Notes

This hotfix resolves empty chart frames on the Dashboard in Microsoft Excel.

## Cause

Dashboard chart datasets are stored in hidden helper columns P:AB. The charts were using Excel's default setting to plot visible cells only, so Excel omitted the hidden data.

## Fix

All four Dashboard charts now explicitly include data from hidden cells. Verification and automated tests enforce this setting.
