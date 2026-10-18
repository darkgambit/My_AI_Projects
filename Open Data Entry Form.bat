@echo off
title Data Entry Form Launcher
echo.
echo Starting Data Entry Form...
echo.

REM Try to open the HTML form in the default browser
if exist "D:\.claude\VBA_Code\DataEntryForm.html" (
    start "" "D:\.claude\VBA_Code\DataEntryForm.html"
    echo Data Entry Form opened successfully!
    echo.
    echo You can now:
    echo   - Use all 4 tabbed sheets (Farmer Profile, Seedling Care, Transplant Volume, IPM & Topping)
    echo   - Add, Delete, Update, and Clear records
    echo   - Records are managed in your browser
    echo.
    echo NOTE: To save data to Excel, copy/paste from the form to the Excel sheets.
    pause
) else (
    echo ERROR: HTML form file not found!
    echo Expected at: D:\.claude\VBA_Code\DataEntryForm.html
    pause
)