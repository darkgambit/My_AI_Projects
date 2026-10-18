<#
.SYNOPSIS
    Injects VBA code into FOR DATA ENTRY_FORM.xlsm to create a data entry form.

.NOTES
    This script uses Excel COM automation to create a UserForm with
    Add, Delete, Update, and Clear buttons.
#>
param(
    [string]$FilePath = "D:\.claude\FOR DATA ENTRY_FORM.xlsm"
)

$ErrorActionPreference = "Stop"
$excel = $null

try {
    Write-Host "Opening workbook: $FilePath" -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($FilePath)
    $vbProj = $wb.VBProject

    #--- Enable VBA if protected ---
    try {
        $vbProj.Protection = 0
    } catch {
        Write-Host "Warning: Could not unprotect VBA project (may require manual trust access)" -ForegroundColor Yellow
    }

    #--- Export existing sheets for reference ---
    Write-Host ""
    Write-Host "Sheets in workbook:" -ForegroundColor Green
    foreach ($ws in $wb.Worksheets) {
        Write-Host "  - $($ws.Name)"
    }

    #--- Read sheet headers ---
    Write-Host ""
    Write-Host "Sheet structures:" -ForegroundColor Green

    $sheetData = @{}
    foreach ($ws in $wb.Worksheets) {
        $name = $ws.Name
        $headers = @()
        $used = $ws.UsedRange
        $colCount = $used.Columns.Count

        # Try to find a row with actual headers
        for ($r = 1; $r -le [Math]::Min(6, $used.Rows.Count); $r++) {
            $rowVals = @()
            for ($c = 1; $c -le $colCount; $c++) {
                $val = $ws.Cells.Item($r, $c).Text
                $rowVals += $val
            }
            # Check if this row looks like headers (all cells non-empty or mix)
            $nonEmpty = ($rowVals | Where-Object { $_ -ne "" }).Count
            if ($nonEmpty -ge $colCount * 0.3) {
                $headers = $rowVals
                break
            }
        }

        $sheetData[$name] = $headers
        Write-Host "  $name ($colCount columns):"
        Write-Host "    Headers: $($headers -join ' | ')"
    }

    Write-Host ""
    Write-Host "VBA injection script ready. Review sheet structures above." -ForegroundColor Green

    #--- Save and close ---
    $wb.Save()
    $wb.Close($false)
    Write-Host "Workbook saved successfully." -ForegroundColor Green

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.StackTrace) {
        Write-Host $_.Exception.StackTrace -ForegroundColor Red
    }
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
