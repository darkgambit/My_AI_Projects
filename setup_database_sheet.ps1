<#
.SYNOPSIS
    Sets up the DATABASE sheet in FOR DATA ENTRY_FORM.xlsm
#>
param(
    [string]$FilePath = "D:\.claude\FOR DATA ENTRY_FORM.xlsm"
)

$ErrorActionPreference = "Stop"
$excel = $null

try {
    Write-Host "Setting up DATABASE sheet in: $FilePath" -ForegroundColor Cyan

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($FilePath)

    # Check if DATABASE sheet already exists
    $dbExists = $false
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq "DATABASE") {
            $dbExists = $true
            break
        }
    }

    if (!$dbExists) {
        Write-Host "Creating DATABASE sheet..." -ForegroundColor Yellow
        $ws = $wb.Worksheets.Add()
        $ws.Name = "DATABASE"

        # Add headers
        $ws.Cells(1, 1).Value2 = "Sheet"
        $ws.Cells(1, 2).Value2 = "Record ID"
        $ws.Cells(1, 3).Value2 = "Timestamp"
        $ws.Cells(1, 4).Value2 = "Data (JSON)"
        $ws.Cells(2, 1).Value2 = "Instructions:"
        $ws.Cells(3, 1).Value2 = "1. Click 'Data Entry' menu -> 'Open Data Entry Form'"
        $ws.Cells(4, 1).Value2 = "2. Enter data in the HTML form"
        $ws.Cells(5, 1).Value2 = "3. Click 'Copy to Clipboard' in the form"
        $ws.Cells(6, 1).Value2 = "4. Return here and click 'Data Entry' -> 'Import from Clipboard'"

        # Format headers
        $headerRange = $ws.Range("A1:D1")
        $headerRange.Font.Bold = $true
        $headerRange.Interior.Color = 5987160  # Blue
        $headerRange.Font.Color = 16777215     # White

        $ws.Columns("A:D").AutoFit
        $ws.Range("A2:D2").EntireRow.Hidden = $true

        Write-Host "DATABASE sheet created!" -ForegroundColor Green
    } else {
        Write-Host "DATABASE sheet already exists - keeping existing data." -ForegroundColor Green
    }

    # Save the workbook
    $wb.Save()
    $wb.Close($false)

    Write-Host "Workbook saved successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "DATABASE sheet ready. The form now saves data directly to this sheet." -ForegroundColor Cyan

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
