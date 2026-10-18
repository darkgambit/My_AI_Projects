<#
.SYNOPSIS
    Injects the VBA module directly into FOR DATA ENTRY_FORM.xlsm
#>
param(
    [string]$FilePath = "D:\.claude\FOR DATA ENTRY_FORM.xlsm"
)

$ErrorActionPreference = "Stop"
$excel = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($FilePath)
    $vbProj = $wb.VBProject

    # Try to remove protection
    try {
        $vbProj.Protection = 0
        Write-Host "VBA project unprotected." -ForegroundColor Green
    } catch {
        Write-Host "Cannot unprotect VBA project. Manual steps needed." -ForegroundColor Yellow
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        $wb.Close($false)
        exit
    }

    # Read the module code
    $moduleCode = Get-Content "D:\.claude\VBA_Code\modDataEntry.bas" -Raw

    # Check if module already exists
    $existingComp = $null
    try {
        Set $existingComp = $vbProj.VBComponents("modDataEntry")
    } catch {}

    if ($existingComp -ne $null) {
        Write-Host "Module modDataEntry already exists. Replacing..." -ForegroundColor Yellow
        $vbProj.VBComponents.Remove($vbProj.VBComponents("modDataEntry")) | Out-Null
    }

    # Add new module
    $vbc = $vbProj.VBComponents
    $comp = $vbc.Add(1)  # 1 = vbext_ct_StdModule
    $comp.Name = "modDataEntry"

    # Set the code
    $codeMod = $comp.CodeModule
    $codeMod.AddFromString($moduleCode)

    Write-Host "Module 'modDataEntry' injected successfully!" -ForegroundColor Green

    # Try to set ThisWorkbook code
    try {
        $twCode = Get-Content "D:\.claude\VBA_Code\ThisWorkbook_cls.cls" -Raw
        $twMod = $vbProj.VBComponents("ThisWorkbook").CodeModule
        $twMod.AddFromString($twCode)
        Write-Host "ThisWorkbook code injected." -ForegroundColor Green
    } catch {
        Write-Host "Could not inject ThisWorkbook code: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Save
    $wb.Save()
    Write-Host "Workbook saved successfully!" -ForegroundColor Green

    # Check if module was added
    $count = $vbProj.VBComponents.Count
    Write-Host "Total VBA components: $count" -ForegroundColor Green

    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host ""
    Write-Host "Module code injected. You still need to create the UserForm manually." -ForegroundColor Yellow
    Write-Host "See: D:\.claude\VBA_Code\SETUP_INSTRUCTIONS.txt" -ForegroundColor Gray

} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
