<#
.SYNOPSIS
    Injects complete VBA code into FOR DATA ENTRY_FORM.xlsm to create a modern data entry form
    with Add, Delete, Update, and Clear buttons.

.NOTES
    This script uses Excel COM automation to add a UserForm and VBA code.
    The form supports 4 tabs matching your 4 sheets.
#>
param(
    [string]$FilePath = "D:\.claude\FOR DATA ENTRY_FORM.xlsm"
)

$ErrorActionPreference = "Stop"
$excel = $null

Write-Host "Starting VBA injection..." -ForegroundColor Cyan
Write-Host "Target file: $FilePath" -ForegroundColor Gray

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($FilePath)

    # Check if VBA project can be modified
    $vbProj = $wb.VBProject
    $canModify = $false
    try {
        $vbProj.Protection = 0  # Remove protection
        $canModify = $true
    } catch {
        Write-Host "Cannot modify VBA project directly. Will export code for manual import." -ForegroundColor Yellow
    }

    # ============================================================
    # Export the VBA code to files so user can import manually
    # ============================================================
    $exportPath = "D:\.claude\VBA_Code"
    if (!(Test-Path $exportPath)) {
        New-Item -ItemType Directory -Path $exportPath | Out-Null
    }

    # --- Module: modDataEntry ---
    $moduleCode = @'
Option Explicit

Public CurrentSheet As String
Public CurrentRow As Long
Public FormMode As String
Public LastDataRow As Long

' Main entry point - show the data entry form
Public Sub ShowDataEntryForm()
    FormMode = "Add"
    frmDataEntry.lblStatus.Caption = "Ready - Select a sheet and start entering data."
    frmDataEntry.Show vbModeless
End Sub

' Get the last data row for a given sheet
Public Function GetLastDataRow(sheetName As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        GetLastDataRow = 2
    Else
        GetLastDataRow = lastRow
    End If
End Function

' Populate a listbox with all data for selection
Public Sub PopulateListBox(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    frmDataEntry.lstRecords.Clear

    If lastRow < 2 Then
        frmDataEntry.lstRecords.AddItem "No data yet"
        Exit Sub
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim displayText As String
        displayText = ""
        Dim c As Long
        For c = 1 To ws.UsedRange.Columns.Count
            If ws.Cells(r, c).Value <> "" Then
                displayText = displayText & ws.Cells(r, c).Value & " | "
            End If
        Next c
        If Len(displayText) > 0 Then
            displayText = Left(displayText, Len(displayText) - 2)
            If Len(displayText) > 100 Then displayText = Left(displayText, 97) & "..."
        End If
        If displayText = "" Then displayText = "(Row " & r & " - empty)"
        frmDataEntry.lstRecords.AddItem displayText
    Next r
    frmDataEntry.lstRecords.ListIndex = 0
End Sub

' Load form fields from a selected row
Public Sub LoadRowToForm(sheetName As String, rowIndex As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Select Case sheetName
        Case "01_Farmer_Profile"
            frmDataEntry.txtRecordID.Text = Nz(ws.Cells(rowIndex, 1).Value)
            frmDataEntry.txtCropYear.Text = Nz(ws.Cells(rowIndex, 2).Value)
            frmDataEntry.txtFCVVariety.Text = Nz(ws.Cells(rowIndex, 3).Value)
            frmDataEntry.txtBurleyVariety.Text = Nz(ws.Cells(rowIndex, 4).Value)
            frmDataEntry.txtRegisteredArea.Text = Nz(ws.Cells(rowIndex, 5).Value)
            frmDataEntry.txtHomeAddress.Text = Nz(ws.Cells(rowIndex, 6).Value)
            frmDataEntry.txtContractNo.Text = Nz(ws.Cells(rowIndex, 7).Value)
            frmDataEntry.txtDateSigned.Text = Nz(ws.Cells(rowIndex, 8).Value)
        Case "02_Seedbed-SEEDLING CARE "
            frmDataEntry.txtSeedSource.Text = Nz(ws.Cells(rowIndex, 1).Value)
            frmDataEntry.txtSeedbedType.Text = Nz(ws.Cells(rowIndex, 2).Value)
            frmDataEntry.txtNumPacks.Text = Nz(ws.Cells(rowIndex, 3).Value)
            frmDataEntry.txtGramPerPack.Text = Nz(ws.Cells(rowIndex, 4).Value)
            frmDataEntry.txtPackNo.Text = Nz(ws.Cells(rowIndex, 5).Value)
            frmDataEntry.txtLotNo.Text = Nz(ws.Cells(rowIndex, 6).Value)
            frmDataEntry.txtNumBeds.Text = Nz(ws.Cells(rowIndex, 7).Value)
            frmDataEntry.txtSeedbedSize.Text = Nz(ws.Cells(rowIndex, 8).Value)
            frmDataEntry.txtNumTrays.Text = Nz(ws.Cells(rowIndex, 9).Value)
            frmDataEntry.txtEstSeedlings.Text = Nz(ws.Cells(rowIndex, 10).Value)
        Case "03_Transplant_Volume"
            frmDataEntry.txtActivity.Text = Nz(ws.Cells(rowIndex, 1).Value)
            frmDataEntry.txtRound.Text = Nz(ws.Cells(rowIndex, 2).Value)
            frmDataEntry.txtDate.Text = Nz(ws.Cells(rowIndex, 3).Value)
            frmDataEntry.txtMethod.Text = Nz(ws.Cells(rowIndex, 4).Value)
            frmDataEntry.txtModel.Text = Nz(ws.Cells(rowIndex, 5).Value)
            frmDataEntry.txtHP.Text = Nz(ws.Cells(rowIndex, 6).Value)
            frmDataEntry.txtFuelL.Text = Nz(ws.Cells(rowIndex, 7).Value)
            frmDataEntry.txtPPE.Text = Nz(ws.Cells(rowIndex, 8).Value)
            frmDataEntry.txtTotalLabor.Text = Nz(ws.Cells(rowIndex, 9).Value)
            frmDataEntry.txtRemarks.Text = Nz(ws.Cells(rowIndex, 10).Value)
        Case "04_IPM_Topping"
            frmDataEntry.txtWeek.Text = Nz(ws.Cells(rowIndex, 1).Value)
            frmDataEntry.txtIPMDate.Text = Nz(ws.Cells(rowIndex, 2).Value)
            frmDataEntry.txtETL.Text = Nz(ws.Cells(rowIndex, 3).Value)
            frmDataEntry.txtInsectFound.Text = Nz(ws.Cells(rowIndex, 4).Value)
            frmDataEntry.txtPlantStage.Text = Nz(ws.Cells(rowIndex, 5).Value)
            frmDataEntry.txtManualPicking.Text = Nz(ws.Cells(rowIndex, 6).Value)
            frmDataEntry.txtCPAUsed.Text = Nz(ws.Cells(rowIndex, 7).Value)
            frmDataEntry.txtDosage.Text = Nz(ws.Cells(rowIndex, 8).Value)
            frmDataEntry.txtIPMPPE.Text = Nz(ws.Cells(rowIndex, 9).Value)
            frmDataEntry.txtWarningSign.Text = Nz(ws.Cells(rowIndex, 10).Value)
            frmDataEntry.txtFamilyM.Text = Nz(ws.Cells(rowIndex, 11).Value)
            frmDataEntry.txtFamilyF.Text = Nz(ws.Cells(rowIndex, 12).Value)
            frmDataEntry.txtFulltimeM.Text = Nz(ws.Cells(rowIndex, 13).Value)
            frmDataEntry.txtFulltimeF.Text = Nz(ws.Cells(rowIndex, 14).Value)
            frmDataEntry.txtTempM.Text = Nz(ws.Cells(rowIndex, 15).Value)
            frmDataEntry.txtTempF.Text = Nz(ws.Cells(rowIndex, 16).Value)
            frmDataEntry.txtTotalLabor2.Text = Nz(ws.Cells(rowIndex, 17).Value)
    End Select
End Sub

' Save form data to the selected sheet
Public Sub SaveFormData(sheetName As String, targetRow As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Application.ScreenUpdating = False

    Select Case sheetName
        Case "01_Farmer_Profile"
            ws.Cells(targetRow, 1).Value = Nz(frmDataEntry.txtRecordID.Text)
            ws.Cells(targetRow, 2).Value = Nz(frmDataEntry.txtCropYear.Text)
            ws.Cells(targetRow, 3).Value = Nz(frmDataEntry.txtFCVVariety.Text)
            ws.Cells(targetRow, 4).Value = Nz(frmDataEntry.txtBurleyVariety.Text)
            ws.Cells(targetRow, 5).Value = Nz(frmDataEntry.txtRegisteredArea.Text)
            ws.Cells(targetRow, 6).Value = Nz(frmDataEntry.txtHomeAddress.Text)
            ws.Cells(targetRow, 7).Value = Nz(frmDataEntry.txtContractNo.Text)
            ws.Cells(targetRow, 8).Value = Nz(frmDataEntry.txtDateSigned.Text)
        Case "02_Seedbed-SEEDLING CARE "
            ws.Cells(targetRow, 1).Value = Nz(frmDataEntry.txtSeedSource.Text)
            ws.Cells(targetRow, 2).Value = Nz(frmDataEntry.txtSeedbedType.Text)
            ws.Cells(targetRow, 3).Value = Nz(frmDataEntry.txtNumPacks.Text)
            ws.Cells(targetRow, 4).Value = Nz(frmDataEntry.txtGramPerPack.Text)
            ws.Cells(targetRow, 5).Value = Nz(frmDataEntry.txtPackNo.Text)
            ws.Cells(targetRow, 6).Value = Nz(frmDataEntry.txtLotNo.Text)
            ws.Cells(targetRow, 7).Value = Nz(frmDataEntry.txtNumBeds.Text)
            ws.Cells(targetRow, 8).Value = Nz(frmDataEntry.txtSeedbedSize.Text)
            ws.Cells(targetRow, 9).Value = Nz(frmDataEntry.txtNumTrays.Text)
            ws.Cells(targetRow, 10).Value = Nz(frmDataEntry.txtEstSeedlings.Text)
        Case "03_Transplant_Volume"
            ws.Cells(targetRow, 1).Value = Nz(frmDataEntry.txtActivity.Text)
            ws.Cells(targetRow, 2).Value = Nz(frmDataEntry.txtRound.Text)
            ws.Cells(targetRow, 3).Value = Nz(frmDataEntry.txtDate.Text)
            ws.Cells(targetRow, 4).Value = Nz(frmDataEntry.txtMethod.Text)
            ws.Cells(targetRow, 5).Value = Nz(frmDataEntry.txtModel.Text)
            ws.Cells(targetRow, 6).Value = Nz(frmDataEntry.txtHP.Text)
            ws.Cells(targetRow, 7).Value = Nz(frmDataEntry.txtFuelL.Text)
            ws.Cells(targetRow, 8).Value = Nz(frmDataEntry.txtPPE.Text)
            ws.Cells(targetRow, 9).Value = Nz(frmDataEntry.txtTotalLabor.Text)
            ws.Cells(targetRow, 10).Value = Nz(frmDataEntry.txtRemarks.Text)
        Case "04_IPM_Topping"
            ws.Cells(targetRow, 1).Value = Nz(frmDataEntry.txtWeek.Text)
            ws.Cells(targetRow, 2).Value = Nz(frmDataEntry.txtIPMDate.Text)
            ws.Cells(targetRow, 3).Value = Nz(frmDataEntry.txtETL.Text)
            ws.Cells(targetRow, 4).Value = Nz(frmDataEntry.txtInsectFound.Text)
            ws.Cells(targetRow, 5).Value = Nz(frmDataEntry.txtPlantStage.Text)
            ws.Cells(targetRow, 6).Value = Nz(frmDataEntry.txtManualPicking.Text)
            ws.Cells(targetRow, 7).Value = Nz(frmDataEntry.txtCPAUsed.Text)
            ws.Cells(targetRow, 8).Value = Nz(frmDataEntry.txtDosage.Text)
            ws.Cells(targetRow, 9).Value = Nz(frmDataEntry.txtIPMPPE.Text)
            ws.Cells(targetRow, 10).Value = Nz(frmDataEntry.txtWarningSign.Text)
            ws.Cells(targetRow, 11).Value = Nz(frmDataEntry.txtFamilyM.Text)
            ws.Cells(targetRow, 12).Value = Nz(frmDataEntry.txtFamilyF.Text)
            ws.Cells(targetRow, 13).Value = Nz(frmDataEntry.txtFulltimeM.Text)
            ws.Cells(targetRow, 14).Value = Nz(frmDataEntry.txtFulltimeF.Text)
            ws.Cells(targetRow, 15).Value = Nz(frmDataEntry.txtTempM.Text)
            ws.Cells(targetRow, 16).Value = Nz(frmDataEntry.txtTempF.Text)
            ws.Cells(targetRow, 17).Value = Nz(frmDataEntry.txtTotalLabor2.Text)
    End Select
    Application.ScreenUpdating = True
End Sub

' Clear all form fields
Public Sub ClearFormFields()
    Dim ctrl As Object
    For Each ctrl In frmDataEntry.Controls
        If TypeName(ctrl) = "TextBox" Then ctrl.Text = ""
        If TypeName(ctrl) = "ComboBox" Then ctrl.Text = ""
    Next ctrl
End Sub

' Helper - null safe
Public Function Nz(v As Variant) As Variant
    If IsEmpty(v) Or IsNull(v) Then Nz = "" Else Nz = v
End Function

' Show status message
Public Sub ShowStatus(msg As String)
    frmDataEntry.lblStatus.Caption = msg
End Sub

' ADD button handler
Public Sub DoAdd()
    Dim sheetName As String
    sheetName = GetCurrentSheetName()
    If sheetName = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    SaveFormData sheetName, nextRow
    PopulateListBox sheetName
    ShowStatus "Record added to row " & nextRow
    Application.SwitchWidth = 0
End Sub

' DELETE button handler
Public Sub DoDelete()
    Dim sheetName As String
    sheetName = GetCurrentSheetName()
    If sheetName = "" Then Exit Sub

    Dim selIndex As Long
    selIndex = frmDataEntry.lstRecords.ListIndex
    If selIndex < 0 Then
        MsgBox "Please select a record to delete.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Map listbox index to actual sheet row
    Dim sheetRow As Long
    If lastRow >= 2 Then sheetRow = selIndex + 2 Else Exit Sub

    If MsgBox("Delete this record?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ws.Rows(sheetRow).Delete Shift:=xlUp
    PopulateListBox sheetName
    ShowStatus "Record deleted."
    ClearFormFields
End Sub

' UPDATE button handler
Public Sub DoUpdate()
    Dim sheetName As String
    sheetName = GetCurrentSheetName()
    If sheetName = "" Then Exit Sub

    Dim selIndex As Long
    selIndex = frmDataEntry.lstRecords.ListIndex
    If selIndex < 0 Then
        MsgBox "Please select a record to update.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim sheetRow As Long
    If lastRow >= 2 Then sheetRow = selIndex + 2 Else Exit Sub

    SaveFormData sheetName, sheetRow
    PopulateListBox sheetName
    ShowStatus "Record updated at row " & sheetRow
End Sub

' CLEAR button handler
Public Sub DoClear()
    ClearFormFields
    frmDataEntry.lstRecords.Clear
    ShowStatus "Form cleared."
End Sub

' Get current active sheet name from tab
Public Function GetCurrentSheetName() As String
    Dim idx As Integer
    idx = frmDataEntry.TabCtl1.Value
    Select Case idx
        Case 0: GetCurrentSheetName = "01_Farmer_Profile"
        Case 1: GetCurrentSheetName = "02_Seedbed-SEEDLING CARE "
        Case 2: GetCurrentSheetName = "03_Transplant_Volume"
        Case 3: GetCurrentSheetName = "04_IPM_Topping"
        Case Else: GetCurrentSheetName = ""
    End Select
End Function

' Tab change - load data for new sheet
Public Sub OnTabChange()
    Dim sheetName As String
    sheetName = GetCurrentSheetName()
    If sheetName = "" Then Exit Sub
    PopulateListBox sheetName
    ClearFormFields
    ShowStatus "Active sheet: " & sheetName
End Sub

' Listbox selection change - load selected row
Public Sub OnListChange()
    Dim sheetName As String
    sheetName = GetCurrentSheetName()
    If sheetName = "" Then Exit Sub

    Dim selIndex As Long
    selIndex = frmDataEntry.lstRecords.ListIndex
    If selIndex < 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim sheetRow As Long
    sheetRow = selIndex + 2

    LoadRowToForm sheetName, sheetRow
    FormMode = "View"
    ShowStatus "Selected row " & sheetRow
End Sub
'@End Module Code
'@ModuleName=modDataEntry

    $moduleCode | Out-File "$exportPath\modDataEntry.bas" -Encoding ascii
    Write-Host "Module file exported: $exportPath\modDataEntry.bas" -ForegroundColor Green

    # --- ThisWorkbook Code ---
    $thisWorkbookCode = @'
Option Explicit

Private Sub Workbook_Open()
    Call CreateDataEntryMenu
End Sub

Public Sub CreateDataEntryMenu()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' Remove existing menu if present
    On Error Resume Next
    wb.Styles("DataEntryMenu").Delete
    On Error GoTo 0

    ' Add ribbon tab via XML (simplified: use menu bar approach)
    ' For modern Excel with ribbon, we add a custom menu
    Dim menuHandler As String
    menuHandler = "ShowDataEntryForm"

    ' Create menu bar
    Dim cb As CommandBar
    Set cb = CommandBars.Add(Name:="DataEntryForm", Position:=msoBarTop)

    Dim ctrl As CommandBarControl
    Set ctrl = cb.Controls.Add(Type:=msoControlButton)
    With ctrl
        .Caption = "Open Data Entry Form"
        .FaceId = 59
        .OnAction = "ShowDataEntryForm"
    End With

    cb.Visible = True
End Sub
'@End ThisWorkbook Code

    $thisWorkbookCode | Out-File "$exportPath\ThisWorkbook_cls.cls" -Encoding ascii
    Write-Host "ThisWorkbook file exported: $exportPath\ThisWorkbook_cls.cls" -ForegroundColor Green

    # --- UserForm code ---
    # Since we can't create a .frm file programmatically easily, create instructions
    $instructions = @"
============================================================
DATA ENTRY FORM - SETUP INSTRUCTIONS
============================================================

This Excel file has been prepared with data sheets from your original
"FOR DATA ENTRY.xlsx" file.

To add the VBA Data Entry Form:

1. Open "FOR DATA ENTRY_FORM.xlsm" in Excel
2. Press Alt + F11 to open the VBA Editor
3. In the VBA Editor:
   a. Go to File -> Import File
   b. Import "modDataEntry.bas" from: $exportPath\modDataEntry.bas
   c. In the Project Explorer, double-click "ThisWorkbook"
   d. Replace its contents with code from "ThisWorkbook_cls.cls"

4. Create a new UserForm:
   a. Insert -> UserForm
   b. Name: frmDataEntry
   c. Caption: "Data Entry Form"
   d. Set properties:
      - Width: 800
      - Height: 550

5. Add these controls to the form:

   TAB CONTROL (Name: TabCtl1, 4 tabs):
   - Tab 1: Caption = "Farmer Profile"
   - Tab 2: Caption = "Seedling Care"
   - Tab 3: Caption = "Transplant Volume"
   - Tab 4: Caption = "IPM & Topping"

   For each tab, add a frame named "Frame[Name]" with appropriate controls:

   TAB 1 - Farmer Profile:
   - txtRecordID, txtCropYear, txtFCVVariety, txtBurleyVariety, txtRegisteredArea
   - txtHomeAddress, txtContractNo, txtDateSigned
   (Labels and TextBoxes arranged in 2-column layout)

   TAB 2 - Seedling Care:
   - txtSeedSource, txtSeedbedType, txtNumPacks, txtGramPerPack
   - txtPackNo, txtLotNo, txtNumBeds, txtSeedbedSize, txtNumTrays, txtEstSeedlings

   TAB 3 - Transplant Volume:
   - txtActivity, txtRound, txtDate, txtMethod, txtModel
   - txtHP, txtFuelL, txtPPE, txtTotalLabor, txtRemarks

   TAB 4 - IPM & Topping:
   - txtWeek, txtIPMDate, txtETL, txtInsectFound, txtPlantStage
   - txtManualPicking, txtCPAUsed, txtDosage, txtIPMPPE
   - txtWarningSign, txtFamilyM, txtFamilyF, txtFulltimeM
   - txtFulltimeF, txtTempM, txtTempF, txtTotalLabor2

   BOTTOM CONTROLS:
   - lstRecords (ListBox) - Width: 760, Height: 120
   - 4 Buttons: cmdAdd, cmdDelete, cmdUpdate, cmdClear
   - lblStatus (Label) - Width: 760, Height: 25

6. Add these event handlers to the UserForm code:
   (See frmDataEntry_Code.txt for the complete code)

7. Save the workbook (Ctrl + S)
8. Close and reopen the file
9. Enable macros when prompted
10. Click the "Data Entry Form" menu at the top

============================================================
FOR IMMEDIATE USE: Run the macro "ShowDataEntryForm"
============================================================
"@

    $instructions | Out-File "$exportPath\SETUP_INSTRUCTIONS.txt" -Encoding utf8
    Write-Host "Instructions exported: $exportPath\SETUP_INSTRUCTIONS.txt" -ForegroundColor Green

    # --- UserForm code behind ---
    $frmCode = @'
Option Explicit

Private Sub UserForm_Initialize()
    ' Set initial sheet selection
    TabCtl1.Value = 0
    Call modDataEntry.OnTabChange
End Sub

Private Sub TabCtl1_Change()
    Call modDataEntry.OnTabChange
End Sub

Private Sub lstRecords_Click()
    Call modDataEntry.OnListChange
End Sub

Private Sub lstRecords_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call modDataEntry.OnListChange
End Sub

Private Sub cmdAdd_Click()
    Call modDataEntry.DoAdd
End Sub

Private Sub cmdDelete_Click()
    Call modDataEntry.DoDelete
End Sub

Private Sub cmdUpdate_Click()
    Call modDataEntry.DoUpdate
End Sub

Private Sub cmdClear_Click()
    Call modDataEntry.DoClear
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Clean up
End Sub
'@End UserForm Code

    $frmCode | Out-File "$exportPath\frmDataEntry_Code.txt" -Encoding ascii
    Write-Host "UserForm code exported: $exportPath\frmDataEntry_Code.txt" -ForegroundColor Green

    # --- Try to inject VBA if possible ---
    if ($canModify) {
        Write-Host ""
        Write-Host "Attempting to inject VBA code directly..." -ForegroundColor Cyan

        try {
            # Add module
            $vbc = $vbProj.VBComponents
            $comp = $vbc.Add(1)  # 1 = vbext_ct_StdModule
            $comp.Name = "modDataEntry"

            # Set the code
            $codeMod = $comp.CodeModule
            $codeMod.AddFromString($moduleCode -replace "'@End Module Code`r`n'", "")
            Write-Host "Module injected successfully!" -ForegroundColor Green

            # Try to set ThisWorkbook code
            try {
                $twMod = $vbProj.VBComponents("ThisWorkbook").CodeModule
                $twCode = $thisWorkbookCode -replace "'@End ThisWorkbook Code`r`n'", ""
                $twMod.AddFromString($twCode)
                Write-Host "ThisWorkbook code injected." -ForegroundColor Green
            } catch {
                Write-Host "Could not inject ThisWorkbook code. Will need manual setup." -ForegroundColor Yellow
            }

            # Note: UserForm creation programmatically is complex - user should create manually
            Write-Host "Module code injected. UserForm needs manual creation." -ForegroundColor Yellow
            Write-Host "See setup instructions for UserForm creation steps." -ForegroundColor Yellow

            $wb.Save()
            $wb.Close($false)
            Write-Host ""
            Write-Host "Workbook saved with VBA code injected!" -ForegroundColor Green
        } catch {
            Write-Host "VBA injection failed: $($_.Exception.Message)" -ForegroundColor Red
            $wb.Close($false)
        }
    } else {
        $wb.Close($false)
        Write-Host ""
        Write-Host "VBA project is protected. Manual import required." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "VBA CODE EXPORTED TO: $exportPath" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "Files created:" -ForegroundColor Gray
    Write-Host "  - modDataEntry.bas (VBA module)" -ForegroundColor Gray
    Write-Host "  - ThisWorkbook_cls.cls (ThisWorkbook code)" -ForegroundColor Gray
    Write-Host "  - frmDataEntry_Code.txt (UserForm code)" -ForegroundColor Gray
    Write-Host "  - SETUP_INSTRUCTIONS.txt (Step-by-step guide)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Follow the setup instructions to complete the form." -ForegroundColor Yellow

} catch {
    Write-Host "FATAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.StackTrace) {
        Write-Host $_.Exception.StackTrace -ForegroundColor Red
    }
} finally {
    if ($excel) {
        if (!$excel.HasWorkbook("FOR DATA ENTRY_FORM.xlsm")) {
            # Already closed
        }
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}