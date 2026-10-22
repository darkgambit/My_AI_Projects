Attribute VB_Name = "SheetScout"
Option Explicit

Private Const MAX_RESULTS As Long = 10000
Private resultRow As Long
Private errorRow As Long
Private searchedCount As Long
Private stopped As Boolean

Public Sub SetupSheetScout()
    Dim ws As Worksheet, btn As Object
    Set ws = GetSheet("Name Search")
    ws.Cells.Clear
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    ws.Range("A1").Value = "SHEETSCOUT - EXCEL NAME SEARCH"
    ws.Range("A3").Value = "Names (separate with commas or new lines)"
    ws.Range("A4").Value = "Folder"
    ws.Range("A5").Value = "Match type (Contains or Exact)"
    ws.Range("B5").Value = "Contains"
    ws.Range("A6").Value = "Header row (0 = no headers)"
    ws.Range("B6").Value = 1
    ws.Range("A8").Value = "1. Enter names in B3. Use Alt+Enter for a new line."
    ws.Range("A9").Value = "2. Choose a folder, then click Search names."
    ws.Range("A10").Value = "Subfolders included. Results replace the previous search."
    ws.Range("A11").Value = "Original files are opened read-only and never saved."
    ws.Range("A12").Value = "Header row applies to all sheets. Rows through that row are skipped."
    ws.Range("A13").Value = "Up to 10,000 results. Name matches ignore capitalization."
    ws.Columns("A").ColumnWidth = 53
    ws.Columns("B").ColumnWidth = 65
    ws.Range("B3:B6").Interior.Color = RGB(226, 243, 235)
    ws.Range("B3:B5").NumberFormat = "@"
    ws.Range("B3").WrapText = True
    ws.Rows(3).RowHeight = 65
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Font.Size = 16
    Set btn = ws.Buttons.Add(ws.Range("D3").Left, ws.Range("D3").Top, 140, 32)
    btn.Caption = "Choose folder"
    btn.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ChooseSearchFolder"
    Set btn = ws.Buttons.Add(ws.Range("D5").Left, ws.Range("D5").Top, 140, 32)
    btn.Caption = "Search names"
    btn.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!SearchNames"
    ws.Activate
End Sub

Public Sub ChooseSearchFolder()
    With Application.FileDialog(4)
        .Title = "Choose folder containing Excel files"
        .AllowMultiSelect = False
        If .Show = -1 Then GetSheet("Name Search").Range("B4").Value = .SelectedItems(1)
    End With
End Sub

Public Sub SearchNames()
    Dim form As Worksheet, output As Worksheet, log As Worksheet
    Dim names As Variant, raw As String, folder As String, mode As String
    Dim headerRow As Long, headerValue As Variant, fso As Object
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldSecurity As Long, oldCalc As Long, failure As String
    Dim oldStatus As Variant
    Set form = GetSheet("Name Search")
    raw = Trim$(CStr(form.Range("B3").Value2))
    folder = Trim$(CStr(form.Range("B4").Value2))
    mode = LCase$(Trim$(CStr(form.Range("B5").Value2)))
    If Len(raw) = 0 Then
        MsgBox "Enter at least one name in B3.", vbExclamation
        Exit Sub
    End If
    If mode <> "contains" And mode <> "exact" Then
        MsgBox "B5 must be Contains or Exact.", vbExclamation
        Exit Sub
    End If
    headerValue = form.Range("B6").Value2
    If Not IsNumeric(headerValue) Then
        MsgBox "B6 must be a whole number from 0 to 1000.", vbExclamation
        Exit Sub
    End If
    If CDbl(headerValue) < 0 Or CDbl(headerValue) > 1000 Or CDbl(headerValue) <> Fix(CDbl(headerValue)) Then
        MsgBox "B6 must be a whole number from 0 to 1000.", vbExclamation
        Exit Sub
    End If
    headerRow = CLng(headerValue)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then
        MsgBox "Choose a valid folder first.", vbExclamation
        Exit Sub
    End If
    raw = Replace(Replace(raw, vbCr, vbLf), ",", vbLf)
    names = Split(raw, vbLf)
    If Len(Trim$(Replace(raw, vbLf, ""))) = 0 Then
        MsgBox "Enter at least one name in B3.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Replace previous Results and Search Log sheets?", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldSecurity = Application.AutomationSecurity
    oldCalc = Application.Calculation
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AutomationSecurity = 3 ' Disable macros in opened workbooks
    Application.Calculation = xlCalculationManual
    Set output = GetSheet("Results")
    Set log = GetSheet("Search Log")
    output.Cells.Clear
    log.Cells.Clear
    output.Range("A1:G1").Value = Array("Workbook", "Sheet", "Excel row", "Names searched", "Matching columns", "Record details", "Full file path")
    log.Range("A1:B1").Value = Array("File / Folder", "Issue")
    output.Columns("A:B").NumberFormat = "@"
    output.Columns("D:G").NumberFormat = "@"
    log.Columns("A:B").NumberFormat = "@"
    resultRow = 2: errorRow = 2: searchedCount = 0: stopped = False
    ScanFolder fso.GetFolder(folder), fso, names, (mode = "exact"), headerRow, output, log
    output.Range("A1:G1").Font.Bold = True
    output.Range("A1:G1").Interior.Color = RGB(205, 235, 222)
    output.Range("A1:G" & CStr(resultRow - 1)).AutoFilter
    output.Columns("A:B").ColumnWidth = 26
    output.Columns("C:E").ColumnWidth = 20
    output.Columns("F").ColumnWidth = 90
    output.Columns("G").ColumnWidth = 45
    log.Columns("A:B").ColumnWidth = 65
    output.Activate
    GoTo Cleanup
Failed:
    failure = "Search stopped: " & Err.Description
Cleanup:
    Application.AutomationSecurity = oldSecurity
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatus
    If Len(failure) > 0 Then
        MsgBox failure, vbExclamation
    Else
        raw = CStr(resultRow - 2) & " matching rows; " & searchedCount & " workbooks searched." & vbCrLf & CStr(errorRow - 2) & " issues (see Search Log)."
        If stopped Then raw = raw & vbCrLf & "Stopped at 10,000 results. Narrow your search for complete results."
        MsgBox raw, vbInformation
    End If
End Sub

Private Sub ScanFolder(ByVal folder As Object, ByVal fso As Object, ByVal names As Variant, ByVal exact As Boolean, ByVal headerRow As Long, ByVal output As Worksheet, ByVal log As Worksheet)
    Dim file As Object, subfolder As Object, ext As String
    On Error GoTo FolderError
    For Each file In folder.Files
        If stopped Then Exit Sub
        ext = LCase$(fso.GetExtensionName(file.Name))
        If (ext = "xlsx" Or ext = "xlsm" Or ext = "xls" Or ext = "xlsb") And Left$(file.Name, 2) <> "~$" Then
            If StrComp(file.Path, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
                ScanWorkbook CStr(file.Path), names, exact, headerRow, output, log
            End If
        End If
    Next file
    For Each subfolder In folder.SubFolders
        If stopped Then Exit Sub
        ' Skip junctions / symbolic links to avoid recursion loops.
        If (subfolder.Attributes And 1024) = 0 Then ScanFolder subfolder, fso, names, exact, headerRow, output, log
    Next subfolder
    Exit Sub
FolderError:
    LogIssue log, CStr(folder.Path), "Folder could not be fully scanned: " & Err.Description
End Sub

Private Sub ScanWorkbook(ByVal path As String, ByVal names As Variant, ByVal exact As Boolean, ByVal headerRow As Long, ByVal output As Worksheet, ByVal log As Worksheet)
    Dim book As Workbook, existing As Workbook, sheet As Worksheet, area As Range
    Dim data As Variant, headers As Variant, r As Long, c As Long, n As Long
    Dim value As String, needle As String, hits As String, cols As String, details As String
    Dim rowHits As Object, columnHit As Boolean, key As Variant, label As String
    On Error GoTo FileError
    For Each existing In Application.Workbooks
        If StrComp(existing.FullName, path, vbTextCompare) = 0 Then
            LogIssue log, path, "Skipped: workbook already open. Close it and search again."
            Exit Sub
        End If
    Next existing
    Application.StatusBar = "Searching " & path
    Set book = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, Password:="", WriteResPassword:="", IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
    For Each sheet In book.Worksheets
        Set area = sheet.UsedRange
        ' Force a 2D array even for a sheet containing a single used cell.
        If area.Cells.CountLarge = 1 Then
            ReDim data(1 To 1, 1 To 1)
            data(1, 1) = area.Value2
        Else
            data = area.Value2
        End If
        If headerRow > 0 Then headers = sheet.Range(sheet.Cells(headerRow, area.Column), sheet.Cells(headerRow, area.Column + area.Columns.Count - 1)).Value2
        For r = 1 To UBound(data, 1)
            If area.Row + r - 1 > headerRow Then
                Set rowHits = CreateObject("Scripting.Dictionary")
                rowHits.CompareMode = vbTextCompare
                cols = "": details = ""
                For c = 1 To UBound(data, 2)
                    value = CellText(data(r, c))
                    columnHit = False
                    For n = LBound(names) To UBound(names)
                        needle = CleanName(CStr(names(n)))
                        If Len(needle) > 0 Then
                            If (exact And StrComp(CleanName(value), needle, vbTextCompare) = 0) Or (Not exact And InStr(1, CleanName(value), needle, vbTextCompare) > 0) Then
                                rowHits(Trim$(CStr(names(n)))) = True
                                columnHit = True
                            End If
                        End If
                    Next n
                    If columnHit Then cols = JoinPart(cols, ColumnLetter(area.Column + c - 1), ", ")
                    If Len(value) > 0 Then
                        label = ""
                        If headerRow > 0 Then
                            If IsArray(headers) Then
                                label = CellText(headers(1, c))
                            Else
                                label = CellText(headers)
                            End If
                        End If
                        If Len(label) = 0 Then label = "Column " & ColumnLetter(area.Column + c - 1)
                        details = Left$(JoinPart(details, label & ": " & value, " | "), 32000)
                    End If
                Next c
                If rowHits.Count > 0 Then
                    If resultRow > MAX_RESULTS + 1 Then
                        stopped = True
                        GoTo CloseBook
                    End If
                    hits = ""
                    For Each key In rowHits.Keys
                        hits = JoinPart(hits, CStr(key), "; ")
                    Next key
                    output.Cells(resultRow, 1).Value2 = book.Name
                    output.Cells(resultRow, 2).Value2 = sheet.Name
                    output.Cells(resultRow, 3).Value2 = area.Row + r - 1
                    output.Cells(resultRow, 4).Value2 = Left$(hits, 32000)
                    output.Cells(resultRow, 5).Value2 = cols
                    output.Cells(resultRow, 6).Value2 = details
                    output.Cells(resultRow, 7).Value2 = path
                    resultRow = resultRow + 1
                End If
            End If
        Next r
    Next sheet
    searchedCount = searchedCount + 1
CloseBook:
    book.Close SaveChanges:=False
    Exit Sub
FileError:
    LogIssue log, path, "Workbook could not be fully scanned: " & Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function GetSheet(ByVal title As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(title)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Set GetSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSheet.Name = title
    End If
End Function

Private Function CellText(ByVal value As Variant) As String
    If IsError(value) Then
        CellText = "[Excel error]"
    ElseIf IsEmpty(value) Or IsNull(value) Then
        CellText = ""
    Else
        CellText = CStr(value)
    End If
End Function

Private Function CleanName(ByVal value As String) As String
    value = Replace(Replace(Replace(value, vbTab, " "), vbCr, " "), vbLf, " ")
    value = Replace(value, ChrW(160), " ")
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    CleanName = Trim$(value)
End Function

Private Function JoinPart(ByVal before As String, ByVal item As String, ByVal separator As String) As String
    If Len(before) = 0 Then JoinPart = item Else JoinPart = before & separator & item
End Function

Private Function ColumnLetter(ByVal number As Long) As String
    Dim result As String
    Do While number > 0
        number = number - 1
        result = Chr$(65 + (number Mod 26)) & result
        number = number \ 26
    Loop
    ColumnLetter = result
End Function

Private Sub LogIssue(ByVal log As Worksheet, ByVal path As String, ByVal issue As String)
    log.Cells(errorRow, 1).Value2 = path
    log.Cells(errorRow, 2).Value2 = issue
    errorRow = errorRow + 1
End Sub
