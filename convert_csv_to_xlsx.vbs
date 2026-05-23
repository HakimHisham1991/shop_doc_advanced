' =============================================
' convert_csv_to_xlsx.vbs - With Data Validation
' Emits SHOPDOC_LOG|Level|Category|Message for TCL error log parsing
' =============================================

Sub ShopDocLog(level, category, msg)
    WScript.Echo "SHOPDOC_LOG|" & level & "|" & category & "|" & msg
End Sub

Function NormalizePath(p)
    NormalizePath = p
End Function

If WScript.Arguments.Count < 1 Then
    ShopDocLog "ERROR", "Script", "Usage: convert_csv_to_xlsx.vbs input.csv [output.xlsx]"
    WScript.Quit 1
End If

csvFile = WScript.Arguments(0)

If WScript.Arguments.Count > 1 Then
    xlsxFile = WScript.Arguments(1)
Else
    Set fso = CreateObject("Scripting.FileSystemObject")
    xlsxFile = fso.GetParentFolderName(csvFile) & "\" & fso.GetBaseName(csvFile) & ".xlsx"
End If

ShopDocLog "INFO", "Script", "convert_csv_to_xlsx.vbs started"
ShopDocLog "INFO", "Paths", "CSV input: " & csvFile
ShopDocLog "INFO", "Paths", "XLSX output: " & xlsxFile

Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(csvFile) Then
    ShopDocLog "ERROR", "CSV", "Input CSV file not found: " & csvFile
    WScript.Quit 1
End If

On Error Resume Next
Set probeFile = fso.CreateTextFile(fso.GetParentFolderName(xlsxFile) & "\_shopdoc_vbs_probe.tmp", True)
If Err.Number <> 0 Then
    ShopDocLog "ERROR", "Permission", "Cannot write to output folder (permission denied, read-only path, or policy block): " & fso.GetParentFolderName(xlsxFile) & " - " & Err.Description
    WScript.Quit 1
End If
probeFile.Close
fso.DeleteFile fso.GetParentFolderName(xlsxFile) & "\_shopdoc_vbs_probe.tmp", True
On Error GoTo 0

ShopDocLog "INFO", "Permission", "Output folder is writable: " & fso.GetParentFolderName(xlsxFile)

If fso.FileExists(xlsxFile) Then
    ShopDocLog "WARN", "XLSX", "Target XLSX already exists and will be overwritten: " & xlsxFile
End If

On Error Resume Next

Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    ShopDocLog "ERROR", "Excel", "Cannot create Excel.Application COM object - " & Err.Number & " " & Err.Description
    ShopDocLog "ERROR", "Excel", "Desktop Microsoft Excel may be missing, not licensed, or COM automation is blocked by IT policy"
    ShopDocLog "ERROR", "Excel", "Office Online / Excel Viewer cannot be used for automated conversion"
    ShopDocLog "ERROR", "Policy", "Ask IT to allow Excel COM automation and cscript.exe for this user"
    WScript.Quit 1
End If

ShopDocLog "INFO", "Excel", "Excel.Application created - version " & objExcel.Version & " build " & objExcel.Build

objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(csvFile)
If Err.Number <> 0 Then
    ShopDocLog "ERROR", "Excel", "Cannot open CSV in Excel: " & csvFile & " - " & Err.Number & " " & Err.Description
    ShopDocLog "ERROR", "CSV", "File may be locked by another program, corrupted, or blocked by antivirus"
    objExcel.Quit
    WScript.Quit 1
End If

ShopDocLog "INFO", "Excel", "Opened CSV workbook successfully"

Set objWorksheet = objWorkbook.Worksheets(1)

' Get the last row
lastRow = objWorksheet.UsedRange.Rows.Count
ShopDocLog "INFO", "CSV", "CSV has " & lastRow & " rows"

' Find column function (row 1 exact match after trim)
Function FindColumn(ws, colName)
    FindColumn = 0
    For col = 1 To ws.UsedRange.Columns.Count
        If Trim(ws.Cells(1, col).Value) = colName Then
            FindColumn = col
            Exit Function
        End If
    Next
End Function

' Normalize newlines in header cell (CSV/LF vs Excel CRLF)
Function NormalizeHeaderText(v)
    Dim s
    s = CStr(v)
    s = Replace(s, vbCrLf, Chr(10))
    s = Replace(s, vbCr, Chr(10))
    NormalizeHeaderText = s
End Function

' Multiline Finish Type header: match after normalizing line breaks, with prefix fallback
Function FindFinishTypeColumn(ws)
    Dim col, v, vNorm, expected
    FindFinishTypeColumn = 0
    expected = NormalizeHeaderText("Finish Type " & Chr(10) & "(Finish / Controlled Roughing / Free Roughing)")
    For col = 1 To ws.UsedRange.Columns.Count
        vNorm = NormalizeHeaderText(ws.Cells(1, col).Value)
        If vNorm = expected Then
            FindFinishTypeColumn = col
            Exit Function
        End If
    Next
    For col = 1 To ws.UsedRange.Columns.Count
        v = NormalizeHeaderText(ws.Cells(1, col).Value)
        If InStr(1, v, "Finish Type", 1) = 1 And InStr(1, v, "(Finish / Controlled Roughing / Free Roughing)", 1) > 0 Then
            FindFinishTypeColumn = col
            Exit Function
        End If
    Next
End Function

' Add validation to a range
Sub AddValidation(ws, colNum, lastRowNum, validationString)
    If colNum > 0 And lastRowNum > 1 Then
        Set rng = ws.Range(ws.Cells(2, colNum), ws.Cells(lastRowNum, colNum))
        On Error Resume Next
        rng.Validation.Delete
        rng.Validation.Add 3, 1, 1, validationString
        rng.Validation.InCellDropdown = True
        If Err.Number = 0 Then
            ShopDocLog "INFO", "Format", "Added validation to column " & colNum
        Else
            ShopDocLog "WARN", "Format", "Could not add validation to column " & colNum & " - " & Err.Description
        End If
        On Error GoTo 0
    End If
End Sub

' ===== CUSTOMIZE THESE LISTS =====
materialList = Join(Array( _
    "Aluminium", _
    "Titanium", _
    "Steel", _
    "Bronze" _
), ",")

surfaceTypeList = "Finish,Controlled Roughing,Free Roughing"

millingTypeList = Join(Array( _
    "End Milling", _
    "Face Milling", _
    "Drilling", _
    "Reaming", _
    "Turning" _
), ",")

toolTypeList = Join(Array( _
    "Carbide", _
    "HSS", _
    "PCD" _
), ",")

strategyTypeList = Join(Array( _
    "Conventional", _
    "HSM" _
), ",")

On Error Resume Next

' Format as Table
Set objRange = objWorksheet.Range("A1").CurrentRegion
If Err.Number <> 0 Then
    ShopDocLog "ERROR", "Format", "Cannot read CSV table region - " & Err.Number & " " & Err.Description
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit 1
End If

Set tbl = objWorksheet.ListObjects.Add(1, objRange, , 1)
If Err.Number <> 0 Then
    ShopDocLog "WARN", "Format", "Could not create Excel table (continuing without table style) - " & Err.Description
    Err.Clear
Else
    tbl.TableStyle = "TableStyleMedium9"
    tbl.ShowAutoFilter = True
    ShopDocLog "INFO", "Format", "Applied Excel table formatting"
End If

' Center align
objRange.HorizontalAlignment = -4108
objRange.VerticalAlignment = -4108

' AutoFit
objWorksheet.Columns.AutoFit

' Multiline header labels (e.g. LF inside CSV quoted cells): wrap row 1
objWorksheet.Rows(1).WrapText = True
objWorksheet.Rows(1).AutoFit

' Data validation (after ListObject: table creation can clear prior validation)
col = FindColumn(objWorksheet, "Material Type")
If col > 0 Then AddValidation objWorksheet, col, lastRow, materialList

col = FindFinishTypeColumn(objWorksheet)
If col > 0 Then AddValidation objWorksheet, col, lastRow, surfaceTypeList

col = FindColumn(objWorksheet, "Cutter Type")
If col > 0 Then AddValidation objWorksheet, col, lastRow, millingTypeList

col = FindColumn(objWorksheet, "Tool Type (Carbide/HSS/PCD)")
If col > 0 Then AddValidation objWorksheet, col, lastRow, toolTypeList

col = FindColumn(objWorksheet, "Strategy Type")
If col > 0 Then AddValidation objWorksheet, col, lastRow, strategyTypeList

' Freeze panes
objWorksheet.Activate
With objExcel.ActiveWindow
    .SplitRow = 1
    .SplitColumn = 1
    .FreezePanes = True
End With

' Save
objWorkbook.SaveAs xlsxFile, 51

If Err.Number <> 0 Then
    ShopDocLog "ERROR", "XLSX", "Cannot save XLSX file: " & xlsxFile & " - " & Err.Number & " " & Err.Description
    ShopDocLog "ERROR", "Permission", "Save failed - folder may be read-only, file locked, blocked by OneDrive/sync, or antivirus"
    objWorkbook.Close False
    objExcel.Quit
    WScript.Quit 1
End If

ShopDocLog "INFO", "XLSX", "Saved XLSX successfully: " & xlsxFile
WScript.Echo "SUCCESS: " & xlsxFile

objWorkbook.Close False
objExcel.Quit

Set objExcel = Nothing

' CSV is deleted by the post (TCL) after error log is appended - do not delete here

WScript.Quit 0
