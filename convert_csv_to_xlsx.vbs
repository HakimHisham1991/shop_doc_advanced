' =============================================
' convert_csv_to_xlsx.vbs - With Data Validation
' =============================================

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: convert_csv_to_xlsx.vbs input.csv [output.xlsx]"
    WScript.Quit 1
End If

csvFile = WScript.Arguments(0)

If WScript.Arguments.Count > 1 Then
    xlsxFile = WScript.Arguments(1)
Else
    Set fso = CreateObject("Scripting.FileSystemObject")
    xlsxFile = fso.GetParentFolderName(csvFile) & "\" & fso.GetBaseName(csvFile) & ".xlsx"
End If

On Error Resume Next

Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot create Excel object"
    WScript.Quit 1
End If

objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open(csvFile)
If Err.Number <> 0 Then
    WScript.Echo "ERROR: Cannot open " & csvFile
    objExcel.Quit
    WScript.Quit 1
End If

Set objWorksheet = objWorkbook.Worksheets(1)

' Get the last row
lastRow = objWorksheet.UsedRange.Rows.Count
WScript.Echo "CSV has " & lastRow & " rows"

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
            WScript.Echo "  Added validation to column " & colNum
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

' Format as Table
Set objRange = objWorksheet.Range("A1").CurrentRegion
Set tbl = objWorksheet.ListObjects.Add(1, objRange, , 1)
tbl.TableStyle = "TableStyleMedium9"
tbl.ShowAutoFilter = True

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
    WScript.Echo "ERROR: Cannot save to " & xlsxFile & " - " & Err.Description
    WScript.Quit 1
End If

WScript.Echo "SUCCESS: " & xlsxFile

objWorkbook.Close False
objExcel.Quit

' Delete CSV
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
objFSO.DeleteFile csvFile, True

Set objExcel = Nothing