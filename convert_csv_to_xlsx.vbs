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

' Find column function
Function FindColumn(ws, colName)
    FindColumn = 0
    For col = 1 To ws.UsedRange.Columns.Count
        If Trim(ws.Cells(1, col).Value) = colName Then
            FindColumn = col
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
    "Steel", _
    "Stainless Steel", _
    "Titanium", _
    "Brass", _
    "Copper", _
    "Plastic" _
), ",")

surfaceTypeList = Join(Array( _
    "Roughing", _
    "Finishing" _
), ",")

millingTypeList = Join(Array( _
    "End milling", _
    "Face milling", _
    "Drilling", _
    "Tapping", _
    "Reaming" _
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

' Find columns and add validation
col = FindColumn(objWorksheet, "Material")
If col > 0 Then AddValidation objWorksheet, col, lastRow, materialList

col = FindColumn(objWorksheet, "Surface Type")
If col > 0 Then AddValidation objWorksheet, col, lastRow, surfaceTypeList

col = FindColumn(objWorksheet, "Milling Type")
If col > 0 Then AddValidation objWorksheet, col, lastRow, millingTypeList

col = FindColumn(objWorksheet, "Tool Type")
If col > 0 Then AddValidation objWorksheet, col, lastRow, toolTypeList

col = FindColumn(objWorksheet, "Strategy Type")
If col > 0 Then AddValidation objWorksheet, col, lastRow, strategyTypeList

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