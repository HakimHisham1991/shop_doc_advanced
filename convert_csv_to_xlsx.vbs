' =============================================
' convert_csv_to_xlsx.vbs - CSV to XLSX + Table + Freeze + Center Alignment + Delete CSV
' =============================================

If WScript.Arguments.Count < 1 Then
    WScript.Echo "Usage: convert_csv_to_xlsx.vbs input.csv [output.xlsx]"
    WScript.Quit
End If

csvFile = WScript.Arguments(0)

' Determine output xlsx filename
If WScript.Arguments.Count > 1 Then
    xlsxFile = WScript.Arguments(1)
Else
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    xlsxFile = fso.GetParentFolderName(csvFile) & "\" & fso.GetBaseName(csvFile) & ".xlsx"
End If

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

' Open CSV
Set objWorkbook = objExcel.Workbooks.Open(csvFile)
Set objWorksheet = objWorkbook.Worksheets(1)

' Select all data
Set objRange = objWorksheet.Range("A1").CurrentRegion

' === Format as Table ===
Dim tbl
Set tbl = objWorksheet.ListObjects.Add(1, objRange, , 1)

tbl.TableStyle = "TableStyleMedium9"
tbl.ShowAutoFilter = True

' === Center Alignment (Horizontal + Vertical) ===
objRange.HorizontalAlignment = -4108   ' xlCenter
objRange.VerticalAlignment   = -4108   ' xlCenter

' AutoFit columns
objWorksheet.Columns.AutoFit

' === Freeze Top Row and First Column ===
objWorksheet.Activate
With objExcel.ActiveWindow
    .SplitRow = 1
    .SplitColumn = 1
    .FreezePanes = True
End With

' Save as XLSX
objWorkbook.SaveAs xlsxFile, 51

' Clean shutdown
objWorkbook.Close False
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing

WScript.Sleep 800

' === Delete CSV with retry ===
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim attempts : attempts = 0
Do While objFSO.FileExists(csvFile) And attempts < 5
    On Error Resume Next
    objFSO.DeleteFile csvFile, True
    On Error Goto 0
    If objFSO.FileExists(csvFile) Then
        WScript.Sleep 300
        attempts = attempts + 1
    Else
        Exit Do
    End If
Loop

If objFSO.FileExists(csvFile) Then
    WScript.Echo "Success: " & xlsxFile & " (CSV could not be deleted)"
Else
    WScript.Echo "Success: " & xlsxFile & " (Table + Center + Frozen Panes)"
End If