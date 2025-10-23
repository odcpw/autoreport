Attribute VB_Name = "modABWorkbookSetup"
Option Explicit

'=============================================================
' AutoBericht workbook bootstrap helpers
'=============================================================

Public Sub EnsureAutoBerichtSheets(Optional ByVal clearExisting As Boolean = False)
    Dim sheetName As Variant
    For Each sheetName In SheetList()
        EnsureSheetWithHeaders CStr(sheetName), SheetHeaders(CStr(sheetName)), clearExisting
    Next sheetName
End Sub

Public Function EnsureSheetWithHeaders(ByVal sheetName As String, _
                                       ByVal headers As Variant, _
                                       Optional ByVal clearExisting As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    If clearExisting Then
        ws.Cells.Clear
    End If

    If Not IsEmpty(headers) Then
        Dim existingHeader As Variant
        existingHeader = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, UBound(headers) + 1)).Value
        If clearExisting Or Not HeaderMatches(existingHeader, headers) Then
            WriteHeaderRow ws, headers
        End If
    End If

    ws.Rows(ROW_HEADER_ROW).Font.Bold = True
    ws.Columns.AutoFit
    Set EnsureSheetWithHeaders = ws
End Function

Public Sub WriteHeaderRow(ws As Worksheet, headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(ROW_HEADER_ROW, i + 1).Value = headers(i)
    Next i
End Sub

Private Function HeaderMatches(existing As Variant, headers As Variant) As Boolean
    On Error GoTo Mismatch
    Dim cols As Long
    Dim i As Long
    cols = UBound(headers) - LBound(headers) + 1
    If TypeName(existing) = "Variant()" Then
        If UBound(existing, 2) - LBound(existing, 2) + 1 <> cols Then GoTo Mismatch
        For i = 1 To cols
            If CStr(existing(1, i)) <> CStr(headers(i - 1)) Then GoTo Mismatch
        Next i
        HeaderMatches = True
        Exit Function
    ElseIf Not IsEmpty(existing) Then
        HeaderMatches = (cols = 1 And CStr(existing) = CStr(headers(0)))
        Exit Function
    End If
Mismatch:
    HeaderMatches = False
End Function

Public Sub ClearDataTables()
    Dim sheetName As Variant
    For Each sheetName In Array(SHEET_ROWS, SHEET_PHOTOS, SHEET_LISTS, SHEET_EXPORT_LOG, SHEET_OVERRIDES_HISTORY)
        ClearTable ThisWorkbook.Worksheets(CStr(sheetName))
    Next sheetName
End Sub

Private Sub ClearTable(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > ROW_HEADER_ROW Then
        ws.Rows(ROW_HEADER_ROW + 1 & ":" & lastRow).Delete
    End If
End Sub

Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

