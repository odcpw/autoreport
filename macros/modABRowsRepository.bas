Attribute VB_Name = "modABRowsRepository"
Option Explicit

'=============================================================
' Accessors for the Rows sheet
'=============================================================

Public Function RowsSheet() As Worksheet
    EnsureAutoBerichtSheets
    Set RowsSheet = ThisWorkbook.Worksheets(SHEET_ROWS)
End Function

Public Function GetRowById(rowId As String) As Scripting.Dictionary
    Dim ws As Worksheet
    Set ws = RowsSheet()
    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, "rowId", rowId)
    If rowIndex = 0 Then
        Set GetRowById = Nothing
        Exit Function
    End If

    Dim headers As Variant
    headers = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, ws.UsedRange.Columns.Count)).Value

    Dim entry As New Scripting.Dictionary
    entry.CompareMode = TextCompare

    Dim c As Long
    For c = LBound(headers, 2) To UBound(headers, 2)
        entry(CStr(headers(1, c))) = ws.Cells(rowIndex, c).Value
    Next c

    Set GetRowById = entry
End Function

Public Sub UpsertRowEntry(entry As Scripting.Dictionary)
    Dim ws As Worksheet
    Set ws = RowsSheet()
    UpsertRow ws, "rowId", entry
End Sub

Public Sub UpdateRowField(rowId As String, fieldName As String, value As Variant)
    Dim ws As Worksheet
    Set ws = RowsSheet()
    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, "rowId", rowId)
    If rowIndex = 0 Then Exit Sub

    Dim colIndex As Long
    colIndex = HeaderIndex(ws, fieldName)
    If colIndex = 0 Then Exit Sub

    Dim oldValue As Variant
    oldValue = ws.Cells(rowIndex, colIndex).Value
    ws.Cells(rowIndex, colIndex).Value = value
    RecordOverride rowId, fieldName, oldValue, value
End Sub

Public Function AllRows() As Collection
    Set AllRows = ReadTableAsCollection(RowsSheet())
End Function

Public Sub TouchRow(rowId As String)
    UpdateRowField rowId, "lastEditedAt", Now
    UpdateRowField rowId, "lastEditedBy", Environ$("USERNAME")
End Sub

Public Sub SetFindingOverride(rowId As String, textValue As String, enableOverride As Boolean)
    UpdateRowField rowId, "overrideFinding", textValue
    UpdateRowField rowId, "useOverrideFinding", NzBool(enableOverride)
    TouchRow rowId
End Sub

Public Sub SetRecommendationOverride(rowId As String, level As Long, textValue As String, enableOverride As Boolean)
    Dim levelKey As String
    levelKey = CStr(level)
    UpdateRowField rowId, "overrideLevel" & levelKey, textValue
    UpdateRowField rowId, "useOverrideLevel" & levelKey, NzBool(enableOverride)
    UpdateRowField rowId, "selectedLevel", level
    TouchRow rowId
End Sub

Public Sub SetIncludeFlags(rowId As String, includeFinding As Boolean, includeRecommendation As Boolean)
    UpdateRowField rowId, "includeFinding", NzBool(includeFinding)
    UpdateRowField rowId, "includeRecommendation", NzBool(includeRecommendation)
    TouchRow rowId
End Sub

Private Sub RecordOverride(rowId As String, fieldName As String, oldValue As Variant, newValue As Variant)
    If NzString(oldValue) = NzString(newValue) Then Exit Sub
    Dim ws As Worksheet
    EnsureAutoBerichtSheets
    Set ws = ThisWorkbook.Worksheets(SHEET_OVERRIDES_HISTORY)

    Dim entry As New Scripting.Dictionary
    entry.CompareMode = TextCompare
    entry("timestamp") = Now
    entry("rowId") = rowId
    entry("fieldName") = fieldName
    entry("oldValue") = oldValue
    entry("newValue") = newValue
    entry("user") = Environ$("USERNAME")

    AppendRow ws, entry
End Sub

Private Sub AppendRow(ws As Worksheet, entry As Scripting.Dictionary)
    Dim headers As Variant
    headers = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, ws.UsedRange.Columns.Count)).Value
    Dim colCount As Long
    colCount = UBound(headers, 2)

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    Dim c As Long
    For c = 1 To colCount
        Dim key As String
        key = CStr(headers(1, c))
        If entry.Exists(key) Then
            ws.Cells(nextRow, c).Value = entry(key)
        Else
            ws.Cells(nextRow, c).Value = Empty
        End If
    Next c
End Sub
