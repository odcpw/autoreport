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

Public Function EnsureRowRecord(rowId As String, Optional chapterId As String = "") As Long
    Dim ws As Worksheet
    Set ws = RowsSheet()

    Dim normalized As String
    normalized = NormalizeReportItemId(rowId)
    If Len(normalized) = 0 Then Exit Function

    Dim colRowId As Long
    colRowId = HeaderIndex(ws, "rowId")
    If colRowId = 0 Then Exit Function

    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, "rowId", normalized)
    If rowIndex = 0 Then
        rowIndex = ws.Cells(ws.Rows.Count, colRowId).End(xlUp).Row + 1
        InitializeRowDefaults ws, rowIndex, normalized, chapterId
    ElseIf Len(chapterId) > 0 Then
        Dim colChapterId As Long
        colChapterId = HeaderIndex(ws, "chapterId")
        If colChapterId > 0 Then ws.Cells(rowIndex, colChapterId).Value = chapterId
    End If

    EnsureRowRecord = rowIndex
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

Private Sub InitializeRowDefaults(ws As Worksheet, rowIndex As Long, _
    ByVal rowId As String, ByVal explicitChapterId As String)

    Dim chapterId As String
    chapterId = explicitChapterId
    If Len(chapterId) = 0 Then chapterId = ParentChapterId(rowId)

    Dim colRowId As Long
    colRowId = HeaderIndex(ws, "rowId")
    If colRowId > 0 Then ws.Cells(rowIndex, colRowId).Value = rowId

    Dim colChapterId As Long
    colChapterId = HeaderIndex(ws, "chapterId")
    If colChapterId > 0 Then ws.Cells(rowIndex, colChapterId).Value = chapterId

    Dim colMasterFinding As Long
    colMasterFinding = HeaderIndex(ws, "masterFinding")
    If colMasterFinding > 0 Then ws.Cells(rowIndex, colMasterFinding).Value = ""

    Dim i As Long
    For i = 1 To 4
        Dim colMasterLevel As Long
        colMasterLevel = HeaderIndex(ws, "masterLevel" & CStr(i))
        If colMasterLevel > 0 Then ws.Cells(rowIndex, colMasterLevel).Value = ""

        Dim colOverrideLevel As Long
        colOverrideLevel = HeaderIndex(ws, "overrideLevel" & CStr(i))
        If colOverrideLevel > 0 Then ws.Cells(rowIndex, colOverrideLevel).Value = ""

        Dim colUseOverrideLevel As Long
        colUseOverrideLevel = HeaderIndex(ws, "useOverrideLevel" & CStr(i))
        If colUseOverrideLevel > 0 Then ws.Cells(rowIndex, colUseOverrideLevel).Value = False
    Next i

    Dim colOverrideFinding As Long
    colOverrideFinding = HeaderIndex(ws, "overrideFinding")
    If colOverrideFinding > 0 Then ws.Cells(rowIndex, colOverrideFinding).Value = ""

    Dim colUseOverrideFinding As Long
    colUseOverrideFinding = HeaderIndex(ws, "useOverrideFinding")
    If colUseOverrideFinding > 0 Then ws.Cells(rowIndex, colUseOverrideFinding).Value = False

    Dim colIncludeFinding As Long
    colIncludeFinding = HeaderIndex(ws, "includeFinding")
    If colIncludeFinding > 0 Then ws.Cells(rowIndex, colIncludeFinding).Value = True

    Dim colIncludeRecommendation As Long
    colIncludeRecommendation = HeaderIndex(ws, "includeRecommendation")
    If colIncludeRecommendation > 0 Then ws.Cells(rowIndex, colIncludeRecommendation).Value = True

    Dim colSelectedLevel As Long
    colSelectedLevel = HeaderIndex(ws, "selectedLevel")
    If colSelectedLevel > 0 Then ws.Cells(rowIndex, colSelectedLevel).Value = 2

    Dim colOverwriteMode As Long
    colOverwriteMode = HeaderIndex(ws, "overwriteMode")
    If colOverwriteMode > 0 Then ws.Cells(rowIndex, colOverwriteMode).Value = "append"

    Dim colDone As Long
    colDone = HeaderIndex(ws, "done")
    If colDone > 0 Then ws.Cells(rowIndex, colDone).Value = False

    Dim colNotes As Long
    colNotes = HeaderIndex(ws, "notes")
    If colNotes > 0 Then ws.Cells(rowIndex, colNotes).Value = ""

    Dim colCustomerAnswer As Long
    colCustomerAnswer = HeaderIndex(ws, "customerAnswer")
    If colCustomerAnswer > 0 Then ws.Cells(rowIndex, colCustomerAnswer).Value = Empty

    Dim colCustomerRemark As Long
    colCustomerRemark = HeaderIndex(ws, "customerRemark")
    If colCustomerRemark > 0 Then ws.Cells(rowIndex, colCustomerRemark).Value = ""

    Dim colCustomerPriority As Long
    colCustomerPriority = HeaderIndex(ws, "customerPriority")
    If colCustomerPriority > 0 Then ws.Cells(rowIndex, colCustomerPriority).Value = Empty

    Dim colLastEditedBy As Long
    colLastEditedBy = HeaderIndex(ws, "lastEditedBy")
    If colLastEditedBy > 0 Then ws.Cells(rowIndex, colLastEditedBy).Value = ""

    Dim colLastEditedAt As Long
    colLastEditedAt = HeaderIndex(ws, "lastEditedAt")
    If colLastEditedAt > 0 Then ws.Cells(rowIndex, colLastEditedAt).Value = ""
End Sub
