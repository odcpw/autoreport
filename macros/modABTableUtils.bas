Attribute VB_Name = "modABTableUtils"
Option Explicit

'=============================================================
' Generic helpers for working with structured tables
'=============================================================

Public Function HeaderIndex(ws As Worksheet, columnName As String) As Long
    Dim headers As Variant
    headers = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, ws.UsedRange.Columns.Count)).Value
    If TypeName(headers) = "Variant()" Then
        Dim i As Long
        For i = LBound(headers, 2) To UBound(headers, 2)
            If StrComp(CStr(headers(1, i)), columnName, vbTextCompare) = 0 Then
                HeaderIndex = i
                Exit Function
            End If
        Next i
    ElseIf Not IsEmpty(headers) Then
        If StrComp(CStr(headers), columnName, vbTextCompare) = 0 Then
            HeaderIndex = 1
            Exit Function
        End If
    End If
    HeaderIndex = 0
End Function

Public Function ReadTableAsCollection(ws As Worksheet) As Collection
    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= ROW_HEADER_ROW Then
        Set ReadTableAsCollection = result
        Exit Function
    End If

    Dim headers As Variant
    headers = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, ws.UsedRange.Columns.Count)).Value

    Dim rowValues As Variant
    rowValues = ws.Range(ws.Cells(ROW_HEADER_ROW + 1, 1), ws.Cells(lastRow, ws.UsedRange.Columns.Count)).Value

    Dim r As Long, c As Long
    For r = LBound(rowValues, 1) To UBound(rowValues, 1)
        Dim entry As New Scripting.Dictionary
        entry.CompareMode = TextCompare
        For c = LBound(headers, 2) To UBound(headers, 2)
            entry(CStr(headers(1, c))) = rowValues(r, c)
        Next c
        result.Add entry
    Next r

    Set ReadTableAsCollection = result
End Function

Public Sub WriteCollectionToTable(ws As Worksheet, data As Collection)
    If data Is Nothing Then Exit Sub

    Dim headers As Variant
    headers = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, ws.UsedRange.Columns.Count)).Value
    Dim colCount As Long
    colCount = UBound(headers, 2)
    Dim rowCount As Long
    rowCount = data.Count

    If rowCount = 0 Then
        ClearTableRange ws
        Exit Sub
    End If

    Dim output() As Variant
    ReDim output(1 To rowCount, 1 To colCount)

    Dim i As Long, c As Long
    For i = 1 To rowCount
        Dim entry As Scripting.Dictionary
        Set entry = data(i)
        For c = 1 To colCount
            Dim key As String
            key = CStr(headers(1, c))
            If entry.Exists(key) Then
                output(i, c) = entry(key)
            Else
                output(i, c) = Empty
            End If
        Next c
    Next i

    ClearTableRange ws
    ws.Cells(ROW_HEADER_ROW + 1, 1).Resize(rowCount, colCount).Value = output
End Sub

Public Function FindRowIndex(ws As Worksheet, keyColumn As String, keyValue As String) As Long
    Dim colIndex As Long
    colIndex = HeaderIndex(ws, keyColumn)
    If colIndex = 0 Then
        FindRowIndex = 0
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    Dim i As Long
    For i = ROW_HEADER_ROW + 1 To lastRow
        If StrComp(CStr(ws.Cells(i, colIndex).Value), keyValue, vbTextCompare) = 0 Then
            FindRowIndex = i
            Exit Function
        End If
    Next i
    FindRowIndex = 0
End Function

Public Sub UpsertRow(ws As Worksheet, keyColumn As String, entry As Scripting.Dictionary)
    Dim colIndex As Long
    colIndex = HeaderIndex(ws, keyColumn)
    If colIndex = 0 Then Exit Sub

    Dim keyValue As String
    keyValue = NzString(entry(keyColumn))

    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, keyColumn, keyValue)
    If rowIndex = 0 Then
        rowIndex = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row + 1
    End If

    Dim headers As Variant
    headers = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, ws.UsedRange.Columns.Count)).Value
    Dim c As Long
    For c = LBound(headers, 2) To UBound(headers, 2)
        Dim key As String
        key = CStr(headers(1, c))
        If entry.Exists(key) Then
            ws.Cells(rowIndex, c).Value = entry(key)
        Else
            ws.Cells(rowIndex, c).Value = Empty
        End If
    Next c
End Sub

Public Function NzString(value As Variant) As String
    If IsMissing(value) Then
        NzString = ""
    ElseIf IsNull(value) Or IsEmpty(value) Then
        NzString = ""
    Else
        NzString = CStr(value)
    End If
End Function

Public Function NzBool(value As Variant, Optional defaultValue As Boolean = False) As Boolean
    If IsMissing(value) Or IsNull(value) Or value = "" Then
        NzBool = defaultValue
    ElseIf VarType(value) = vbBoolean Then
        NzBool = value
    ElseIf IsNumeric(value) Then
        NzBool = (CLng(value) <> 0)
    Else
        NzBool = (LCase$(CStr(value)) = "true")
    End If
End Function

Public Function NzNumber(value As Variant, Optional defaultValue As Double = 0#) As Double
    If IsMissing(value) Or IsNull(value) Or value = "" Then
        NzNumber = defaultValue
    ElseIf IsNumeric(value) Then
        NzNumber = CDbl(value)
    Else
        NzNumber = defaultValue
    End If
End Function

Public Function GetDictValue(dict As Dictionary, key As String, Optional defaultValue As Variant) As Variant
    If dict Is Nothing Then
        GetDictValue = defaultValue
    ElseIf dict.Exists(key) Then
        GetDictValue = dict(key)
    Else
        GetDictValue = defaultValue
    End If
End Function

Public Sub ClearTableRange(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > ROW_HEADER_ROW Then
        ws.Rows(ROW_HEADER_ROW + 1 & ":" & lastRow).ClearContents
    End If
End Sub
