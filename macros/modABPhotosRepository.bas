Attribute VB_Name = "modABPhotosRepository"
Option Explicit

'=============================================================
' Photo metadata and button catalog helpers
'=============================================================

Public Function PhotosSheet() As Worksheet
    EnsureAutoBerichtSheets
    Set PhotosSheet = ThisWorkbook.Worksheets(SHEET_PHOTOS)
End Function

Public Function ListsSheet() As Worksheet
    EnsureAutoBerichtSheets
    Set ListsSheet = ThisWorkbook.Worksheets(SHEET_LISTS)
End Function

Public Sub EnsurePhotoRecord(fileName As String)
    Dim entry As Scripting.Dictionary
    Set entry = GetPhotoEntry(fileName)
    If entry Is Nothing Then
        Dim newEntry As New Scripting.Dictionary
        newEntry.CompareMode = TextCompare
        newEntry("fileName") = fileName
        newEntry("displayName") = fileName
        newEntry("notes") = ""
        newEntry("tagChapters") = ""
        newEntry("tagCategories") = ""
        newEntry("tagTraining") = ""
        newEntry("tagSubfolders") = ""
        newEntry("preferredLocale") = ""
        newEntry("capturedAt") = ""
        UpsertPhoto newEntry
    End If
End Sub

Public Function GetPhotoEntry(fileName As String) As Scripting.Dictionary
    Dim ws As Worksheet
    Set ws = PhotosSheet()
    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, "fileName", fileName)
    If rowIndex = 0 Then
        Set GetPhotoEntry = Nothing
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
    Set GetPhotoEntry = entry
End Function

Public Sub UpsertPhoto(entry As Scripting.Dictionary)
    Dim ws As Worksheet
    Set ws = PhotosSheet()
    UpsertRow ws, "fileName", entry
End Sub

Public Sub SetPhotoTags(fileName As String, tagField As String, tags As Variant)
    EnsurePhotoRecord fileName
    Dim ws As Worksheet
    Set ws = PhotosSheet()
    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, "fileName", fileName)
    If rowIndex = 0 Then Exit Sub

    Dim colIndex As Long
    colIndex = HeaderIndex(ws, tagField)
    If colIndex = 0 Then Exit Sub

    ws.Cells(rowIndex, colIndex).Value = JoinTags(tags)
End Sub

Public Sub TogglePhotoTag(fileName As String, tagField As String, tagValue As String)
    EnsurePhotoRecord fileName
    Dim current As Collection
    Set current = GetTagList(fileName, tagField)

    Dim dict As New Scripting.Dictionary
    dict.CompareMode = TextCompare

    Dim i As Long
    For i = 1 To current.Count
        dict(current(i)) = True
    Next i

    If dict.Exists(tagValue) Then
        dict.Remove tagValue
    Else
        dict(tagValue) = True
    End If

    SetPhotoTags fileName, tagField, dict.Keys
End Sub

Public Function GetTagList(fileName As String, tagField As String) As Collection
    Dim entry As Scripting.Dictionary
    Set entry = GetPhotoEntry(fileName)
    Dim result As New Collection
    If entry Is Nothing Then
        Set GetTagList = result
        Exit Function
    End If

    Dim raw As String
    raw = NzString(entry(tagField))
    If Len(raw) = 0 Then
        Set GetTagList = result
        Exit Function
    End If

    Dim parts() As String
    parts = Split(raw, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim val As String
        val = Trim$(parts(i))
        If Len(val) > 0 Then
            result.Add val
        End If
    Next i
    Set GetTagList = result
End Function

Public Function GetButtonList(listName As String, locale As String) As Collection
    Dim ws As Worksheet
    Set ws = ListsSheet()

    Dim listColumn As Long
    listColumn = HeaderIndex(ws, "listName")
    If listColumn = 0 Then
        Set GetButtonList = New Collection
        Exit Function
    End If

    Dim valueCol As Long
    valueCol = HeaderIndex(ws, "value")
    Dim groupCol As Long
    groupCol = HeaderIndex(ws, "group")
    Dim sortCol As Long
    sortCol = HeaderIndex(ws, "sortOrder")
    Dim chapterCol As Long
    chapterCol = HeaderIndex(ws, "chapterId")

    Dim labelCol As Long
    labelCol = HeaderIndex(ws, "label_" & Replace(locale, "-", "_"))
    If labelCol = 0 Then labelCol = HeaderIndex(ws, "label_de")

    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, listColumn).End(xlUp).Row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        If StrComp(NzString(ws.Cells(r, listColumn).Value), listName, vbTextCompare) = 0 Then
            Dim item As New Scripting.Dictionary
            item.CompareMode = TextCompare
            item("value") = NzString(ws.Cells(r, valueCol).Value)
            item("label") = NzString(ws.Cells(r, labelCol).Value)
            If groupCol > 0 Then item("group") = NzString(ws.Cells(r, groupCol).Value)
            If sortCol > 0 Then item("sortOrder") = ws.Cells(r, sortCol).Value
            If chapterCol > 0 Then item("chapterId") = NzString(ws.Cells(r, chapterCol).Value)
            result.Add item
        End If
    Next r

    SortButtonCollection result
    Set GetButtonList = result
End Function

Private Sub SortButtonCollection(items As Collection)
    Dim i As Long, j As Long
    For i = 1 To items.Count - 1
        For j = i + 1 To items.Count
            Dim a As Scripting.Dictionary, b As Scripting.Dictionary
            Set a = items(i)
            Set b = items(j)
            Dim aOrder As Double, bOrder As Double
            aOrder = NzNumber(GetDictValue(a, "sortOrder"), 99999)
            bOrder = NzNumber(GetDictValue(b, "sortOrder"), 99999)
            If aOrder > bOrder Then
                items.Remove j
                items.Add b, , i
            End If
        Next j
    Next i
End Sub

Public Function JoinTags(tags As Variant) As String
    Dim result As String
    result = ""

    If TypeName(tags) = "Collection" Then
        Dim count As Long
        count = tags.Count
        If count = 0 Then
            JoinTags = ""
            Exit Function
        End If
        Dim temp() As String
        ReDim temp(1 To count)
        Dim i As Long
        For i = 1 To count
            temp(i) = CStr(tags(i))
        Next i
        result = Join(temp, ", ")
    ElseIf IsArray(tags) Then
        Dim lb As Long, ub As Long
        lb = LBound(tags)
        ub = UBound(tags)
        If ub - lb + 1 <= 0 Then
            JoinTags = ""
            Exit Function
        End If
        Dim parts() As String
        ReDim parts(lb To ub)
        Dim idx As Long
        For idx = lb To ub
            parts(idx) = CStr(tags(idx))
        Next idx
        result = Join(parts, ", ")
    Else
        result = NzString(tags)
    End If

    JoinTags = result
End Function

Public Sub NormalizePhotoButtonKeys()
    NormalizeListKeys PHOTO_LIST_BERICHT, syncChapterId:=True
End Sub

Private Sub NormalizeListKeys(listName As String, Optional syncChapterId As Boolean = False)
    Dim ws As Worksheet
    Set ws = ListsSheet()

    Dim colListName As Long
    Dim colValue As Long
    Dim colLabelDe As Long
    Dim colChapterId As Long

    colListName = HeaderIndex(ws, "listName")
    colValue = HeaderIndex(ws, "value")
    colLabelDe = HeaderIndex(ws, "label_de")
    colChapterId = HeaderIndex(ws, "chapterId")

    If colListName = 0 Or colValue = 0 Or colLabelDe = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colListName).End(xlUp).Row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        If StrComp(NzString(ws.Cells(r, colListName).Value), listName, vbTextCompare) = 0 Then
            Dim labelText As String
            labelText = NzString(ws.Cells(r, colLabelDe).Value)
            Dim guessedId As String
            guessedId = GuessListId(labelText)
            If Len(guessedId) = 0 Then GoTo ContinueRow

            If StrComp(NzString(ws.Cells(r, colValue).Value), guessedId, vbTextCompare) <> 0 Then
                ws.Cells(r, colValue).Value = guessedId
            End If

            If syncChapterId And colChapterId > 0 Then
                If StrComp(NzString(ws.Cells(r, colChapterId).Value), guessedId, vbTextCompare) <> 0 Then
                    ws.Cells(r, colChapterId).Value = guessedId
                End If
            End If
        End If
ContinueRow:
    Next r
End Sub

Private Function GuessListId(ByVal labelText As String) As String
    Dim candidate As String
    candidate = Trim$(labelText)
    If Len(candidate) = 0 Then Exit Function

    Dim firstToken As String
    Dim spacePos As Long
    spacePos = InStr(candidate, " ")
    If spacePos > 0 Then
        firstToken = Left$(candidate, spacePos - 1)
    Else
        firstToken = candidate
    End If

    firstToken = Replace$(firstToken, Chr$(160), " ")
    firstToken = Trim$(firstToken)
    firstToken = modABIdUtils.NormalizeReportItemId(firstToken)

    If firstToken Like "*[0-9]*" Then
        GuessListId = firstToken
    Else
        GuessListId = ""
    End If
End Function

EOF
