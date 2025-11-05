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
        newEntry("tagBericht") = ""
        newEntry("tagSeminar") = ""
        newEntry("tagTopic") = ""
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
            Dim item As Scripting.Dictionary
            Set item = New Scripting.Dictionary
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

Public Function BuildFolderTagLookup() As Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    map.CompareMode = TextCompare

    Dim ws As Worksheet
    Set ws = ListsSheet()

    Dim colListName As Long
    Dim colValue As Long
    Dim colLabelDe As Long
    Dim colLabelFr As Long
    Dim colLabelIt As Long
    Dim colLabelEn As Long

    colListName = HeaderIndex(ws, "listName")
    colValue = HeaderIndex(ws, "value")
    colLabelDe = HeaderIndex(ws, "label_de")
    colLabelFr = HeaderIndex(ws, "label_fr")
    colLabelIt = HeaderIndex(ws, "label_it")
    colLabelEn = HeaderIndex(ws, "label_en")

    If colListName = 0 Then
        Set BuildFolderTagLookup = map
        Exit Function
    End If

    Dim listToField As New Scripting.Dictionary
    listToField.CompareMode = TextCompare
    listToField(PHOTO_LIST_BERICHT) = PHOTO_TAG_BERICHT
    listToField(PHOTO_LIST_SEMINAR) = PHOTO_TAG_SEMINAR
    listToField(PHOTO_LIST_TOPIC) = PHOTO_TAG_TOPIC

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colListName).End(xlUp).Row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim listName As String
        listName = NzString(ws.Cells(r, colListName).Value)
        If Not listToField.Exists(listName) Then GoTo ContinueRow

        Dim tagField As String
        tagField = CStr(listToField(listName))

        Dim tagValue As String
        If colValue > 0 Then
            tagValue = NzString(ws.Cells(r, colValue).Value)
        Else
            tagValue = ""
        End If
        If Len(tagValue) = 0 Then GoTo ContinueRow

        AddFolderMapping map, NzString(ws.Cells(r, colLabelDe).Value), tagField, tagValue
        AddFolderMapping map, NzString(ws.Cells(r, colLabelFr).Value), tagField, tagValue
        AddFolderMapping map, NzString(ws.Cells(r, colLabelIt).Value), tagField, tagValue
        AddFolderMapping map, NzString(ws.Cells(r, colLabelEn).Value), tagField, tagValue
ContinueRow:
    Next r

    Set BuildFolderTagLookup = map
End Function

Public Sub ApplyFolderTags(record As Scripting.Dictionary, ByVal relativePath As String, folderMap As Scripting.Dictionary)
    If folderMap Is Nothing Then Exit Sub
    If folderMap.Count = 0 Then Exit Sub
    If record Is Nothing Then Exit Sub

    Dim normalizedPath As String
    normalizedPath = Replace(relativePath, "/", "\")

    Dim segments() As String
    segments = Split(normalizedPath, "\")
    If UBound(segments) < 1 Then Exit Sub

    Dim tagBuckets As New Scripting.Dictionary
    tagBuckets.CompareMode = TextCompare

    Dim fields As Variant
    fields = Array(PHOTO_TAG_BERICHT, PHOTO_TAG_SEMINAR, PHOTO_TAG_TOPIC)

    Dim field As Variant
    For Each field In fields
        Dim initialBucket As Scripting.Dictionary
        Set initialBucket = ExistingTagDictionary(record, CStr(field))
        tagBuckets(field) = initialBucket
    Next field

    Dim i As Long
    For i = LBound(segments) To UBound(segments) - 1
        Dim folderToken As String
        folderToken = NormalizeFolderName(segments(i))
        If Len(folderToken) = 0 Then GoTo ContinueSegment
        Dim lookupKey As String
        lookupKey = LCase$(folderToken)
        If folderMap.Exists(lookupKey) Then
            Dim entries As Collection
            Set entries = folderMap(lookupKey)
            Dim desc As Scripting.Dictionary
            For Each desc In entries
                tagBuckets(desc("field"))(desc("value")) = True
            Next desc
        End If
ContinueSegment:
    Next i

    Dim bucket As Scripting.Dictionary
    For Each field In tagBuckets.Keys
        Set bucket = tagBuckets(field)
        record(CStr(field)) = JoinDictionaryKeys(bucket)
    Next field
End Sub

Private Sub AddFolderMapping(ByRef map As Scripting.Dictionary, ByVal labelValue As Variant, ByVal tagField As String, ByVal tagValue As String)
    Dim token As String
    token = NormalizeFolderName(labelValue)
    If Len(token) = 0 Then Exit Sub

    Dim key As String
    key = LCase$(token)

    Dim entries As Collection
    If map.Exists(key) Then
        Set entries = map(key)
    Else
        Set entries = New Collection
        Set map(key) = entries
    End If

    Dim existing As Scripting.Dictionary
    For Each existing In entries
        If StrComp(existing("field"), tagField, vbTextCompare) = 0 _
            And StrComp(existing("value"), tagValue, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next existing

    Dim descriptor As New Scripting.Dictionary
    descriptor.CompareMode = TextCompare
    descriptor("field") = tagField
    descriptor("value") = tagValue
    entries.Add descriptor
End Sub

Private Function NormalizeFolderName(ByVal rawValue As Variant) As String
    Dim textValue As String
    textValue = NzString(rawValue)
    textValue = Replace(textValue, Chr$(160), " ")
    textValue = Trim$(textValue)
    Do While InStr(textValue, "  ") > 0
        textValue = Replace(textValue, "  ", " ")
    Loop
    NormalizeFolderName = textValue
End Function

Private Function ExistingTagDictionary(ByVal record As Scripting.Dictionary, ByVal fieldName As String) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.CompareMode = TextCompare

    If Not record Is Nothing Then
        If record.Exists(fieldName) Then
            Dim raw As String
            raw = NzString(record(fieldName))
            If Len(raw) > 0 Then
                Dim parts() As String
                parts = Split(raw, ",")
                Dim idx As Long
                For idx = LBound(parts) To UBound(parts)
                    Dim trimmed As String
                    trimmed = Trim$(parts(idx))
                    If Len(trimmed) > 0 Then dict(trimmed) = True
                Next idx
            End If
        End If
    End If

    Set ExistingTagDictionary = dict
End Function

Private Function JoinDictionaryKeys(ByVal dict As Scripting.Dictionary) As String
    If dict Is Nothing Then Exit Function
    If dict.Count = 0 Then
        JoinDictionaryKeys = ""
    Else
        Dim keys As Variant
        keys = dict.Keys
        Dim i As Long, j As Long
        For i = LBound(keys) To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                    Dim tmp As Variant
                    tmp = keys(i)
                    keys(i) = keys(j)
                    keys(j) = tmp
                End If
            Next j
        Next i
        JoinDictionaryKeys = Join(keys, ", ")
    End If
End Function
