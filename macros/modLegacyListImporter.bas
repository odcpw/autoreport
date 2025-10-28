Attribute VB_Name = "modLegacyListImporter"
Option Explicit

'=============================================================
' Legacy PSCategoryLabels → Lists importer
'=============================================================
' Converts the columns from the old PSCategoryLabels sheet to
' rows in the new Lists table.
'
' Expected legacy layout:
'   Column A – Bericht buttons (chapter mapped)
'   Column B – Audit buttons
'   Column C – Training buttons
'   Column D – Subfolder labels
'
' Adjust LIST_MAPPING below if your columns differ.
'=============================================================

Private Type LegacyColumnMap
    ColumnIndex As Long
    listName As String
    tagField As String
    UseChapterId As Boolean
End Type

Private Const LEGACY_SHEET_NAME As String = "PSCategoryLabels"

Private LegacyMappings() As LegacyColumnMap

Public Sub ImportLegacyCategoryLabels()
    Dim wsLegacy As Worksheet
    On Error Resume Next
    Set wsLegacy = ThisWorkbook.Worksheets(LEGACY_SHEET_NAME)
    On Error GoTo 0

    If wsLegacy Is Nothing Then
        MsgBox "Legacy sheet '" & LEGACY_SHEET_NAME & "' not found.", vbExclamation
        Exit Sub
    End If

    EnsureAutoBerichtSheets

    Dim wsLists As Worksheet
    Set wsLists = ThisWorkbook.Worksheets(SHEET_LISTS)

    Dim wsChapters As Worksheet
    On Error Resume Next
    Set wsChapters = ThisWorkbook.Worksheets("BerichtKapitel")
    On Error GoTo 0

    PrepareMappings

    Dim lastRow As Long
    lastRow = wsLegacy.Cells(wsLegacy.Rows.count, 1).End(xlUp).row

    Dim chapterMap As Scripting.Dictionary
    Set chapterMap = BuildChapterLookup(wsChapters)

    Dim output As Scripting.Dictionary
    Set output = CreateObject("Scripting.Dictionary")
    output.CompareMode = vbTextCompare

    Dim r As Long
    For r = 1 To lastRow
        Dim mapIndex As Long
        For mapIndex = LBound(LegacyMappings) To UBound(LegacyMappings)
            Dim map As LegacyColumnMap
            map = LegacyMappings(mapIndex)
            Dim labelText As String
            labelText = Trim$(NzAny(wsLegacy.Cells(r, map.ColumnIndex).value))
            If Len(labelText) = 0 Then GoTo NextLabel

            Dim key As String
            key = map.listName & "|" & r
            Dim entry As Scripting.Dictionary
            If output.Exists(key) Then
                Set entry = output(key)
            Else
                Set entry = CreateObject("Scripting.Dictionary")
                entry.CompareMode = vbTextCompare
                entry("listName") = map.listName
                entry("sortOrder") = r
                Set output(key) = entry
            End If

            entry("label_de") = labelText
            entry("label_fr") = labelText
            entry("label_it") = labelText
            entry("label_en") = labelText

            Dim value As String
            value = labelText
            Dim chapterId As String
            chapterId = ""

            If map.UseChapterId Then
                chapterId = NormalizeChapterId(Trim$(NzAny(wsLegacy.Cells(r, 1).value)))
                If Len(chapterId) = 0 And Not chapterMap Is Nothing Then
                    chapterId = chapterMapLookup(chapterMap, labelText)
                End If
                If Len(chapterId) = 0 Then
                    chapterId = r
                End If
                value = chapterId
            End If

            entry("value") = value
            entry("chapterId") = chapterId
            entry("group") = map.listName
NextLabel:
        Next mapIndex
    Next r

    WriteListsFromEntries wsLists, output
    MsgBox "Legacy labels imported into Lists sheet.", vbInformation
End Sub

Private Sub PrepareMappings()
    ReDim LegacyMappings(0 To 3)

    LegacyMappings(0).ColumnIndex = 1
    LegacyMappings(0).listName = modABPhotoConstants.PHOTO_LIST_BERICHT
    LegacyMappings(0).tagField = modABPhotoConstants.PHOTO_TAG_CHAPTERS
    LegacyMappings(0).UseChapterId = True

    LegacyMappings(1).ColumnIndex = 2
    LegacyMappings(1).listName = modABPhotoConstants.PHOTO_LIST_AUDIT
    LegacyMappings(1).tagField = modABPhotoConstants.PHOTO_TAG_CATEGORIES
    LegacyMappings(1).UseChapterId = False

    LegacyMappings(2).ColumnIndex = 3
    LegacyMappings(2).listName = modABPhotoConstants.PHOTO_LIST_TRAINING
    LegacyMappings(2).tagField = modABPhotoConstants.PHOTO_TAG_TRAINING
    LegacyMappings(2).UseChapterId = False

    LegacyMappings(3).ColumnIndex = 4
    LegacyMappings(3).listName = modABPhotoConstants.PHOTO_LIST_SUBFOLDERS
    LegacyMappings(3).tagField = modABPhotoConstants.PHOTO_TAG_SUBFOLDERS
    LegacyMappings(3).UseChapterId = False
End Sub

Private Function BuildChapterLookup(wsChapters As Worksheet) As Object
    If wsChapters Is Nothing Then
        Set BuildChapterLookup = Nothing
        Exit Function
    End If

    Dim lookup As Object
    Set lookup = CreateObject("Scripting.Dictionary")
    lookup.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = wsChapters.Cells(wsChapters.Rows.count, 1).End(xlUp).row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim chapterId As String
        chapterId = Trim$(NzAny(wsChapters.Cells(r, HeaderIndex(wsChapters, "chapterId")).value))
        Dim title As String
        title = Trim$(NzAny(wsChapters.Cells(r, HeaderIndex(wsChapters, "defaultTitle_de")).value))
        If Len(chapterId) > 0 Then
            lookup(chapterId) = chapterId
        End If
        If Len(title) > 0 Then
            lookup(title) = chapterId
        End If
    Next r

    Set BuildChapterLookup = lookup
End Function

Private Function HeaderIndex(ws As Worksheet, headerName As String, Optional headerRow As Long = 1) As Long
    ' Finds the column number of the header name (case-insensitive, exact match)
    Dim f As Range
    Set f = ws.Rows(headerRow).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        Err.Raise vbObjectError + 5000, "HeaderIndex", _
            "Header '" & headerName & "' not found in row " & headerRow & " on sheet '" & ws.Name & "'."
    End If
    HeaderIndex = f.column
End Function


Private Function chapterMapLookup(ByVal dict As Object, ByVal key As String) As String
    If dict Is Nothing Then Exit Function
    If dict.Exists(key) Then chapterMapLookup = dict(key)
End Function

Private Sub WriteListsFromEntries(ByVal ws As Worksheet, ByVal entries As Object)
    Dim rowIndex As Long
    rowIndex = ROW_HEADER_ROW + 1

    Dim key As Variant
    For Each key In entries.Keys
        Dim entry As Scripting.Dictionary
        Set entry = entries(key)
        ws.Cells(rowIndex, HeaderIndex(ws, "listName")).value = entry("listName")
        ws.Cells(rowIndex, HeaderIndex(ws, "value")).value = NzAny(entry("value"))
        ws.Cells(rowIndex, HeaderIndex(ws, "label_de")).value = NzAny(entry("label_de"))
        ws.Cells(rowIndex, HeaderIndex(ws, "label_fr")).value = NzAny(entry("label_fr"))
        ws.Cells(rowIndex, HeaderIndex(ws, "label_it")).value = NzAny(entry("label_it"))
        ws.Cells(rowIndex, HeaderIndex(ws, "label_en")).value = NzAny(entry("label_en"))
        ws.Cells(rowIndex, HeaderIndex(ws, "group")).value = NzAny(entry("group"))
        ws.Cells(rowIndex, HeaderIndex(ws, "sortOrder")).value = entry("sortOrder")
        ws.Cells(rowIndex, HeaderIndex(ws, "chapterId")).value = NzAny(entry("chapterId"))
        rowIndex = rowIndex + 1
    Next key
End Sub

Private Function NzAny(ByVal value As Variant) As Variant
    If IsMissing(value) Or IsNull(value) Then
        NzAny = ""
    Else
        NzAny = value
    End If
End Function

Private Function NormalizeChapterId(ByVal rawId As String) As String
    Dim trimmed As String
    trimmed = Trim$(rawId)
    Do While Len(trimmed) > 0 And Right$(trimmed, 1) = "."
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    Loop
    NormalizeChapterId = trimmed
End Function
