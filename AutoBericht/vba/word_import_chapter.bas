Attribute VB_Name = "modWordImportChapter"
Option Explicit

' Requires JsonConverter.bas (VBA-JSON) in the Word project.

' === STYLE CONFIG (edit these to match your template) ===
Private Const STYLE_BODY As String = "Normal"
Private Const STYLE_SECTION As String = "Heading 2"
Private Const STYLE_FINDING As String = "Heading 3"
Private Const STYLE_TABLE As String = "Grid Table Light"
Private Const STYLE_LIST As String = "List Paragraph"

' === TABLE CONFIG (edit widths as needed) ===
Private Const COL1_WIDTH_CM As Double = 6.5
Private Const COL2_WIDTH_CM As Double = 9.5
Private Const COL3_WIDTH_CM As Double = 1.3
Private Const HEADER_CHECKMARK As String = "✓"

Public Sub ImportChapter1Table()
    Dim jsonPath As String
    jsonPath = ResolveSidecarPath()
    If Len(jsonPath) = 0 Then Exit Sub

    Dim jsonText As String
    jsonText = ReadAllText(jsonPath)
    If Len(jsonText) = 0 Then
        MsgBox "Sidecar JSON is empty.", vbExclamation
        Exit Sub
    End If

    Dim root As Object
    Set root = JsonConverter.ParseJson(jsonText)

    Dim report As Object
    Set report = GetObject(root, "report")
    If report Is Nothing Then
        MsgBox "Missing report section in JSON.", vbExclamation
        Exit Sub
    End If

    Dim project As Object
    Set project = GetObject(report, "project")
    If project Is Nothing Then
        MsgBox "Missing report.project in JSON.", vbExclamation
        Exit Sub
    End If

    Dim chapters As Object
    Set chapters = GetObject(project, "chapters")
    If chapters Is Nothing Or chapters.Count = 0 Then
        MsgBox "No chapters found in JSON.", vbExclamation
        Exit Sub
    End If

    Dim chapter As Object
    Set chapter = chapters.Item(1)
    Dim chapterId As String
    chapterId = SafeText(chapter, "id")

    Dim insertRng As Range
    Set insertRng = ResolveBookmarkRange("Chapter1_start", "Chapter1_end")
    If insertRng Is Nothing Then
        Set insertRng = ResolveInsertRange("Chapter1")
    End If
    If insertRng Is Nothing Then
        MsgBox "Bookmark range 'Chapter1_start'/'Chapter1_end' not found.", vbExclamation
        Exit Sub
    End If

    Dim rows As Object
    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then
        MsgBox "No rows in Chapter 1.", vbExclamation
        Exit Sub
    End If

    Dim includedSections As Object
    Set includedSections = BuildIncludedSections(rows)

    Dim renumberMap As Object
    Set renumberMap = BuildRenumberMap(rows, chapterId)

    Dim tableRowCount As Long
    tableRowCount = CountDataRows(rows, includedSections)
    If tableRowCount = 0 Then
        MsgBox "No data rows found in Chapter 1.", vbExclamation
        Exit Sub
    End If

    ClearRangeSafe insertRng
    insertRng.Collapse wdCollapseStart

    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add(insertRng, tableRowCount + 3, 3)
    On Error Resume Next
    tbl.Style = STYLE_TABLE
    On Error GoTo 0
    tbl.Borders.Enable = True

    ' Header row 1: blank + checkmark
    On Error Resume Next
    tbl.Cell(1, 1).Merge tbl.Cell(1, 2)
    On Error GoTo 0
    tbl.Cell(1, 1).Range.Text = ""
    tbl.Cell(1, 3).Range.Text = HEADER_CHECKMARK
    tbl.Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    ' Header row 2: title
    On Error Resume Next
    tbl.Cell(2, 1).Merge tbl.Cell(2, 2)
    On Error GoTo 0
    tbl.Cell(2, 1).Range.Text = "Systempunkte mit Verbesserungspotenzial"
    tbl.Cell(2, 1).Range.Font.Bold = True

    ' Header row 3: column labels
    tbl.Cell(3, 1).Range.Text = "Ist-Zustand"
    tbl.Cell(3, 2).Range.Text = "Lösungsansätze"
    tbl.Cell(3, 3).Range.Text = "Prio"
    tbl.Rows(3).Range.Font.Bold = True
    tbl.Cell(3, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    Dim row As Variant
    Dim targetRow As Long
    targetRow = 4

    For Each row In rows
        If IsSectionRow(row) Then
            If ShouldIncludeSection(row, includedSections) Then
                On Error Resume Next
                tbl.Cell(targetRow, 1).Merge tbl.Cell(targetRow, 2)
                On Error GoTo 0
                tbl.Cell(targetRow, 1).Range.Text = SafeSectionTitle(row, renumberMap)
                tbl.Rows(targetRow).Range.Style = STYLE_SECTION
                targetRow = targetRow + 1
            End If
        ElseIf IsIncludedRow(row) Then
            tbl.Cell(targetRow, 1).Range.Text = BuildFindingHeading(row, renumberMap)
            tbl.Cell(targetRow, 1).Range.Style = STYLE_FINDING
            tbl.Cell(targetRow, 2).Range.Text = ResolveRecommendation(row)
            tbl.Cell(targetRow, 2).Range.Style = STYLE_BODY
            tbl.Cell(targetRow, 3).Range.Text = ""
            tbl.Cell(targetRow, 3).Range.Style = STYLE_BODY
            tbl.Cell(targetRow, 3).Range.Font.Bold = True
            tbl.Cell(targetRow, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            targetRow = targetRow + 1
        End If
    Next row

    On Error Resume Next
    tbl.Columns(1).PreferredWidthType = wdPreferredWidthPoints
    tbl.Columns(1).PreferredWidth = CentimetersToPoints(COL1_WIDTH_CM)
    tbl.Columns(2).PreferredWidthType = wdPreferredWidthPoints
    tbl.Columns(2).PreferredWidth = CentimetersToPoints(COL2_WIDTH_CM)
    tbl.Columns(3).PreferredWidthType = wdPreferredWidthPoints
    tbl.Columns(3).PreferredWidth = CentimetersToPoints(COL3_WIDTH_CM)
    On Error GoTo 0
    tbl.AutoFitBehavior wdAutoFitFixed

    Dim i As Long
    For i = 1 To tbl.Rows.Count
        On Error Resume Next
        With tbl.Cell(i, 3).Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
        End With
        On Error GoTo 0
    Next i

    Dim colCell As Cell
    For Each colCell In tbl.Columns(3).Cells
        On Error Resume Next
        colCell.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        On Error GoTo 0
    Next colCell

    MsgBox "Chapter 1 table imported.", vbInformation
End Sub

Public Sub ImportChapter0Summary()
    Dim jsonPath As String
    jsonPath = ResolveSidecarPath()
    If Len(jsonPath) = 0 Then Exit Sub

    Dim jsonText As String
    jsonText = ReadAllText(jsonPath)
    If Len(jsonText) = 0 Then
        MsgBox "Sidecar JSON is empty.", vbExclamation
        Exit Sub
    End If

    Dim root As Object
    Set root = JsonConverter.ParseJson(jsonText)

    Dim report As Object
    Set report = GetObject(root, "report")
    If report Is Nothing Then
        MsgBox "Missing report section in JSON.", vbExclamation
        Exit Sub
    End If

    Dim project As Object
    Set project = GetObject(report, "project")
    If project Is Nothing Then
        MsgBox "Missing report.project in JSON.", vbExclamation
        Exit Sub
    End If

    Dim chapters As Object
    Set chapters = GetObject(project, "chapters")
    If chapters Is Nothing Or chapters.Count = 0 Then
        MsgBox "No chapters found in JSON.", vbExclamation
        Exit Sub
    End If

    Dim chapter As Object
    Set chapter = FindChapterById(chapters, "0")
    If chapter Is Nothing Then
        Set chapter = chapters.Item(1)
    End If

    Dim rows As Object
    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then
        MsgBox "No rows in Chapter 0.", vbExclamation
        Exit Sub
    End If

    Dim insertRng As Range
    Set insertRng = ResolveBookmarkRange("Chapter0_start", "Chapter0_end")
    If insertRng Is Nothing Then
        Set insertRng = ResolveInsertRange("Chapter0")
    End If
    If insertRng Is Nothing Then
        MsgBox "Bookmark range 'Chapter0_start'/'Chapter0_end' not found.", vbExclamation
        Exit Sub
    End If

    ClearRangeSafe insertRng
    insertRng.Collapse wdCollapseStart

    Dim writer As Range
    Set writer = insertRng.Duplicate
    writer.Collapse wdCollapseStart
    Dim startPos As Long
    startPos = writer.Start

    Dim wrote As Boolean
    Dim row As Variant
    For Each row In rows
        If Not IsSectionRow(row) Then
            If IsIncludedRow(row) Then
                Dim summaryText As String
                summaryText = ResolveRecommendation(row)
                If Len(Trim$(summaryText)) > 0 Then
                    summaryText = Replace(summaryText, vbCrLf, " ")
                    summaryText = Replace(summaryText, vbCr, " ")
                    summaryText = Replace(summaryText, vbLf, " ")
                    If wrote Then
                        writer.InsertParagraphAfter
                        writer.Collapse wdCollapseEnd
                    End If
                    Dim lineRange As Range
                    Set lineRange = writer.Duplicate
                    lineRange.Text = summaryText
                    On Error Resume Next
                    lineRange.Style = STYLE_BODY
                    On Error GoTo 0
                    writer.SetRange lineRange.End, lineRange.End
                    wrote = True
                End If
            End If
        End If
    Next row

    If Not wrote Then
        MsgBox "No summary rows to import for Chapter 0.", vbExclamation
        Exit Sub
    End If

    Dim listRange As Range
    Set listRange = ActiveDocument.Range(startPos, writer.End)
    On Error Resume Next
    listRange.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, _
        ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
    listRange.ListFormat.ListTemplate.ListLevels(1).NumberStyle = wdListNumberStyleUppercaseLetter
    On Error GoTo 0

    MsgBox "Chapter 0 summary imported.", vbInformation
End Sub

Private Function ResolveSidecarPath() As String
    Dim defaultPath As String
    If Len(ActiveDocument.Path) > 0 Then
        defaultPath = ActiveDocument.Path & "\\project_sidecar.json"
        If FileExists(defaultPath) Then
            ResolveSidecarPath = defaultPath
            Exit Function
        End If
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select project_sidecar.json"
    fd.Filters.Clear
    fd.Filters.Add "JSON", "*.json"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Function
    ResolveSidecarPath = fd.SelectedItems(1)
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim text As String
    text = ReadAllTextUtf8(path)
    If Len(text) > 0 Then
        ReadAllText = text
        Exit Function
    End If
    Dim f As Integer
    f = FreeFile
    On Error GoTo CleanFail
    Open path For Input As #f
    ReadAllText = Input$(LOF(f), f)
CleanExit:
    Close #f
    Exit Function
CleanFail:
    ReadAllText = ""
    Resume CleanExit
End Function

Private Function ReadAllTextUtf8(ByVal path As String) As String
    On Error GoTo CleanFail
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadAllTextUtf8 = stream.ReadText(-1)
    stream.Close
    Exit Function
CleanFail:
    ReadAllTextUtf8 = ""
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(path) <> "")
End Function

Private Function FindContentControl(ByVal title As String) As ContentControl
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If LCase$(cc.Title) = LCase$(title) Or LCase$(cc.Tag) = LCase$(title) Then
            Set FindContentControl = cc
            Exit Function
        End If
    Next cc
End Function

Private Function FindBookmark(ByVal name As String) As Bookmark
    On Error Resume Next
    Set FindBookmark = ActiveDocument.Bookmarks(name)
    On Error GoTo 0
End Function

Private Function ResolveInsertRange(ByVal anchorName As String) As Range
    Dim cc As ContentControl
    Set cc = FindContentControl(anchorName)
    If Not cc Is Nothing Then
        Set ResolveInsertRange = cc.Range
        Exit Function
    End If

    Dim bm As Bookmark
    Set bm = FindBookmark(anchorName)
    If bm Is Nothing Then Exit Function

    Dim anchorPara As Paragraph
    Set anchorPara = bm.Range.Paragraphs(1)

    If anchorPara.Next Is Nothing Then
        anchorPara.Range.InsertParagraphAfter
    End If

    Dim blankPara As Paragraph
    Set blankPara = anchorPara.Next
    If Len(Trim$(Replace(blankPara.Range.Text, vbCr, ""))) > 0 Then
        blankPara.Range.InsertParagraphBefore
        Set blankPara = anchorPara.Next
    End If

    If blankPara.Next Is Nothing Then
        blankPara.Range.InsertParagraphAfter
    End If

    Dim targetPara As Paragraph
    Set targetPara = blankPara.Next
    Set ResolveInsertRange = targetPara.Range
End Function

Private Function ResolveBookmarkRange(ByVal startName As String, ByVal endName As String) As Range
    Dim bmStart As Bookmark
    Dim bmEnd As Bookmark
    Set bmStart = FindBookmark(startName)
    Set bmEnd = FindBookmark(endName)
    If bmStart Is Nothing Or bmEnd Is Nothing Then Exit Function
    Set ResolveBookmarkRange = ActiveDocument.Range(bmStart.Range.End, bmEnd.Range.Start)
End Function

Private Sub ClearRangeSafe(ByVal rng As Range)
    If rng Is Nothing Then Exit Sub
    If rng.Start >= rng.End Then Exit Sub
    On Error Resume Next
    rng.Delete
    If Err.Number <> 0 Then
        Err.Clear
        rng.Text = ""
    End If
    On Error GoTo 0
End Sub

Private Function FindChapterById(ByVal chapters As Object, ByVal chapterId As String) As Object
    Dim chapter As Variant
    For Each chapter In chapters
        Dim cid As String
        cid = SafeText(chapter, "id")
        If cid = chapterId Then
            Set FindChapterById = chapter
            Exit Function
        End If
    Next chapter
End Function

Private Function IsSectionRow(ByVal row As Object) As Boolean
    On Error GoTo SafeExit
    If row.Exists("kind") Then
        IsSectionRow = (LCase$(CStr(row("kind"))) = "section")
    End If
SafeExit:
End Function

Private Function CountDataRows(ByVal rows As Object, ByVal includedSections As Object) As Long
    Dim row As Variant
    Dim count As Long
    For Each row In rows
        If IsSectionRow(row) Then
            If ShouldIncludeSection(row, includedSections) Then count = count + 1
        ElseIf IsIncludedRow(row) Then
            count = count + 1
        End If
    Next row
    CountDataRows = count
End Function

Private Function BuildIncludedSections(ByVal rows As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim row As Variant
    For Each row In rows
        If Not IsSectionRow(row) Then
            If IsIncludedRow(row) Then
                Dim sectionKey As String
                sectionKey = ""
                On Error Resume Next
                If row.Exists("sectionId") Then sectionKey = CStr(row("sectionId"))
                On Error GoTo 0
                If Len(sectionKey) > 0 Then dict(sectionKey) = True
            End If
        End If
    Next row
    Set BuildIncludedSections = dict
End Function

Private Function ShouldIncludeSection(ByVal row As Object, ByVal includedSections As Object) As Boolean
    Dim key As String
    key = ""
    On Error Resume Next
    If row.Exists("id") Then key = CStr(row("id"))
    On Error GoTo 0
    If Len(key) = 0 Then Exit Function
    If includedSections.Exists(key) Then ShouldIncludeSection = True
End Function

Private Function SafeSectionTitle(ByVal row As Object, ByVal renumberMap As Object) As String
    Dim title As String
    title = ""
    On Error Resume Next
    If row.Exists("title") Then title = CStr(row("title"))
    On Error GoTo 0
    If Len(title) = 0 Then
        title = SafeText(row, "id")
    Else
        title = title
    End If
    Dim sectionId As String
    sectionId = ResolveSectionId(row)
    Dim displayId As String
    displayId = ResolveSectionDisplayId(sectionId, renumberMap)
    If Len(displayId) > 0 Then
        SafeSectionTitle = displayId & " " & title
    Else
        SafeSectionTitle = title
    End If
End Function

Private Function IsIncludedRow(ByVal row As Object) As Boolean
    Dim ws As Object
    Set ws = GetObject(row, "workstate")
    If ws Is Nothing Then
        IsIncludedRow = True
        Exit Function
    End If
    On Error Resume Next
    If ws.Exists("includeFinding") Then
        IsIncludedRow = CBool(ws("includeFinding"))
    Else
        IsIncludedRow = True
    End If
    On Error GoTo 0
End Function

Private Function BuildRenumberMap(ByVal rows As Object, ByVal chapterId As String) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    Dim sectionMap As Object
    Set sectionMap = CreateObject("Scripting.Dictionary")
    Dim sectionCounts As Object
    Set sectionCounts = CreateObject("Scripting.Dictionary")
    Dim itemCount As Long

    Dim row As Variant
    For Each row In rows
        If IsSectionRow(row) Then
            ' section rows handled via sectionMap
        ElseIf IsIncludedRow(row) Then
            Dim rowId As String
            rowId = SafeText(row, "id")
            If Len(rowId) = 0 Then GoTo ContinueLoop
            If IsFieldObservationChapter(chapterId) Then
                itemCount = itemCount + 1
                map(rowId) = chapterId & "." & CStr(itemCount)
            Else
                Dim sectionKey As String
                sectionKey = ResolveSectionId(row)
                If Len(sectionKey) = 0 Then sectionKey = chapterId & ".1"
                If Not sectionMap.Exists(sectionKey) Then
                    sectionMap(sectionKey) = sectionMap.Count + 1
                End If
                Dim count As Long
                If sectionCounts.Exists(sectionKey) Then
                    count = CLng(sectionCounts(sectionKey)) + 1
                Else
                    count = 1
                End If
                sectionCounts(sectionKey) = count
                map(rowId) = chapterId & "." & CStr(sectionMap(sectionKey)) & "." & CStr(count)
            End If
        End If
ContinueLoop:
    Next row

    Set map("_sectionMap") = sectionMap
    Set BuildRenumberMap = map
End Function

Private Function ResolveDisplayId(ByVal row As Object, ByVal renumberMap As Object) As String
    Dim rowId As String
    rowId = SafeText(row, "id")
    If Len(rowId) = 0 Then Exit Function
    On Error Resume Next
    If Not renumberMap Is Nothing Then
        If renumberMap.Exists(rowId) Then ResolveDisplayId = CStr(renumberMap(rowId))
    End If
    On Error GoTo 0
    If Len(ResolveDisplayId) = 0 Then ResolveDisplayId = rowId
End Function

Private Function ResolveSectionId(ByVal row As Object) As String
    On Error Resume Next
    If row.Exists("sectionId") Then
        ResolveSectionId = CStr(row("sectionId"))
        Exit Function
    End If
    On Error GoTo 0
    Dim rowId As String
    rowId = SafeText(row, "id")
    Dim parts() As String
    parts = Split(rowId, ".")
    If UBound(parts) >= 1 Then
        ResolveSectionId = parts(0) & "." & parts(1)
    End If
End Function

Private Function ResolveSectionDisplayId(ByVal sectionId As String, ByVal renumberMap As Object) As String
    If Len(sectionId) = 0 Then Exit Function
    On Error Resume Next
    If Not renumberMap Is Nothing Then
        Dim sectionMap As Object
        If renumberMap.Exists("_sectionMap") Then
            Set sectionMap = renumberMap("_sectionMap")
            If Not sectionMap Is Nothing Then
                If sectionMap.Exists(sectionId) Then
                    Dim chapterPart As String
                    chapterPart = Split(sectionId, ".")(0)
                    ResolveSectionDisplayId = chapterPart & "." & CStr(sectionMap(sectionId))
                End If
            End If
        End If
    End If
    On Error GoTo 0
    If Len(ResolveSectionDisplayId) = 0 Then ResolveSectionDisplayId = sectionId
End Function

Private Function IsFieldObservationChapter(ByVal chapterId As String) As Boolean
    If InStr(chapterId, ".") > 0 Then
        IsFieldObservationChapter = True
    End If
End Function

Private Function BuildFindingHeading(ByVal row As Object, ByVal renumberMap As Object) As String
    Dim displayId As String
    displayId = ResolveDisplayId(row, renumberMap)
    Dim finding As String
    finding = ResolveFinding(row)
    If Len(displayId) > 0 And Len(finding) > 0 Then
        BuildFindingHeading = displayId & " " & finding
    ElseIf Len(finding) > 0 Then
        BuildFindingHeading = finding
    Else
        BuildFindingHeading = displayId
    End If
End Function

Private Function ResolveFinding(ByVal row As Object) As String
    Dim ws As Object
    Set ws = GetObject(row, "workstate")
    If Not ws Is Nothing Then
        If GetBool(ws, "useFindingOverride") Then
            ResolveFinding = SafeText(ws, "findingOverride")
            Exit Function
        End If
    End If

    Dim master As Object
    Set master = GetObject(row, "master")
    If Not master Is Nothing Then
        ResolveFinding = ToPlainText(master("finding"))
    End If
End Function

Private Function ResolveRecommendation(ByVal row As Object) As String
    Dim levelKey As String
    levelKey = "1"

    Dim ws As Object
    Set ws = GetObject(row, "workstate")
    If Not ws Is Nothing Then
        If ws.Exists("includeRecommendation") Then
            If Not CBool(ws("includeRecommendation")) Then
                ResolveRecommendation = ""
                Exit Function
            End If
        End If
        If ws.Exists("selectedLevel") Then
            levelKey = CStr(ws("selectedLevel"))
        End If
        If ws.Exists("useLevelOverride") Then
            Dim overrides As Object
            Set overrides = GetObject(ws, "levelOverrides")
            If Not overrides Is Nothing Then
                If GetBoolFromDict(ws("useLevelOverride"), levelKey) Then
                    ResolveRecommendation = ToPlainText(overrides(levelKey))
                    Exit Function
                End If
            End If
        End If
    End If

    Dim master As Object
    Set master = GetObject(row, "master")
    If Not master Is Nothing Then
        Dim levels As Object
        Set levels = GetObject(master, "levels")
        If Not levels Is Nothing Then
            If levels.Exists(levelKey) Then
                ResolveRecommendation = ToPlainText(levels(levelKey))
                Exit Function
            End If
        End If
    End If
End Function

Private Function GetObject(ByVal dict As Object, ByVal key As String) As Object
    On Error GoTo SafeExit
    If dict Is Nothing Then Exit Function
    If dict.Exists(key) Then
        If IsObject(dict(key)) Then
            Set GetObject = dict(key)
        End If
    End If
SafeExit:
End Function

Private Function SafeText(ByVal dict As Object, ByVal key As String) As String
    On Error GoTo SafeExit
    If dict.Exists(key) Then
        SafeText = ToPlainText(dict(key))
    End If
SafeExit:
End Function

Private Function ToPlainText(ByVal value As Variant) As String
    If IsObject(value) Then
        If TypeName(value) = "Collection" Then
            Dim parts As Collection
            Set parts = value
            Dim item As Variant
            Dim buff As String
            For Each item In parts
                If Len(buff) > 0 Then buff = buff & vbCrLf
                buff = buff & CStr(item)
            Next item
            ToPlainText = buff
            Exit Function
        End If
    End If
    If IsNull(value) Then
        ToPlainText = ""
    Else
        ToPlainText = CStr(value)
    End If
End Function

Private Function GetBool(ByVal dict As Object, ByVal key As String) As Boolean
    On Error GoTo SafeExit
    If dict.Exists(key) Then
        GetBool = CBool(dict(key))
    End If
SafeExit:
End Function

Private Function GetBoolFromDict(ByVal dict As Variant, ByVal key As String) As Boolean
    On Error GoTo SafeExit
    If IsObject(dict) Then
        If dict.Exists(key) Then
            GetBoolFromDict = CBool(dict(key))
        End If
    End If
SafeExit:
End Function
