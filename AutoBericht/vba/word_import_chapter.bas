Attribute VB_Name = "modWordImportChapter"
Option Explicit

' Requires JsonConverter.bas (VBA-JSON) in the Word project.

' Config constants live in modAutoBerichtConfig.

Public Sub ImportChapter0Summary()
    LogDebug "ImportChapter0Summary: start"
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
    LogDebug "ImportChapter0Summary: JSON parsed"

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
    LogDebug "ImportChapter0Summary: chapter selected"

    Dim rows As Object
    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then
        MsgBox "No rows in Chapter 0.", vbExclamation
        Exit Sub
    End If

    Dim insertRng As Range
    Set insertRng = ResolveBookmarkInsertRange("Chapter0_start", "Chapter0_end", "")
    If insertRng Is Nothing Then
        Set insertRng = ResolveInsertRange("Chapter0")
    End If
    If insertRng Is Nothing Then
        MsgBox "Bookmark range 'Chapter0_start'/'Chapter0_end' not found.", vbExclamation
        Exit Sub
    End If
    LogDebug "ImportChapter0Summary: insert range " & insertRng.Start & "-" & insertRng.End

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
                    LogDebug "Summary row: " & ResolveDisplayId(row, Nothing)
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
                    lineRange.Style = AB_STYLE_BODY
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

    LogDebug "ImportChapter0Summary: done"
End Sub

' ========= Generalized chapter table =========

Public Sub ImportChapterDialog()
    Dim chapterId As String
    chapterId = PromptChapterId(AB_PROMPT_IMPORT_CHAPTER)
    If Len(chapterId) = 0 Then Exit Sub
    If chapterId = "0" Then
        ImportChapter0Summary
    Else
        ImportChapterTable chapterId, BuildStartBookmark(chapterId), BuildEndBookmark(chapterId)
    End If
End Sub

Public Sub ImportChapterAll()
    Dim ids() As String
    ids = Split(AB_DEFAULT_CHAPTER_IDS, ",")
    Dim i As Long
    ImportChapter0Summary
    For i = LBound(ids) To UBound(ids)
        Dim cid As String
        cid = Trim$(ids(i))
        If cid <> "0" Then
            ImportChapterTable cid, BuildStartBookmark(cid), BuildEndBookmark(cid)
        End If
    Next i
    SaveReportAsFromSidecar
End Sub

Private Function BuildStartBookmark(ByVal chapterId As String) As String
    BuildStartBookmark = Replace$("Chapter" & chapterId & "_start", ".", "_")
End Function

Private Function BuildEndBookmark(ByVal chapterId As String) As String
    BuildEndBookmark = Replace$("Chapter" & chapterId & "_end", ".", "_")
End Function

Private Function PromptChapterId(ByVal prompt As String) As String
    Dim userChoice As String
    userChoice = InputBox(prompt, AB_PROMPT_CHOOSE_CHAPTER_TITLE, AB_PROMPT_CHAPTER_DEFAULT)
    userChoice = Trim$(userChoice)
    If Len(userChoice) = 0 Then Exit Function
    If Not IsValidChapterId(userChoice) Then
        MsgBox "Invalid chapter ID: " & userChoice, vbExclamation
        Exit Function
    End If
    PromptChapterId = userChoice
End Function

Private Function IsValidChapterId(ByVal chapterId As String) As Boolean
    Dim ids() As String
    ids = Split(AB_DEFAULT_CHAPTER_IDS, ",")
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If Trim$(ids(i)) = Trim$(chapterId) Then
            IsValidChapterId = True
            Exit Function
        End If
    Next i
End Function

Private Sub ImportChapterTable(ByVal chapterId As String, ByVal startBm As String, ByVal endBm As String)
    LogDebug "ImportChapterTable: start " & chapterId
    Dim jsonPath As String
    jsonPath = ResolveSidecarPath()
    If Len(jsonPath) = 0 Then Exit Sub

    Dim jsonText As String
    jsonText = ReadAllText(jsonPath)
    If Len(jsonText) = 0 Then
        MsgBox "Sidecar JSON is empty.", vbExclamation
        Exit Sub
    End If
    LogDebug "ImportChapter1Table: loaded JSON (" & Len(jsonText) & " chars)"

    Dim root As Object
    Set root = JsonConverter.ParseJson(jsonText)
    LogDebug "ImportChapter1Table: JSON parsed"

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
    LogDebug "ImportChapter1Table: chapters=" & chapters.Count

    Dim chapter As Object
    Set chapter = FindChapterById(chapters, chapterId)
    If chapter Is Nothing Then
        MsgBox "Chapter id '" & chapterId & "' not found in JSON.", vbExclamation
        Exit Sub
    End If
    chapterId = SafeText(chapter, "id")
    LogDebug "ImportChapterTable: chapterId=" & chapterId

    Dim insertRng As Range
    Set insertRng = ResolveBookmarkInsertRange(startBm, endBm, "")
    If insertRng Is Nothing Then
        Set insertRng = ResolveInsertRange(Replace$("Chapter" & chapterId, ".", "_"))
    End If
    If insertRng Is Nothing Then
        MsgBox "Bookmark range '" & startBm & "' / '" & endBm & "' not found.", vbExclamation
        Exit Sub
    End If
    LogDebug "ImportChapterTable: insert range " & insertRng.Start & "-" & insertRng.End

    Dim rows As Object
    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then
        MsgBox "No rows in Chapter " & chapterId & ".", vbExclamation
        Exit Sub
    End If
    LogDebug "ImportChapterTable: rows=" & rows.Count

    Dim includedSections As Object
    Set includedSections = BuildIncludedSections(rows)

    Dim renumberMap As Object
    Set renumberMap = BuildRenumberMap(rows, chapterId)

    Dim tableRowCount As Long
    tableRowCount = CountDataRows(rows, includedSections)
    If tableRowCount = 0 Then
        MsgBox "No data rows found in Chapter " & chapterId & ".", vbExclamation
        Exit Sub
    End If
    LogDebug "ImportChapterTable: table rows=" & tableRowCount

    insertRng.Collapse wdCollapseStart

    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add(insertRng, tableRowCount + 3, 3)
    tbl.Rows.Alignment = wdAlignRowLeft
    tbl.Rows.AllowBreakAcrossPages = True
    tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    On Error Resume Next
    tbl.Rows.WrapAroundText = True ' place table in text-wrapping (floating) mode instead of inline
    On Error GoTo 0
    LogDebug "ImportChapterTable: table created cols=" & tbl.Columns.Count & " row1cells=" & tbl.Rows(1).Cells.Count
    On Error Resume Next
    tbl.Style = AB_STYLE_TABLE
    On Error GoTo 0
    tbl.Borders.Enable = True
    If tbl.Columns.Count < 3 Then
        MsgBox "Table creation failed (missing columns).", vbExclamation
        Exit Sub
    End If

    ' Remove borders everywhere, then apply only the bottom border of row 1.
    On Error Resume Next
    tbl.Borders.Enable = False
    With tbl.Rows(1).Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth050pt
    End With
    On Error GoTo 0

    ' Header row 1: blank + checkmark (merge later)
    If tbl.Rows(1).Cells.Count < 3 Then
        MsgBox "Table header row 1 has fewer than 3 cells.", vbExclamation
        Exit Sub
    End If
    tbl.Cell(1, 1).Range.Text = ""
    tbl.Cell(1, 2).Range.Text = ""
    tbl.Cell(1, 3).Range.Text = AB_TABLE_HEADER_CHECKMARK
    tbl.Cell(1, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    ' Header row 2: title (merge later)
    If tbl.Rows(2).Cells.Count < 3 Then
        MsgBox "Table header row 2 has fewer than 3 cells.", vbExclamation
        Exit Sub
    End If
    tbl.Cell(2, 1).Range.Text = AB_TABLE_HEADER_TITLE
    tbl.Cell(2, 2).Range.Text = ""
    tbl.Cell(2, 3).Range.Text = ""
    tbl.Cell(2, 1).Range.Font.Bold = True

    ' Header row 3: column labels
    tbl.Cell(3, 1).Range.Text = AB_TABLE_HEADER_COL1
    tbl.Cell(3, 2).Range.Text = AB_TABLE_HEADER_COL2
    tbl.Cell(3, 3).Range.Text = AB_TABLE_HEADER_COL3
    tbl.Rows(3).Range.Font.Bold = True
    tbl.Cell(3, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    ' Column widths after headers populated but before any merges (percent-based)
    On Error Resume Next
    tbl.AllowAutoFit = False
    tbl.PreferredWidthType = wdPreferredWidthPercent
    tbl.PreferredWidth = 100
    tbl.Columns(1).PreferredWidthType = wdPreferredWidthPercent
    tbl.Columns(1).PreferredWidth = AB_TABLE_COL1_WIDTH_PCT
    tbl.Columns(2).PreferredWidthType = wdPreferredWidthPercent
    tbl.Columns(2).PreferredWidth = AB_TABLE_COL2_WIDTH_PCT
    tbl.Columns(3).PreferredWidthType = wdPreferredWidthPercent
    tbl.Columns(3).PreferredWidth = AB_TABLE_COL3_WIDTH_PCT
    On Error GoTo 0

    Dim row As Variant
    Dim targetRow As Long
    targetRow = 4
    Dim sectionRows As Collection
    Set sectionRows = New Collection

    For Each row In rows
        If IsSectionRow(row) Then
            If ShouldIncludeSection(row, includedSections) Then
                LogDebug "Section row: " & SafeSectionTitle(row, renumberMap)
                tbl.Cell(targetRow, 1).Range.Text = SafeSectionTitle(row, renumberMap)
                tbl.Cell(targetRow, 2).Range.Text = ""
                tbl.Cell(targetRow, 3).Range.Text = ""
                sectionRows.Add targetRow
                targetRow = targetRow + 1
            End If
        ElseIf IsIncludedRow(row) Then
            LogDebug "Finding row: " & ResolveDisplayId(row, renumberMap)
            tbl.Cell(targetRow, 1).Range.Text = BuildFindingHeading(row, renumberMap)
            tbl.Cell(targetRow, 1).Range.Style = AB_STYLE_FINDING
            tbl.Cell(targetRow, 2).Range.Text = ResolveRecommendation(row)
            tbl.Cell(targetRow, 2).Range.Style = AB_STYLE_BODY
            tbl.Cell(targetRow, 3).Range.Text = ""
            tbl.Cell(targetRow, 3).Range.Style = AB_STYLE_BODY
            tbl.Cell(targetRow, 3).Range.Font.Bold = True
            tbl.Cell(targetRow, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
            targetRow = targetRow + 1
        End If
    Next row

    ' Merge header rows after content is filled
    On Error Resume Next
    tbl.Cell(1, 1).Merge tbl.Cell(1, 2)
    tbl.Cell(2, 1).Merge tbl.Cell(2, 2)
    On Error GoTo 0
    tbl.Cell(2, 1).Range.Font.Bold = True

    Dim idx As Variant
    For Each idx In sectionRows
        On Error Resume Next
        tbl.Cell(CLng(idx), 1).Merge tbl.Cell(CLng(idx), 3)
        On Error GoTo 0
        On Error Resume Next
        tbl.Cell(CLng(idx), 1).Range.Style = AB_STYLE_SECTION
        On Error GoTo 0
    Next idx

    ' Keep header rows with the first data row to avoid page break after row 3.
    Dim h As Long
    For h = 1 To 3
        On Error Resume Next
        With tbl.Rows(h).Range.ParagraphFormat
            .KeepWithNext = True
            .KeepTogether = True
            .PageBreakBefore = False
        End With
        tbl.Rows(h).AllowBreakAcrossPages = False
        On Error GoTo 0
    Next h

    LogDebug "ImportChapterTable: done"
    ResetTableBookmarks startBm, endBm, tbl
End Sub

Public Sub InsertLogos()
    ' Pick a logo once and place at BIG/SMALL markers.
    Dim logoPath As String
    logoPath = PickLogoFile()
    If Len(logoPath) = 0 Then Exit Sub

    Dim insertedBig As Boolean
    insertedBig = InsertLogoAtMarker(AB_LOGO_MARKER_BIG, logoPath, AB_LOGO_HEIGHT_MAIN_CM, False)

    Dim insertedSmall As Boolean
    insertedSmall = InsertLogoAtMarker(AB_LOGO_MARKER_SMALL, logoPath, AB_LOGO_HEIGHT_HEADER_CM, True)

    If Not insertedBig And Not insertedSmall Then
        MsgBox "Keine Logo-Marker gefunden (" & AB_LOGO_MARKER_BIG & " / " & AB_LOGO_MARKER_SMALL & ").", vbExclamation
    End If
End Sub

Private Function InsertLogoAtMarker(ByVal marker As String, ByVal logoPath As String, ByVal heightCm As Double, ByVal searchHeaders As Boolean) As Boolean
    Dim markerRange As Range
    Set markerRange = FindMarkerRange(marker)
    If markerRange Is Nothing And searchHeaders Then
        Set markerRange = FindMarkerRangeInHeaders(marker)
    End If

    If markerRange Is Nothing Then Exit Function

    markerRange.Text = ""
    Dim inline As InlineShape
    Set inline = markerRange.InlineShapes.AddPicture(FileName:=logoPath, LinkToFile:=False, SaveWithDocument:=True)
    inline.LockAspectRatio = True
    inline.Height = CentimetersToPoints(heightCm)
    InsertLogoAtMarker = True
End Function

Public Sub ImportTextFields()
    ' Replace text field markers from sidecar metadata (NAME$$, COMPANY$$, etc.)
    LogDebug "ImportTextFields: start"

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

    Dim meta As Object
    Set meta = GetObject(project, "meta")
    If meta Is Nothing Then
        MsgBox "Missing report.project.meta in JSON.", vbExclamation
        Exit Sub
    End If

    ' Replace text markers from metadata
    ReplaceTextMarker AB_TEXT_MARKER_NAME, SafeText(meta, "projectName")
    ReplaceTextMarker AB_TEXT_MARKER_COMPANY, SafeText(meta, "company")
    ReplaceTextMarker AB_TEXT_MARKER_COMPANY_ID, SafeText(meta, "companyId")
    ReplaceTextMarker AB_TEXT_MARKER_AUTHOR, SafeText(meta, "author")
    ReplaceTextMarker AB_TEXT_MARKER_MODERATOR, SafeText(meta, "moderator")
    ReplaceTextMarker AB_TEXT_MARKER_CO_MODERATOR, SafeText(meta, "coModerator")

    ' Format date from ISO to DD.MM.YYYY
    Dim dateValue As String
    dateValue = SafeText(meta, "createdAt")
    If Len(dateValue) >= 10 Then
        ' ISO format: YYYY-MM-DD...
        dateValue = Mid$(dateValue, 9, 2) & "." & Mid$(dateValue, 6, 2) & "." & Mid$(dateValue, 1, 4)
    End If
    ReplaceTextMarker AB_TEXT_MARKER_DATE, dateValue

    MsgBox "Text fields replaced.", vbInformation
    LogDebug "ImportTextFields: done"
End Sub

Public Sub InsertSpiderChart()
    ' Insert spider/radar chart at SPIDER$$ marker using sidecar spider data.
    LogDebug "InsertSpiderChart: start"

    Const XL_AXIS_VALUE As Long = 2

    Dim spiderRange As Range
    Set spiderRange = FindMarkerRange(AB_SPIDER_MARKER)
    If spiderRange Is Nothing Then
        MsgBox "Spider marker (" & AB_SPIDER_MARKER & ") not found in document.", vbExclamation
        Exit Sub
    End If

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
    Dim project As Object
    Dim meta As Object
    If Not report Is Nothing Then
        Set project = GetObject(report, "project")
        If Not project Is Nothing Then
            Set meta = GetObject(project, "meta")
        End If
    End If

    Dim spider As Object
    Set spider = GetObject(root, "spider")
    If spider Is Nothing Then
        MsgBox "Spider block not found in sidecar.", vbExclamation
        Exit Sub
    End If

    Dim effective As Object
    Set effective = GetObject(spider, "effective")
    If effective Is Nothing Then
        MsgBox "Spider.effective not found in sidecar.", vbExclamation
        Exit Sub
    End If

    Dim chapters11 As Object, chapters14 As Object, selected As Object
    Set chapters11 = GetObject(effective, "chapters_1_11")
    Set chapters14 = GetObject(effective, "chapters_1_14")

    If Not chapters14 Is Nothing And chapters14.Count > 0 And Not chapters11 Is Nothing And chapters11.Count > 0 Then
        If AB_SPIDER_PROMPT_WHEN_BOTH Then
            Dim choice As VbMsgBoxResult
            choice = MsgBox("Use spider for chapters 1–14? (Yes = 1–14, No = 1–11)", vbYesNoCancel + vbQuestion, "Choose spider range")
            If choice = vbCancel Then Exit Sub
            If choice = vbYes Then
                Set selected = chapters14
            Else
                Set selected = chapters11
            End If
        Else
            If AB_SPIDER_PREFER_14 Then
                Set selected = chapters14
            Else
                Set selected = chapters11
            End If
        End If
    ElseIf Not chapters14 Is Nothing And chapters14.Count > 0 Then
        Set selected = chapters14
    ElseIf Not chapters11 Is Nothing And chapters11.Count > 0 Then
        Set selected = chapters11
    Else
        MsgBox "No spider data (chapters_1_11 / chapters_1_14) found.", vbExclamation
        Exit Sub
    End If

    If Not IsValidSpiderSeries(selected) Then
        MsgBox "Spider data is invalid or empty.", vbExclamation
        Exit Sub
    End If

    ' Clear marker text
    spiderRange.Text = ""

    ' Insert radar chart
    Dim ish As InlineShape
    Set ish = spiderRange.InlineShapes.AddChart(Type:=AB_SPIDER_CHART_TYPE, Range:=spiderRange)
    Dim cht As Object
    Set cht = ish.Chart

    ' Use chart data workbook to avoid category/value swap issues
    cht.ChartData.Activate
    Dim wbData As Object
    Set wbData = cht.ChartData.Workbook
    Dim wsData As Object
    Set wsData = wbData.Worksheets(1)
    wsData.Cells.Clear

    wsData.Cells(1, 1).Value = "Kapitel"
    wsData.Cells(1, 2).Value = IIf(Not meta Is Nothing And SafeText(meta, "company") <> "", "Selbstbeurteilung " & SafeText(meta, "company"), AB_SPIDER_SERIES_COMPANY)
    wsData.Cells(1, 3).Value = AB_SPIDER_SERIES_CONSULTANT

    Dim r As Long: r = 2
    Dim item As Variant
    For Each item In selected
        Dim label As String
        label = SafeText(item, "label")
        If Len(label) = 0 Then label = SafeText(item, "id")
        wsData.Cells(r, 1).Value = label
        wsData.Cells(r, 2).Value = CDbl(Val(SafeText(item, "company")))
        wsData.Cells(r, 3).Value = CDbl(Val(SafeText(item, "consultant")))
        r = r + 1
    Next item

    Dim lastRow As Long
    lastRow = r - 1

    ' Bind series to worksheet ranges (series in columns: A=categories, B/C=values)
    Dim dataRange As Object
    Set dataRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, 3))

    ' Manual series wiring to avoid host PlotBy quirks and ensure both series render
    On Error Resume Next
    Dim sc As Object
    For Each sc In cht.SeriesCollection
        sc.Delete
    Next sc
    On Error GoTo 0

    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(1).Name = wsData.Cells(1, 2).Value
    cht.SeriesCollection(1).Values = wsData.Range(wsData.Cells(2, 2), wsData.Cells(lastRow, 2))
    cht.SeriesCollection(1).XValues = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, 1))

    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(2).Name = wsData.Cells(1, 3).Value
    cht.SeriesCollection(2).Values = wsData.Range(wsData.Cells(2, 3), wsData.Cells(lastRow, 3))
    cht.SeriesCollection(2).XValues = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, 1))

    ' Style chart
    cht.HasTitle = False
    cht.Legend.IncludeInLayout = AB_SPIDER_SHOW_LEGEND
    If AB_SPIDER_SHOW_LEGEND Then
        cht.Legend.Position = AB_SPIDER_LEGEND_POS
    Else
        cht.HasLegend = False
    End If
    On Error Resume Next
    cht.Axes(XL_AXIS_VALUE).MinimumScale = AB_SPIDER_AXIS_MIN
    cht.Axes(XL_AXIS_VALUE).MaximumScale = AB_SPIDER_AXIS_MAX
    On Error GoTo 0

    MsgBox "Spider chart inserted.", vbInformation
    LogDebug "InsertSpiderChart: done"
End Sub

Private Function IsValidSpiderSeries(ByVal seriesData As Object) As Boolean
    On Error GoTo Fail
    If seriesData Is Nothing Then Exit Function
    If seriesData.Count = 0 Then Exit Function
    Dim item As Variant
    For Each item In seriesData
        If Not IsObject(item) Then Exit Function
        If Not item.Exists("id") Then Exit Function
        If Not item.Exists("company") Then Exit Function
        If Not item.Exists("consultant") Then Exit Function
    Next item
    IsValidSpiderSeries = True
    Exit Function
Fail:
    IsValidSpiderSeries = False
End Function

Private Sub ReplaceTextMarker(ByVal marker As String, ByVal value As String)
    Dim rng As Range
    Set rng = FindMarkerRange(marker)
    If Not rng Is Nothing Then
        rng.Text = value
        LogDebug "Replaced " & marker & " with: " & value
    End If
End Sub

Private Function PickLogoFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select logo image"
    fd.Filters.Clear
    fd.Filters.Add "Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Function
    PickLogoFile = fd.SelectedItems(1)
End Function

Private Function ResolveSidecarPath() As String
    Dim defaultPath As String
    If Len(ActiveDocument.Path) > 0 Then
        defaultPath = ActiveDocument.Path & "\\" & AB_SIDECAR_FILENAME
        If FileExists(defaultPath) Then
            ResolveSidecarPath = defaultPath
            Exit Function
        End If
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = AB_SIDECAR_DIALOG_TITLE
    fd.Filters.Clear
    fd.Filters.Add "JSON", "*.json"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Function
    ResolveSidecarPath = fd.SelectedItems(1)
End Function

Private Sub SaveReportAsFromSidecar()
    On Error GoTo CleanFail
    Dim jsonPath As String
    jsonPath = ResolveSidecarPath()
    If Len(jsonPath) = 0 Then Exit Sub

    Dim jsonText As String
    jsonText = ReadAllText(jsonPath)
    If Len(jsonText) = 0 Then Exit Sub

    Dim root As Object
    Set root = JsonConverter.ParseJson(jsonText)
    Dim meta As Object
    Set meta = GetObject(GetObject(GetObject(root, "report"), "project"), "meta")

    Dim company As String
    company = SafeText(meta, "company")
    If Len(company) = 0 Then company = "Unbenannt"
    company = SanitizeFileName(company)

    Dim folder As String
    If Len(ActiveDocument.Path) > 0 Then
        folder = ActiveDocument.Path
    Else
        folder = GetFolderFromPath(jsonPath)
    End If
    If Len(folder) = 0 Then Exit Sub

    Dim fileName As String
    fileName = Format(Date, "yyyy-mm-dd") & " " & company & " Ist Aufnahme Bericht.docx"
    Dim fullPath As String
    fullPath = folder & Application.PathSeparator & fileName

    ActiveDocument.SaveAs2 FileName:=fullPath, FileFormat:=AB_WD_FORMAT_DOCX, AddToRecentFiles:=False
    Exit Sub
CleanFail:
    ' silent fail; avoid blocking imports
End Sub

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

Private Function GetFolderFromPath(ByVal path As String) As String
    Dim pos As Long
    pos = InStrRev(path, Application.PathSeparator)
    If pos > 0 Then
        GetFolderFromPath = Left$(path, pos - 1)
    End If
End Function

Private Function SanitizeFileName(ByVal value As String) As String
    Dim badChars As Variant
    badChars = Array("\\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    SanitizeFileName = value
    For i = LBound(badChars) To UBound(badChars)
        SanitizeFileName = Replace$(SanitizeFileName, badChars(i), "_")
    Next i
    SanitizeFileName = Trim$(SanitizeFileName)
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

Private Function ResolveBookmarkInsertRange(ByVal startName As String, ByVal endName As String, ByVal markerText As String) As Range
    Dim bmStart As Bookmark
    Dim bmEnd As Bookmark
    Set bmStart = FindBookmark(startName)
    Set bmEnd = FindBookmark(endName)
    If bmStart Is Nothing Or bmEnd Is Nothing Then Exit Function

    ' Clear existing content between bookmarks (including prior tables)
    Dim clearRange As Range
    Set clearRange = ActiveDocument.Range(bmStart.Range.End, bmEnd.Range.Start)
    ClearRangeSafe clearRange

    ' Create a clean insertion point just after the start bookmark
    Dim insertRange As Range
    Set insertRange = ActiveDocument.Range(bmStart.Range.End, bmStart.Range.End)
    insertRange.Collapse wdCollapseEnd
    insertRange.InsertParagraphAfter
    insertRange.Collapse wdCollapseEnd

    ' If we're inside a table, move out to the next paragraph to avoid nested tables
    If insertRange.Information(wdWithInTable) Then
        insertRange = insertRange.Paragraphs(1).Range
        insertRange.Collapse wdCollapseEnd
        insertRange.InsertParagraphAfter
        insertRange.Collapse wdCollapseEnd
    End If

    Set ResolveBookmarkInsertRange = insertRange
End Function

Private Function FindMarkerRange(ByVal markerText As String) As Range
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = markerText
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        If .Execute Then
            Set FindMarkerRange = rng
        End If
    End With
End Function

Private Function FindMarkerRangeInHeaders(ByVal markerText As String) As Range
    Dim sec As Section
    Dim hdr As HeaderFooter
    For Each sec In ActiveDocument.Sections
        For Each hdr In sec.Headers
            Dim rng As Range
            Set rng = hdr.Range.Duplicate
            With rng.Find
                .Text = markerText
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
                If .Execute Then
                    Set FindMarkerRangeInHeaders = rng
                    Exit Function
                End If
            End With
        Next hdr
    Next sec
End Function

Private Sub ResetTableBookmarks(ByVal startName As String, ByVal endName As String, ByVal tbl As Table)
    On Error Resume Next
    If ActiveDocument.Bookmarks.Exists(startName) Then ActiveDocument.Bookmarks(startName).Delete
    If ActiveDocument.Bookmarks.Exists(endName) Then ActiveDocument.Bookmarks(endName).Delete
    Dim startRng As Range
    Dim endRng As Range
    Set startRng = tbl.Range.Duplicate
    startRng.End = startRng.Start
    Set endRng = tbl.Range.Duplicate
    endRng.Start = endRng.End
    ActiveDocument.Bookmarks.Add startName, startRng
    ActiveDocument.Bookmarks.Add endName, endRng
    On Error GoTo 0
End Sub

Private Sub LogDebug(ByVal message As String)
    If Not AB_DEBUG_WORD_IMPORT Then Exit Sub
    Debug.Print Format$(Now, "hh:nn:ss") & " | " & message
End Sub

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
    End If
    Dim cleaned As String
    cleaned = StripLeadingNumber(title)
    If Len(cleaned) = 0 Then
        SafeSectionTitle = title
    Else
        SafeSectionTitle = cleaned
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

Private Function StripLeadingNumber(ByVal value As String) As String
    Dim trimmed As String
    trimmed = LTrim$(value)
    Dim i As Long
    i = 1
    Do While i <= Len(trimmed)
        Dim ch As String
        ch = Mid$(trimmed, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            i = i + 1
        ElseIf ch = " " Or ch = "-" Or ch = ":" Then
            i = i + 1
            Exit Do
        Else
            Exit Do
        End If
    Loop
    StripLeadingNumber = LTrim$(Mid$(trimmed, i))
End Function

Private Function IsFieldObservationChapter(ByVal chapterId As String) As Boolean
    If InStr(chapterId, ".") > 0 Then
        IsFieldObservationChapter = True
    End If
End Function

Private Function BuildFindingHeading(ByVal row As Object, ByVal renumberMap As Object) As String
    BuildFindingHeading = ResolveFinding(row)
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
