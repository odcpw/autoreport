Attribute VB_Name = "modPptExport"
Option Explicit

' Config constants live in modAutoBerichtConfig.

Public Sub ExportTrainingPptD()
    ExportTrainingPptInternal "D"
End Sub

Public Sub ExportTrainingPptF()
    ExportTrainingPptInternal "F"
End Sub

Public Sub ExportBesprechungPpt()
    Dim sidecarPath As String
    Dim sidecarText As String
    Dim root As Object
    Dim report As Object
    Dim project As Object
    Dim chapters As Object
    Dim meta As Object
    Dim photosDoc As Object
    Dim projectFolder As String
    Dim templatesFolder As String
    Dim templatePath As String
    Dim pptApp As Object
    Dim pres As Object
    Dim photoMap As Object
    Dim sectionMapCache As Object
    Dim chapter As Variant
    Dim chapterId As String
    Dim chapterTitle As String
    Dim outPath As String

    ' Export Bericht Besprechung presentation
    LogDebug "ExportBesprechungPpt: start"

    sidecarPath = ResolveSidecarPathPpt()
    If Len(sidecarPath) = 0 Then Exit Sub

    sidecarText = ReadAllTextUtf8(sidecarPath)
    If Len(sidecarText) = 0 Then
        MsgBox "Sidecar JSON is empty.", vbExclamation
        Exit Sub
    End If

    Set root = JsonConverter.ParseJson(sidecarText)

    Set report = GetObject(root, "report")
    If report Is Nothing Then
        MsgBox "Missing report section in sidecar.", vbExclamation
        Exit Sub
    End If

    Set project = GetObject(report, "project")
    If project Is Nothing Then
        MsgBox "Missing report.project in sidecar.", vbExclamation
        Exit Sub
    End If

    Set chapters = GetObject(project, "chapters")
    If chapters Is Nothing Or chapters.Count = 0 Then
        MsgBox "No chapters found in sidecar.", vbExclamation
        Exit Sub
    End If

    Set meta = GetObject(project, "meta")

    Set photosDoc = GetObject(root, "photos")

    projectFolder = GetParentFolder(sidecarPath)
    If Len(AB_PPT_TEMPLATE_FOLDER) > 0 Then
        templatesFolder = projectFolder & "\\" & AB_PPT_TEMPLATE_FOLDER
    Else
        templatesFolder = projectFolder
    End If
    templatePath = templatesFolder & "\\" & AB_PPT_REPORT_TEMPLATE
    If Dir(templatePath) = "" Then
        MsgBox "Report template not found: " & templatePath, vbExclamation
        Exit Sub
    End If

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    Set pres = pptApp.Presentations.Open(templatePath, 0, 0, -1)

    If AB_PPT_REPORT_INCLUDE_TITLE Then
        AddReportTitleSlide pres, meta
    End If

    Set photoMap = BuildReportPhotoMap(photosDoc, projectFolder)

    Set sectionMapCache = CreateObject("Scripting.Dictionary")

    For Each chapter In chapters
        chapterId = SafeText(chapter, "id")
        If Len(chapterId) = 0 Then GoTo ContinueChapter

        chapterTitle = ResolveChapterTitle(chapterId, chapter)

        If chapterId = "0" Then
            AddChapterSeparatorSlide pres, chapterTitle
            AddChapter0SummarySlides pres, chapter, photoMap
        ElseIf chapterId = "4.8" Then
            AddSpecialChapter48Slides pres, chapter, chapters, photoMap, sectionMapCache
        Else
            AddChapterSeparatorSlide pres, BuildNumberedTitle(chapterId, chapterTitle)
            AddChapterScreenshotSlide pres, chapterId, chapterTitle
            AddChapterSectionSlides pres, chapter, chapterId, photoMap, sectionMapCache
        End If

ContinueChapter:
    Next chapter

    outPath = projectFolder & "\\" & Format$(Date, "yyyy-mm-dd") & "_" & AB_PPT_REPORT_OUTPUT_BASE & ".pptx"
    pres.SaveAs outPath
    MsgBox "Besprechung deck exported to: " & outPath, vbInformation

    LogDebug "ExportBesprechungPpt: done"
End Sub

Private Sub ExportTrainingPptInternal(ByVal langSuffix As String)
    Dim sidecarPath As String
    Dim sidecarText As String
    Dim root As Object
    Dim photosDoc As Object
    Dim projectFolder As String
    Dim templatesFolder As String
    Dim templatePath As String
    Dim pptApp As Object
    Dim pres As Object
    Dim seminarLayout As Object
    Dim seminarSlide As Object
    Dim tagMap As Object
    Dim orderedTags As Collection
    Dim tag As Variant
    Dim photos As Collection
    Dim chapterLayout As Object
    Dim chapterSlide As Object
    Dim layoutName As String
    Dim layout As Object
    Dim slots As Long
    Dim idx As Long
    Dim slide As Object
    Dim s As Long
    Dim photoArray() As Variant
    Dim outPath As String

    LogDebug "ExportTrainingPpt: start (" & langSuffix & ")"
    sidecarPath = ResolveSidecarPathPpt()
    If Len(sidecarPath) = 0 Then Exit Sub

    sidecarText = ReadAllTextUtf8(sidecarPath)
    If Len(sidecarText) = 0 Then
        MsgBox "Sidecar JSON is empty.", vbExclamation
        Exit Sub
    End If

    Set root = JsonConverter.ParseJson(sidecarText)
    Set photosDoc = GetObject(root, "photos")
    If photosDoc Is Nothing Then
        MsgBox "No photos section in sidecar.", vbExclamation
        Exit Sub
    End If

    projectFolder = GetParentFolder(sidecarPath)
    If Len(AB_PPT_TEMPLATE_FOLDER) > 0 Then
        templatesFolder = projectFolder & "\\" & AB_PPT_TEMPLATE_FOLDER
    Else
        templatesFolder = projectFolder
    End If
    templatePath = templatesFolder & "\\" & AB_PPT_TEMPLATE
    If Dir(templatePath) = "" Then
        MsgBox "Training template not found: " & templatePath, vbExclamation
        Exit Sub
    End If

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    Set pres = pptApp.Presentations.Open(templatePath, 0, 0, -1)

    If AB_PPT_INCLUDE_SEMINAR_SLIDE Then
        Set seminarLayout = FindLayoutByName(pres, AB_SeminarLayoutName(langSuffix))
        If Not seminarLayout Is Nothing Then
            Set seminarSlide = pres.Slides.AddSlide(pres.Slides.Count + 1, seminarLayout)
            SetTitleIfPresent seminarSlide, AB_PPT_SEMINAR_TITLE
        End If
    End If

    Set tagMap = BuildTrainingTagMap(photosDoc, projectFolder)

    Set orderedTags = BuildTrainingTagOrder(tagMap)

    For Each tag In orderedTags
        Set photos = tagMap(tag)
        If photos.Count = 0 Then GoTo ContinueTag

        Set chapterLayout = FindLayoutByName(pres, AB_PPT_LAYOUT_CHAPTER)
        If Not chapterLayout Is Nothing Then
            Set chapterSlide = pres.Slides.AddSlide(pres.Slides.Count + 1, chapterLayout)
            SetTitleIfPresent chapterSlide, CStr(tag)
        End If

        layoutName = ResolveTrainingLayoutName(CStr(tag), langSuffix)
        Set layout = FindLayoutByName(pres, layoutName)
        If layout Is Nothing Then
            Set layout = FindLayoutByName(pres, AB_PPT_LAYOUT_PICTURE)
        End If

        slots = CountPictureSlots(layout)
        If slots < 1 Then slots = 1

        photoArray = CollectionToArrayDistinct(photos)

        idx = LBound(photoArray)
        Do While idx <= UBound(photoArray)
            Set slide = pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
            SetTitleIfPresent slide, CStr(tag)
            For s = 1 To slots
                If idx > UBound(photoArray) Then Exit For
                InsertPictureSlot slide, CStr(photoArray(idx)), s
                idx = idx + 1
            Next s
        Loop
ContinueTag:
    Next tag

    outPath = projectFolder & "\\" & Format$(Date, "yyyy-mm-dd") & "_" & AB_PPT_OUTPUT_TRAINING_BASE & "_" & langSuffix & ".pptx"
    pres.SaveAs outPath
    MsgBox "Training deck exported to: " & outPath, vbInformation
    LogDebug "ExportTrainingPpt: done"
End Sub

Private Function BuildTrainingTagMap(ByVal photosDoc As Object, ByVal projectFolder As String) As Object
    Dim map As Object
    Dim photos As Object
    Dim key As Variant
    Dim photo As Object
    Dim tags As Object
    Dim training As Object
    Dim absPath As String
    Dim t As Variant
    Dim coll As Collection

    Set map = CreateObject("Scripting.Dictionary")
    Set photos = GetObject(photosDoc, "photos")
    If photos Is Nothing Then
        Set BuildTrainingTagMap = map
        Exit Function
    End If
    For Each key In photos.Keys
        Set photo = photos(key)
        Set tags = GetObject(photo, "tags")
        If tags Is Nothing Then GoTo NextPhoto
        Set training = GetObject(tags, "training")
        If training Is Nothing Then GoTo NextPhoto
        absPath = projectFolder & "\\" & Replace(CStr(key), "/", "\\")
        For Each t In training
            If Not map.Exists(CStr(t)) Then
                Set coll = New Collection
                map.Add CStr(t), coll
            End If
            map(CStr(t)).Add absPath
        Next t
NextPhoto:
    Next key
    Set BuildTrainingTagMap = map
End Function

Private Function BuildTrainingTagOrder(ByVal tagMap As Object) As Collection
    Dim order As Collection
    Dim knownTags As Variant
    Dim i As Long
    Dim key As Variant

    Set order = New Collection
    knownTags = Split(AB_PPT_TAG_ORDER, "|")
    For i = LBound(knownTags) To UBound(knownTags)
        If tagMap.Exists(knownTags(i)) Then order.Add knownTags(i)
    Next i
    For Each key In tagMap.Keys
        If Not ExistsInCollection(order, CStr(key)) Then order.Add CStr(key)
    Next key
    Set BuildTrainingTagOrder = order
End Function

Private Function ResolveTrainingLayoutName(ByVal tag As String, ByVal langSuffix As String) As String
    ResolveTrainingLayoutName = AB_PptTrainingLayoutForTag(tag, langSuffix)
End Function

Private Function FindLayoutByName(ByVal pres As Object, ByVal layoutName As String) As Object
    Dim d As Object
    Dim layout As Object

    For Each d In pres.Designs
        For Each layout In d.SlideMaster.CustomLayouts
            If LCase$(layout.Name) = LCase$(layoutName) Then
                Set FindLayoutByName = layout
                Exit Function
            End If
        Next layout
    Next d
End Function

Private Sub SetTitleIfPresent(ByVal slide As Object, ByVal titleText As String)
    Dim shape As Object
    For Each shape In slide.Shapes
        If shape.HasTextFrame Then
            On Error Resume Next
            If shape.PlaceholderFormat.Type = 1 Then
                shape.TextFrame.TextRange.Text = titleText
                Exit Sub
            End If
            On Error GoTo 0
        End If
    Next shape
End Sub

Private Sub InsertPictureSlot(ByVal slide As Object, ByVal filePath As String, ByVal slotIndex As Long)
    Dim shape As Object
    Dim slot As Long
    slot = 0
    For Each shape In slide.Shapes
        On Error Resume Next
        If shape.Type = AB_MSO_PLACEHOLDER Or shape.Type = AB_PP_PLACEHOLDER_PICTURE Then
            If shape.PlaceholderFormat.Type = AB_PP_PLACEHOLDER_PICTURE Or shape.Type = AB_PP_PLACEHOLDER_PICTURE Then
                slot = slot + 1
                If slot = slotIndex Then
                    InsertPictureIntoBounds slide, shape, filePath
                    Exit Sub
                End If
            End If
        End If
        On Error GoTo 0
    Next shape
End Sub

Private Sub InsertPictureIntoBounds(ByVal slide As Object, ByVal placeholder As Object, ByVal filePath As String)
    Dim l As Single, t As Single, w As Single, h As Single
    Dim pic As Object
    Dim scaleFactor As Double

    l = placeholder.Left: t = placeholder.Top: w = placeholder.Width: h = placeholder.Height
    Set pic = slide.Shapes.AddPicture(filePath, msoFalse, msoTrue, l, t, w, h)
    On Error Resume Next
    pic.LockAspectRatio = msoTrue
    If pic.Height = 0 Or pic.Width = 0 Then GoTo Done
    scaleFactor = w / pic.Width
    If (pic.Height * scaleFactor) > h Then scaleFactor = h / pic.Height
    pic.Width = pic.Width * scaleFactor
    pic.Height = pic.Height * scaleFactor
    pic.Left = l + (w - pic.Width) / 2
    pic.Top = t + (h - pic.Height) / 2
Done:
    placeholder.Delete
    On Error GoTo 0
End Sub

Private Function CountPictureSlots(ByVal layout As Object) As Long
    Dim shape As Object
    Dim count As Long
    For Each shape In layout.Shapes
        On Error Resume Next
        If shape.Type = AB_MSO_PLACEHOLDER Or shape.Type = AB_PP_PLACEHOLDER_PICTURE Then
            If shape.PlaceholderFormat.Type = AB_PP_PLACEHOLDER_PICTURE Or shape.Type = AB_PP_PLACEHOLDER_PICTURE Then
                count = count + 1
            End If
        End If
        On Error GoTo 0
    Next shape
    CountPictureSlots = count
End Function

Private Function CollectionToArray(ByVal coll As Collection) As Variant()
    Dim arr() As Variant
    Dim i As Long

    ReDim arr(1 To coll.Count)
    For i = 1 To coll.Count
        arr(i) = coll(i)
    Next i
    CollectionToArray = arr
End Function

Private Function CollectionToArrayDistinct(ByVal coll As Collection) As Variant()
    ' Deduplicate collection items using Dictionary (case-insensitive)
    Dim dict As Object
    Dim item As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim key As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For Each item In coll
        If Not dict.Exists(CStr(item)) Then
            dict.Add CStr(item), True
        End If
    Next item

    ' Convert unique keys to array (1-based for VBA consistency)
    If dict.Count = 0 Then
        ReDim arr(1 To 1)
        arr(1) = ""
    Else
        ReDim arr(1 To dict.Count)
        i = 1
        For Each key In dict.Keys
            arr(i) = key
            i = i + 1
        Next key
    End If

    CollectionToArrayDistinct = arr
End Function

Private Function ExistsInCollection(ByVal coll As Collection, ByVal value As String) As Boolean
    Dim item As Variant
    For Each item In coll
        If LCase$(CStr(item)) = LCase$(value) Then
            ExistsInCollection = True
            Exit Function
        End If
    Next item
End Function

' === Bericht Besprechung helpers ===

Private Sub AddReportTitleSlide(ByVal pres As Object, ByVal meta As Object)
    Dim slide As Object
    Dim titleText As String

    Set slide = AddReportSlide(pres, AB_PPT_REPORT_LAYOUT_TITLE)
    If slide Is Nothing Then Exit Sub

    titleText = SafeText(meta, "projectName")
    If Len(titleText) = 0 Then titleText = SafeText(meta, "company")
    If Len(titleText) = 0 Then titleText = "Report"
    SetTitleIfPresent slide, titleText
End Sub

Private Sub AddChapterSeparatorSlide(ByVal pres As Object, ByVal titleText As String)
    Dim slide As Object
    Set slide = AddReportSlide(pres, AB_PPT_REPORT_LAYOUT_CHAPTER_SEPARATOR)
    If slide Is Nothing Then Exit Sub
    SetTitleIfPresent slide, titleText
End Sub

Private Sub AddChapterScreenshotSlide(ByVal pres As Object, ByVal chapterId As String, ByVal chapterTitle As String)
    Dim slide As Object
    Dim pageRange As Range
    Dim pastedRange As Object
    Dim pastedShape As Object

    Set slide = AddReportSlide(pres, AB_PPT_REPORT_LAYOUT_CHAPTER_SCREENSHOT)
    If slide Is Nothing Then Exit Sub
    SetTitleIfPresent slide, BuildNumberedTitle(chapterId, chapterTitle)

    Set pageRange = GetChapterFirstPageRange(chapterId)
    If pageRange Is Nothing Then Exit Sub

    On Error Resume Next
    pageRange.CopyAsPicture
    If Err.Number <> 0 Then
        Err.Clear
        pageRange.Copy
    End If
    On Error GoTo 0

    On Error Resume Next
    Set pastedRange = slide.Shapes.PasteSpecial(AB_PP_PASTE_ENHANCED_METAFILE)
    If pastedRange Is Nothing Then
        Set pastedRange = slide.Shapes.Paste
    End If
    On Error GoTo 0
    If pastedRange Is Nothing Then Exit Sub
    Set pastedShape = pastedRange(1)
    FitPastedShapeToPlaceholder slide, pastedShape
End Sub

Private Sub AddChapter0SummarySlides(ByVal pres As Object, ByVal chapter As Object, ByVal photoMap As Object)
    Dim rows As Object
    Dim items As Collection
    Dim row As Variant
    Dim lines As Collection

    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then Exit Sub

    Set items = New Collection
    For Each row In rows
        If Not IsSectionRow(row) Then
            If IsIncludedRow(row) Then items.Add row
        End If
    Next row

    Set lines = BuildFindingTextLines(items, Nothing, True, False)
    If lines.Count = 0 Then Exit Sub

    AddFindingSlides pres, AB_PPT_REPORT_CH0_TITLE, lines, AB_PPT_REPORT_LAYOUT_SECTION_TEXT
End Sub

Private Sub AddChapterSectionSlides(ByVal pres As Object, ByVal chapter As Object, ByVal chapterId As String, ByVal photoMap As Object, ByVal sectionMapCache As Object)
    Dim rows As Object
    Dim renumberMap As Object
    Dim sections As Collection
    Dim section As Variant
    Dim sectionId As String
    Dim items As Collection
    Dim photos As Collection
    Dim sectionDisplayId As String
    Dim sectionTitle As String
    Dim lines As Collection

    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then Exit Sub

    Set renumberMap = BuildRenumberMap(rows, chapterId)
    CacheSectionMap sectionMapCache, chapterId, renumberMap

    Set sections = BuildSectionList(rows, chapterId)
    For Each section In sections
        sectionId = CStr(section("id"))

        Set items = section("items")
        Set photos = GetPhotoCollection(photoMap, sectionId)

        If items.Count = 0 And photos.Count = 0 Then GoTo ContinueSection

        sectionDisplayId = ResolveSectionDisplayId(sectionId, renumberMap)
        If Len(sectionDisplayId) = 0 Then sectionDisplayId = sectionId

        sectionTitle = BuildNumberedTitle(sectionDisplayId, CStr(section("title")))

        Set lines = BuildFindingTextLines(items, renumberMap, False, True)
        If lines.Count > 0 Then
            AddFindingSlides pres, sectionTitle, lines, AB_PPT_REPORT_LAYOUT_SECTION_TEXT
        End If

        If photos.Count > 0 Then
            AddPhotoSlides pres, sectionTitle, photos, AB_PPT_REPORT_LAYOUT_SECTION_PHOTO_3, AB_PPT_REPORT_LAYOUT_SECTION_PHOTO_6
        End If

ContinueSection:
    Next section
End Sub

Private Sub AddSpecialChapter48Slides(ByVal pres As Object, ByVal chapter As Object, ByVal chapters As Object, ByVal photoMap As Object, ByVal sectionMapCache As Object)
    Dim displaySectionId As String
    Dim chapterTitle As String
    Dim separatorTitle As String
    Dim separatorSlide As Object
    Dim rows As Object
    Dim orderedRows As Collection
    Dim itemIndex As Long
    Dim row As Variant
    Dim findingText As String
    Dim displayId As String
    Dim observationTag As String
    Dim observationTitle As String
    Dim displayTitle As String
    Dim sectionId As String
    Dim photos As Collection

    displaySectionId = ResolveSpecial48DisplaySectionId(chapters, sectionMapCache)
    If Len(displaySectionId) = 0 Then displaySectionId = "4.8"

    chapterTitle = ResolveChapterTitle("4.8", chapter)
    separatorTitle = BuildNumberedTitle(displaySectionId, chapterTitle)

    Set separatorSlide = AddReportSlide(pres, AB_PPT_REPORT_LAYOUT_48_SEPARATOR)
    If Not separatorSlide Is Nothing Then
        SetTitleIfPresent separatorSlide, separatorTitle
    End If

    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then Exit Sub
    Set orderedRows = OrderObservationRows(rows, chapter)
    If orderedRows Is Nothing Then Exit Sub
    If orderedRows.Count = 0 Then Exit Sub

    For Each row In orderedRows
        If IsSectionRow(row) Then GoTo ContinueRow
        If Not IsIncludedRow(row) Then GoTo ContinueRow

        findingText = ResolveFinding(row)

        itemIndex = itemIndex + 1
        displayId = displaySectionId & "." & CStr(itemIndex)

        observationTag = ResolveObservationTag(row)
        observationTitle = ResolveObservationTitle(row, observationTag)
        displayTitle = BuildNumberedTitle(displayId, observationTitle)
        sectionId = ResolveSectionId(row, "4.8")
        If Len(observationTag) = 0 Then observationTag = sectionId
        Set photos = GetPhotoCollection(photoMap, observationTag)

        AddObservationSlides pres, displayTitle, findingText, photos

ContinueRow:
    Next row
End Sub

Private Function OrderObservationRows(ByVal rows As Object, ByVal chapter As Object) As Collection
    Dim ordered As Collection
    Dim meta As Object
    Dim orderList As Object
    Dim rowIndex As Object
    Dim added As Object
    Dim row As Variant
    Dim rowId As String
    Dim orderId As Variant

    Set ordered = New Collection
    If rows Is Nothing Then
        Set OrderObservationRows = ordered
        Exit Function
    End If

    Set rowIndex = CreateObject("Scripting.Dictionary")
    rowIndex.CompareMode = vbTextCompare
    For Each row In rows
        If Not IsSectionRow(row) Then
            rowId = SafeText(row, "id")
            If Len(rowId) > 0 Then rowIndex(rowId) = row
        End If
    Next row

    Set meta = GetObject(chapter, "meta")
    If Not meta Is Nothing Then
        On Error Resume Next
        Set orderList = meta("order")
        On Error GoTo 0
    End If

    Set added = CreateObject("Scripting.Dictionary")
    added.CompareMode = vbTextCompare
    If Not orderList Is Nothing Then
        For Each orderId In orderList
            rowId = CStr(orderId)
            If rowIndex.Exists(rowId) Then
                ordered.Add rowIndex(rowId)
                added(rowId) = True
            End If
        Next orderId
    End If

    For Each row In rows
        If IsSectionRow(row) Then GoTo ContinueRow
        rowId = SafeText(row, "id")
        If Len(rowId) = 0 Then
            ordered.Add row
        ElseIf Not added.Exists(rowId) Then
            ordered.Add row
            added(rowId) = True
        End If
ContinueRow:
    Next row

    Set OrderObservationRows = ordered
End Function

Private Function ResolveObservationTag(ByVal row As Object) As String
    Dim tagValue As String

    On Error Resume Next
    If row.Exists("tag") Then tagValue = CStr(row("tag"))
    On Error GoTo 0

    If Len(tagValue) = 0 Then tagValue = SafeText(row, "titleOverride")
    If Len(tagValue) = 0 Then tagValue = SafeText(row, "title")
    If Len(tagValue) = 0 Then tagValue = SafeText(row, "id")
    ResolveObservationTag = Trim$(tagValue)
End Function

Private Function ResolveObservationTitle(ByVal row As Object, ByVal fallback As String) As String
    Dim titleValue As String

    titleValue = SafeText(row, "titleOverride")
    If Len(titleValue) = 0 Then titleValue = SafeText(row, "title")
    If Len(titleValue) = 0 Then titleValue = fallback
    ResolveObservationTitle = Trim$(titleValue)
End Function

Private Sub AddObservationSlides(ByVal pres As Object, ByVal titleText As String, ByVal bodyText As String, ByVal photos As Collection)
    Dim layoutName As String
    Dim slots As Long
    Dim photoCount As Long
    Dim slide As Object
    Dim photoArray() As Variant
    Dim idx As Long
    Dim s As Long
    Dim remaining As Collection

    If photos Is Nothing Then
        photoCount = 0
    Else
        photoCount = photos.Count
    End If

    If photoCount > 3 Then
        layoutName = AB_PPT_REPORT_LAYOUT_48_TEXT_PHOTO_6
        slots = 6
    Else
        layoutName = AB_PPT_REPORT_LAYOUT_48_TEXT_PHOTO_3
        slots = 3
    End If

    Set slide = AddReportSlide(pres, layoutName)
    If slide Is Nothing Then Exit Sub
    SetTitleIfPresent slide, titleText
    SetBodyIfPresent slide, bodyText

    If photoCount > 0 Then
        photoArray = CollectionToArrayDistinct(photos)
        idx = LBound(photoArray)
        For s = 1 To slots
            If idx > UBound(photoArray) Then Exit For
            InsertPictureSlot slide, CStr(photoArray(idx)), s
            idx = idx + 1
        Next s

        If idx <= UBound(photoArray) Then
            Set remaining = New Collection
            Do While idx <= UBound(photoArray)
                remaining.Add photoArray(idx)
                idx = idx + 1
            Loop
            AddPhotoSlides pres, titleText, remaining, AB_PPT_REPORT_LAYOUT_SECTION_PHOTO_3, AB_PPT_REPORT_LAYOUT_SECTION_PHOTO_6
        End If
    End If
End Sub

Private Sub AddFindingSlides(ByVal pres As Object, ByVal titleText As String, ByVal lines As Collection, ByVal layoutName As String)
    Dim idx As Long
    Dim bodyText As String
    Dim slide As Object

    idx = 1
    Do While idx <= lines.Count
        bodyText = JoinLines(lines, idx, AB_PPT_REPORT_FINDINGS_PER_SLIDE)

        Set slide = AddReportSlide(pres, layoutName)
        If slide Is Nothing Then Exit Sub
        SetTitleIfPresent slide, titleText
        SetBodyIfPresent slide, bodyText

        idx = idx + AB_PPT_REPORT_FINDINGS_PER_SLIDE
    Loop
End Sub

Private Sub AddPhotoSlides(ByVal pres As Object, ByVal titleText As String, ByVal photos As Collection, ByVal layout3 As String, ByVal layout6 As String)
    Dim photoArray() As Variant
    Dim idx As Long
    Dim remaining As Long
    Dim layoutName As String
    Dim slots As Long
    Dim slide As Object
    Dim s As Long

    If photos Is Nothing Then Exit Sub
    If photos.Count = 0 Then Exit Sub

    photoArray = CollectionToArrayDistinct(photos)
    idx = LBound(photoArray)
    Do While idx <= UBound(photoArray)
        remaining = UBound(photoArray) - idx + 1
        If remaining <= 3 Then
            layoutName = layout3
            slots = 3
        Else
            layoutName = layout6
            slots = 6
        End If

        Set slide = AddReportSlide(pres, layoutName)
        If slide Is Nothing Then Exit Sub
        SetTitleIfPresent slide, titleText

        For s = 1 To slots
            If idx > UBound(photoArray) Then Exit For
            InsertPictureSlot slide, CStr(photoArray(idx)), s
            idx = idx + 1
        Next s
    Loop
End Sub

Private Function AddReportSlide(ByVal pres As Object, ByVal layoutName As String) As Object
    Dim layout As Object
    Set layout = FindLayoutByName(pres, layoutName)
    If layout Is Nothing Then Exit Function
    Set AddReportSlide = pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
End Function

Private Sub SetBodyIfPresent(ByVal slide As Object, ByVal bodyText As String)
    Dim target As Object
    Set target = FindBestTextPlaceholder(slide)
    If target Is Nothing Then Exit Sub
    target.TextFrame.TextRange.Text = bodyText
End Sub

Private Function FindBestTextPlaceholder(ByVal slide As Object) As Object
    Dim shape As Object
    Dim best As Object
    Dim bestArea As Double
    Dim isTitle As Boolean
    Dim area As Double

    For Each shape In slide.Shapes
        If shape.HasTextFrame Then
            On Error Resume Next
            isTitle = (shape.PlaceholderFormat.Type = 1)
            On Error GoTo 0
            If Not isTitle Then
                area = shape.Width * shape.Height
                If area > bestArea Then
                    bestArea = area
                    Set best = shape
                End If
            End If
        End If
    Next shape
    Set FindBestTextPlaceholder = best
End Function

Private Sub FitPastedShapeToPlaceholder(ByVal slide As Object, ByVal pastedShape As Object)
    Dim placeholder As Object
    Set placeholder = FindPicturePlaceholder(slide)
    If placeholder Is Nothing Then Exit Sub

    On Error Resume Next
    pastedShape.LockAspectRatio = msoTrue
    On Error GoTo 0
    FitShapeInBounds pastedShape, placeholder.Left, placeholder.Top, placeholder.Width, placeholder.Height
    placeholder.Delete
End Sub

Private Sub FitShapeInBounds(ByVal shape As Object, ByVal leftPos As Single, ByVal topPos As Single, ByVal width As Single, ByVal height As Single)
    Dim scaleFactor As Double

    If shape Is Nothing Then Exit Sub
    On Error Resume Next
    scaleFactor = width / shape.Width
    If (shape.Height * scaleFactor) > height Then scaleFactor = height / shape.Height
    shape.Width = shape.Width * scaleFactor
    shape.Height = shape.Height * scaleFactor
    shape.Left = leftPos + (width - shape.Width) / 2
    shape.Top = topPos + (height - shape.Height) / 2
    On Error GoTo 0
End Sub

Private Function FindPicturePlaceholder(ByVal slide As Object) As Object
    Dim shape As Object
    For Each shape In slide.Shapes
        On Error Resume Next
        If shape.Type = AB_MSO_PLACEHOLDER Or shape.Type = AB_PP_PLACEHOLDER_PICTURE Then
            If shape.PlaceholderFormat.Type = AB_PP_PLACEHOLDER_PICTURE Or shape.Type = AB_PP_PLACEHOLDER_PICTURE Then
                Set FindPicturePlaceholder = shape
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next shape
End Function

Private Function GetChapterFirstPageRange(ByVal chapterId As String) As Range
    Dim startName As String
    Dim bm As Bookmark
    Dim pageNum As Long
    Dim pageStart As Range
    Dim pageEnd As Range
    Dim rng As Range

    startName = BuildStartBookmark(chapterId)
    Set bm = FindBookmark(startName)
    If bm Is Nothing Then Exit Function

    pageNum = bm.Range.Information(wdActiveEndPageNumber)

    Set pageStart = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNum)

    On Error Resume Next
    Set pageEnd = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNum + 1)
    On Error GoTo 0

    If pageEnd Is Nothing Then
        Set rng = ActiveDocument.Range(pageStart.Start, ActiveDocument.Content.End)
    ElseIf pageEnd.Start > pageStart.Start Then
        Set rng = ActiveDocument.Range(pageStart.Start, pageEnd.Start - 1)
    Else
        Set rng = ActiveDocument.Range(pageStart.Start, ActiveDocument.Content.End)
    End If

    Set GetChapterFirstPageRange = rng
End Function

Private Function FindBookmark(ByVal name As String) As Bookmark
    On Error Resume Next
    Set FindBookmark = ActiveDocument.Bookmarks(name)
    On Error GoTo 0
End Function

Private Function BuildStartBookmark(ByVal chapterId As String) As String
    BuildStartBookmark = Replace$("Chapter" & chapterId & "_start", ".", "_")
End Function

Private Function BuildSectionList(ByVal rows As Object, ByVal chapterId As String) As Collection
    Dim sections As Collection
    Dim current As Object
    Dim row As Variant
    Dim items As Collection
    Dim newItems As Collection

    Set sections = New Collection
    For Each row In rows
        If IsSectionRow(row) Then
            If Not current Is Nothing Then sections.Add current
            Set current = CreateObject("Scripting.Dictionary")
            current("id") = SafeText(row, "id")
            current("title") = ResolveSectionTitle(row, chapterId)
            Set items = New Collection
            Set current("items") = items
        ElseIf IsIncludedRow(row) Then
            If current Is Nothing Then
                Set current = CreateObject("Scripting.Dictionary")
                current("id") = ResolveSectionId(row, chapterId)
                current("title") = ResolveSectionTitleFromItem(row)
                Set newItems = New Collection
                Set current("items") = newItems
            End If
            current("items").Add row
        End If
    Next row
    If Not current Is Nothing Then sections.Add current
    Set BuildSectionList = sections
End Function

Private Function ResolveSectionTitle(ByVal row As Object, ByVal chapterId As String) As String
    ResolveSectionTitle = SafeText(row, "title")
    If Len(ResolveSectionTitle) = 0 Then ResolveSectionTitle = ResolveSectionId(row, chapterId)
End Function

Private Function ResolveSectionTitleFromItem(ByVal row As Object) As String
    ResolveSectionTitleFromItem = SafeText(row, "sectionLabel")
    If Len(ResolveSectionTitleFromItem) = 0 Then ResolveSectionTitleFromItem = SafeText(row, "sectionId")
End Function

Private Function ResolveChapterTitle(ByVal chapterId As String, ByVal chapter As Object) As String
    Dim titleText As String
    titleText = SafeText(chapter, "title")
    If chapterId = "0" Then
        If Len(titleText) = 0 Then titleText = AB_PPT_REPORT_CH0_TITLE
        ResolveChapterTitle = titleText
        Exit Function
    End If
    ResolveChapterTitle = titleText
End Function

Private Function BuildNumberedTitle(ByVal numberText As String, ByVal titleText As String) As String
    Dim trimmedTitle As String
    trimmedTitle = Trim$(titleText)
    If Len(trimmedTitle) = 0 Then
        BuildNumberedTitle = numberText
    Else
        If InStr(numberText, ".") > 0 Then
            BuildNumberedTitle = numberText & " " & trimmedTitle
        Else
            BuildNumberedTitle = numberText & ". " & trimmedTitle
        End If
    End If
End Function

Private Function JoinLines(ByVal lines As Collection, ByVal startIndex As Long, ByVal count As Long) As String
    Dim parts As String
    Dim i As Long
    For i = startIndex To startIndex + count - 1
        If i > lines.Count Then Exit For
        If Len(parts) > 0 Then parts = parts & vbCrLf & vbCrLf
        parts = parts & CStr(lines(i))
    Next i
    JoinLines = parts
End Function

Private Function BuildFindingTextLines(ByVal items As Object, ByVal renumberMap As Object, ByVal useRecommendation As Boolean, ByVal includeIds As Boolean) As Collection
    Dim lines As Collection
    Dim row As Variant
    Dim textValue As String
    Dim lineText As String
    Dim displayId As String

    Set lines = New Collection
    For Each row In items
        If useRecommendation Then
            textValue = ResolveRecommendation(row)
        Else
            textValue = ResolveFinding(row)
        End If
        textValue = Trim$(textValue)
        If Len(textValue) = 0 Then GoTo ContinueRow

        If includeIds Then
            displayId = ResolveDisplayId(row, renumberMap)
            If Len(displayId) > 0 Then
                lineText = displayId & " " & textValue
            Else
                lineText = textValue
            End If
        Else
            lineText = textValue
        End If
        lines.Add lineText
ContinueRow:
    Next row
    Set BuildFindingTextLines = lines
End Function

Private Function ResolveFinding(ByVal row As Object) As String
    Dim ws As Object
    Dim master As Object

    Set ws = GetObject(row, "workstate")
    If Not ws Is Nothing Then
        If GetBool(ws, "useFindingOverride") Then
            ResolveFinding = SafeText(ws, "findingOverride")
            Exit Function
        End If
    End If

    Set master = GetObject(row, "master")
    If Not master Is Nothing Then
        ResolveFinding = ToPlainText(master("finding"))
    End If
End Function

Private Function ResolveRecommendation(ByVal row As Object) As String
    Dim levelKey As String
    Dim ws As Object
    Dim overrides As Object
    Dim master As Object
    Dim levels As Object

    levelKey = "1"

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
            Set overrides = GetObject(ws, "levelOverrides")
            If Not overrides Is Nothing Then
                If GetBoolFromDict(ws("useLevelOverride"), levelKey) Then
                    ResolveRecommendation = ToPlainText(overrides(levelKey))
                    Exit Function
                End If
            End If
        End If
    End If

    Set master = GetObject(row, "master")
    If Not master Is Nothing Then
        Set levels = GetObject(master, "levels")
        If Not levels Is Nothing Then
            If levels.Exists(levelKey) Then
                ResolveRecommendation = ToPlainText(levels(levelKey))
                Exit Function
            End If
        End If
    End If
End Function

Private Function BuildReportPhotoMap(ByVal photosDoc As Object, ByVal projectFolder As String) As Object
    Dim map As Object
    Dim photos As Object
    Dim key As Variant
    Dim photo As Object
    Dim tags As Object
    Dim absPath As String

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    If photosDoc Is Nothing Then
        Set BuildReportPhotoMap = map
        Exit Function
    End If

    Set photos = GetObject(photosDoc, "photos")
    If photos Is Nothing Then
        Set BuildReportPhotoMap = map
        Exit Function
    End If

    For Each key In photos.Keys
        Set photo = photos(key)
        Set tags = GetObject(photo, "tags")
        If tags Is Nothing Then GoTo NextPhoto

        absPath = projectFolder & "\\" & Replace(CStr(key), "/", "\\")

        AddReportPhotoTags map, GetObject(tags, "report"), absPath
        AddReportPhotoTags map, GetObject(tags, "observations"), absPath
NextPhoto:
    Next key

    Set BuildReportPhotoMap = map
End Function

Private Sub AddReportPhotoTags(ByVal map As Object, ByVal tagList As Object, ByVal absPath As String)
    Dim tag As Variant
    Dim tagKey As String
    Dim coll As Collection

    If tagList Is Nothing Then Exit Sub
    For Each tag In tagList
        tagKey = CStr(tag)
        If Len(tagKey) = 0 Then GoTo ContinueTag
        If Not map.Exists(tagKey) Then
            Set coll = New Collection
            map.Add tagKey, coll
        End If
        map(tagKey).Add absPath
ContinueTag:
    Next tag
End Sub

Private Function GetPhotoCollection(ByVal map As Object, ByVal sectionId As String) As Collection
    Dim coll As Collection
    If map Is Nothing Then
        Set GetPhotoCollection = New Collection
        Exit Function
    End If
    If map.Exists(sectionId) Then
        Set GetPhotoCollection = map(sectionId)
    Else
        Set GetPhotoCollection = New Collection
    End If
End Function

Private Sub CacheSectionMap(ByVal cache As Object, ByVal chapterId As String, ByVal renumberMap As Object)
    If cache Is Nothing Then Exit Sub
    On Error Resume Next
    If Not renumberMap Is Nothing Then
        If renumberMap.Exists("_sectionMap") Then
            cache(chapterId) = renumberMap("_sectionMap")
        End If
    End If
    On Error GoTo 0
End Sub

Private Function ResolveSpecial48DisplaySectionId(ByVal chapters As Object, ByVal cache As Object) As String
    Dim sectionMap As Object
    Set sectionMap = GetSectionMapForChapter(chapters, "4", cache)
    If sectionMap Is Nothing Then Exit Function
    ResolveSpecial48DisplaySectionId = "4." & CStr(sectionMap.Count + 1)
End Function

Private Function GetSectionMapForChapter(ByVal chapters As Object, ByVal chapterId As String, ByVal cache As Object) As Object
    Dim chapter As Variant
    Dim rows As Object
    Dim renumberMap As Object

    If Not cache Is Nothing Then
        If cache.Exists(chapterId) Then
            Set GetSectionMapForChapter = cache(chapterId)
            Exit Function
        End If
    End If

    For Each chapter In chapters
        If SafeText(chapter, "id") = chapterId Then
            Set rows = GetObject(chapter, "rows")
            Set renumberMap = BuildRenumberMap(rows, chapterId)
            If Not renumberMap Is Nothing Then
                If renumberMap.Exists("_sectionMap") Then
                    Set GetSectionMapForChapter = renumberMap("_sectionMap")
                    If Not cache Is Nothing Then cache(chapterId) = renumberMap("_sectionMap")
                    Exit Function
                End If
            End If
        End If
    Next chapter
End Function

Private Function BuildRenumberMap(ByVal rows As Object, ByVal chapterId As String) As Object
    Dim map As Object
    Dim sectionMap As Object
    Dim sectionCounts As Object
    Dim itemCount As Long
    Dim row As Variant
    Dim rowId As String
    Dim sectionKey As String
    Dim count As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set sectionMap = CreateObject("Scripting.Dictionary")
    Set sectionCounts = CreateObject("Scripting.Dictionary")

    For Each row In rows
        If IsSectionRow(row) Then
            ' section rows handled via sectionMap
        ElseIf IsIncludedRow(row) Then
            rowId = SafeText(row, "id")
            If Len(rowId) = 0 Then GoTo ContinueLoop
            If IsFieldObservationChapter(chapterId) Then
                itemCount = itemCount + 1
                map(rowId) = chapterId & "." & CStr(itemCount)
            Else
                sectionKey = ResolveSectionId(row, chapterId)
                If Len(sectionKey) = 0 Then sectionKey = chapterId & ".1"
                If Not sectionMap.Exists(sectionKey) Then
                    sectionMap(sectionKey) = sectionMap.Count + 1
                End If
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

    map("_sectionMap") = sectionMap
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

Private Function ResolveSectionDisplayId(ByVal sectionId As String, ByVal renumberMap As Object) As String
    Dim sectionMap As Object
    Dim chapterPart As String

    If Len(sectionId) = 0 Then Exit Function
    On Error Resume Next
    If Not renumberMap Is Nothing Then
        If renumberMap.Exists("_sectionMap") Then
            Set sectionMap = renumberMap("_sectionMap")
            If Not sectionMap Is Nothing Then
                If sectionMap.Exists(sectionId) Then
                    chapterPart = Split(sectionId, ".")(0)
                    ResolveSectionDisplayId = chapterPart & "." & CStr(sectionMap(sectionId))
                End If
            End If
        End If
    End If
    On Error GoTo 0
    If Len(ResolveSectionDisplayId) = 0 Then ResolveSectionDisplayId = sectionId
End Function

Private Function ResolveSectionId(ByVal row As Object, ByVal chapterId As String) As String
    Dim rowId As String
    Dim parts() As String

    On Error Resume Next
    If row.Exists("sectionId") Then
        ResolveSectionId = CStr(row("sectionId"))
        Exit Function
    End If
    On Error GoTo 0

    rowId = SafeText(row, "id")
    If IsFieldObservationChapter(chapterId) Then
        ResolveSectionId = rowId
        Exit Function
    End If

    parts = Split(rowId, ".")
    If UBound(parts) >= 1 Then
        ResolveSectionId = parts(0) & "." & parts(1)
    End If
End Function

Private Function IsFieldObservationChapter(ByVal chapterId As String) As Boolean
    If InStr(chapterId, ".") > 0 Then
        IsFieldObservationChapter = True
    End If
End Function

Private Function IsSectionRow(ByVal row As Object) As Boolean
    On Error GoTo SafeExit
    If row.Exists("kind") Then
        IsSectionRow = (LCase$(CStr(row("kind"))) = "section")
    End If
SafeExit:
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

Private Function SafeText(ByVal dict As Object, ByVal key As String) As String
    On Error GoTo SafeExit
    If dict Is Nothing Then Exit Function
    If dict.Exists(key) Then
        SafeText = ResolveLocalizedText(dict(key))
    End If
SafeExit:
End Function

Private Function ResolveLocalizedText(ByVal value As Variant) As String
    Dim langKey As Variant
    Dim key As Variant

    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            For Each langKey In Array("de", "fr", "it", "en")
                If value.Exists(langKey) Then
                    ResolveLocalizedText = ToPlainText(value(langKey))
                    Exit Function
                End If
            Next langKey
            For Each key In value.Keys
                ResolveLocalizedText = ToPlainText(value(key))
                Exit Function
            Next key
        End If
    End If
    ResolveLocalizedText = ToPlainText(value)
End Function

Private Function ToPlainText(ByVal value As Variant) As String
    Dim parts As Collection
    Dim item As Variant
    Dim buff As String

    If IsObject(value) Then
        If TypeName(value) = "Collection" Then
            Set parts = value
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

Private Function ResolveSidecarPathPpt() As String
    Dim defaultPath As String
    Dim fd As FileDialog

    If Len(ActiveDocument.Path) > 0 Then
        defaultPath = ActiveDocument.Path & "\\" & AB_SIDECAR_FILENAME
        If Dir(defaultPath) <> "" Then
            ResolveSidecarPathPpt = defaultPath
            Exit Function
        End If
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = AB_SIDECAR_DIALOG_TITLE
    fd.Filters.Clear
    fd.Filters.Add "JSON", "*.json"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Function
    ResolveSidecarPathPpt = fd.SelectedItems(1)
End Function

Private Function GetParentFolder(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = fso.GetParentFolderName(filePath)
End Function

Private Function ReadAllTextUtf8(ByVal path As String) As String
    Dim stream As Object

    On Error GoTo CleanFail
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

Private Sub LogDebug(ByVal message As String)
    If Not AB_DEBUG_PPT_EXPORT Then Exit Sub
    Debug.Print Format$(Now, "hh:nn:ss") & " | " & message
End Sub
