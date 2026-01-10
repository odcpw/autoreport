Attribute VB_Name = "modPptExport"
Option Explicit

' === PPT EXPORT CONFIG ===
Private Const DEBUG_ENABLED As Boolean = True
Private Const TEMPLATE_FOLDER As String = "Templates"
Private Const TRAINING_TEMPLATE_D As String = "Training_D.pptx"
Private Const TRAINING_TEMPLATE_F As String = "Training_F.pptx"
Private Const OUTPUT_TRAINING_BASE As String = "Seminar_Slides"
Private Const INCLUDE_SEMINAR_SLIDE As Boolean = True

' Layout names (match template)
Private Const LAYOUT_CHAPTER As String = "chapterorange"
Private Const LAYOUT_SEMINAR As String = "seminar_d"
Private Const LAYOUT_PICTURE As String = "picture"

Public Sub ExportTrainingPptD()
    ExportTrainingPptInternal "D", TRAINING_TEMPLATE_D
End Sub

Public Sub ExportTrainingPptF()
    ExportTrainingPptInternal "F", TRAINING_TEMPLATE_F
End Sub

Private Sub ExportTrainingPptInternal(ByVal langSuffix As String, ByVal templateFile As String)
    LogDebug "ExportTrainingPpt: start (" & langSuffix & ")"
    Dim sidecarPath As String
    sidecarPath = ResolveSidecarPathPpt()
    If Len(sidecarPath) = 0 Then Exit Sub

    Dim sidecarText As String
    sidecarText = ReadAllTextUtf8(sidecarPath)
    If Len(sidecarText) = 0 Then
        MsgBox "Sidecar JSON is empty.", vbExclamation
        Exit Sub
    End If

    Dim root As Object
    Set root = JsonConverter.ParseJson(sidecarText)
    Dim photosDoc As Object
    Set photosDoc = GetObject(root, "photos")
    If photosDoc Is Nothing Then
        MsgBox "No photos section in sidecar.", vbExclamation
        Exit Sub
    End If

    Dim projectFolder As String
    projectFolder = GetParentFolder(sidecarPath)
    Dim templatesFolder As String
    templatesFolder = projectFolder & "\\" & TEMPLATE_FOLDER
    Dim templatePath As String
    templatePath = templatesFolder & "\\" & templateFile
    If Dir(templatePath) = "" Then
        MsgBox "Training template not found: " & templatePath, vbExclamation
        Exit Sub
    End If

    Dim pptApp As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True

    Dim pres As Object
    Set pres = pptApp.Presentations.Open(templatePath, 0, 0, -1)

    If INCLUDE_SEMINAR_SLIDE Then
        Dim seminarLayout As Object
        Set seminarLayout = FindLayoutByName(pres, LAYOUT_SEMINAR)
        If Not seminarLayout Is Nothing Then
            Dim seminarSlide As Object
            Set seminarSlide = pres.Slides.AddSlide(pres.Slides.Count + 1, seminarLayout)
            SetTitleIfPresent seminarSlide, "Seminar"
        End If
    End If

    Dim tagMap As Object
    Set tagMap = BuildTrainingTagMap(photosDoc, projectFolder)

    Dim orderedTags As Collection
    Set orderedTags = BuildTrainingTagOrder(tagMap)

    Dim tag As Variant
    For Each tag In orderedTags
        Dim photos As Collection
        Set photos = tagMap(tag)
        If photos.Count = 0 Then GoTo ContinueTag

        Dim chapterLayout As Object
        Set chapterLayout = FindLayoutByName(pres, LAYOUT_CHAPTER)
        If Not chapterLayout Is Nothing Then
            Dim chapterSlide As Object
            Set chapterSlide = pres.Slides.AddSlide(pres.Slides.Count + 1, chapterLayout)
            SetTitleIfPresent chapterSlide, CStr(tag)
        End If

        Dim layoutName As String
        layoutName = ResolveTrainingLayoutName(CStr(tag), langSuffix)
        Dim layout As Object
        Set layout = FindLayoutByName(pres, layoutName)
        If layout Is Nothing Then
            Set layout = FindLayoutByName(pres, LAYOUT_PICTURE)
        End If

        Dim photoPath As Variant
        For Each photoPath In photos
            Dim slide As Object
            Set slide = pres.Slides.AddSlide(pres.Slides.Count + 1, layout)
            SetTitleIfPresent slide, CStr(tag)
            InsertFirstPicture slide, CStr(photoPath)
        Next photoPath
ContinueTag:
    Next tag

    Dim outPath As String
    outPath = projectFolder & "\\" & Format$(Date, "yyyy-mm-dd") & "_" & OUTPUT_TRAINING_BASE & "_" & langSuffix & ".pptx"
    pres.SaveAs outPath
    MsgBox "Training deck exported to: " & outPath, vbInformation
    LogDebug "ExportTrainingPpt: done"
End Sub

Private Function BuildTrainingTagMap(ByVal photosDoc As Object, ByVal projectFolder As String) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    Dim photos As Object
    Set photos = GetObject(photosDoc, "photos")
    If photos Is Nothing Then
        Set BuildTrainingTagMap = map
        Exit Function
    End If
    Dim key As Variant
    For Each key In photos.Keys
        Dim photo As Object
        Set photo = photos(key)
        Dim tags As Object
        Set tags = GetObject(photo, "tags")
        If tags Is Nothing Then GoTo NextPhoto
        Dim training As Object
        Set training = GetObject(tags, "training")
        If training Is Nothing Then GoTo NextPhoto
        Dim absPath As String
        absPath = projectFolder & "\\" & Replace(CStr(key), "/", "\\")
        Dim t As Variant
        For Each t In training
            If Not map.Exists(CStr(t)) Then
                Dim coll As New Collection
                map.Add CStr(t), coll
            End If
            map(CStr(t)).Add absPath
        Next t
NextPhoto:
    Next key
    Set BuildTrainingTagMap = map
End Function

Private Function BuildTrainingTagOrder(ByVal tagMap As Object) As Collection
    Dim order As New Collection
    Dim knownTags As Variant
    knownTags = Array("Unterlassen", "Dulden", "Handeln", "Vorbild", "Iceberg", "Pyramide", "STOP", "SOS", _
                      "Verhindern", "Audit", "Risikobeurteilung", "AVIVA")
    Dim i As Long
    For i = LBound(knownTags) To UBound(knownTags)
        If tagMap.Exists(knownTags(i)) Then order.Add knownTags(i)
    Next i
    Dim key As Variant
    For Each key In tagMap.Keys
        If Not ExistsInCollection(order, CStr(key)) Then order.Add CStr(key)
    Next key
    Set BuildTrainingTagOrder = order
End Function

Private Function ResolveTrainingLayoutName(ByVal tag As String, ByVal langSuffix As String) As String
    Dim suffix As String
    suffix = LCase$(langSuffix)
    If suffix <> "d" And suffix <> "f" Then suffix = "d"
    Select Case LCase$(tag)
        Case "unterlassen": ResolveTrainingLayoutName = "unterlassen_" & suffix
        Case "dulden": ResolveTrainingLayoutName = "dulden_" & suffix
        Case "handeln": ResolveTrainingLayoutName = "handeln_" & suffix
        Case "vorbild": ResolveTrainingLayoutName = "vorbild_" & suffix
        Case "audit": ResolveTrainingLayoutName = "audit_" & suffix
        Case "risikobeurteilung": ResolveTrainingLayoutName = "risikobeurteilung_" & suffix
        Case "aviva": ResolveTrainingLayoutName = "aviva_" & suffix
        Case "verhindern": ResolveTrainingLayoutName = "verhindern_" & suffix
        Case "iceberg", "pyramide", "stop", "sos": ResolveTrainingLayoutName = "picture"
        Case Else: ResolveTrainingLayoutName = "picture"
    End Select
End Function

Private Function FindLayoutByName(ByVal pres As Object, ByVal layoutName As String) As Object
    Dim d As Object
    For Each d In pres.Designs
        Dim layout As Object
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

Private Sub InsertFirstPicture(ByVal slide As Object, ByVal filePath As String)
    Dim shape As Object
    For Each shape In slide.Shapes
        On Error Resume Next
        If shape.Type = 14 Then
            shape.Fill.UserPicture filePath
            Exit Sub
        End If
        If shape.PlaceholderFormat.Type = 18 Then
            shape.Fill.UserPicture filePath
            Exit Sub
        End If
        On Error GoTo 0
    Next shape
End Sub

Private Function ExistsInCollection(ByVal coll As Collection, ByVal value As String) As Boolean
    Dim item As Variant
    For Each item In coll
        If LCase$(CStr(item)) = LCase$(value) Then
            ExistsInCollection = True
            Exit Function
        End If
    Next item
End Function

Private Function ResolveSidecarPathPpt() As String
    Dim defaultPath As String
    If Len(ActiveDocument.Path) > 0 Then
        defaultPath = ActiveDocument.Path & "\\project_sidecar.json"
        If Dir(defaultPath) <> "" Then
            ResolveSidecarPathPpt = defaultPath
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
    ResolveSidecarPathPpt = fd.SelectedItems(1)
End Function

Private Function GetParentFolder(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = fso.GetParentFolderName(filePath)
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
    If Not DEBUG_ENABLED Then Exit Sub
    Debug.Print Format$(Now, "hh:nn:ss") & " | " & message
End Sub
