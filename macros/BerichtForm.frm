VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BerichtForm 
   Caption         =   "BerichtForm"
   ClientHeight    =   11310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22440
   OleObjectBlob   =   "BerichtForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BerichtForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==================== STATE ====================

Private didRender   As Boolean
Private currentChapter As Long
Private currentPage As Long
Private totalPages As Long

Private rowSeq      As Long
Private RowUIs      As Collection
Private RowFrames   As Collection

Private chapterItems As Collection
Private lastParentID As String  ' Track when parent context changes

'==================== LIFECYCLE ====================

Private Sub UserForm_Initialize()
    modReportData.LoadAllCaches
    currentChapter = 1
    currentPage = 1
End Sub

Private Sub UserForm_Activate()
    If Not didRender Then
        didRender = True
        LoadChapterData currentChapter
        RenderCurrentPage
        modBerichtUI.UpdateChapterButtons Me, currentChapter
        modBerichtUI.UpdatePageButtons Me, currentPage, totalPages
    End If
End Sub

'==================== CHAPTER BUTTON HANDLERS ====================

Private Sub btnChap1_Click(): SelectChapter 1: End Sub
Private Sub btnChap2_Click(): SelectChapter 2: End Sub
Private Sub btnChap3_Click(): SelectChapter 3: End Sub
Private Sub btnChap4_Click(): SelectChapter 4: End Sub
Private Sub btnChap5_Click(): SelectChapter 5: End Sub
Private Sub btnChap6_Click(): SelectChapter 6: End Sub
Private Sub btnChap7_Click(): SelectChapter 7: End Sub
Private Sub btnChap8_Click(): SelectChapter 8: End Sub
Private Sub btnChap9_Click(): SelectChapter 9: End Sub
Private Sub btnChap10_Click(): SelectChapter 10: End Sub
Private Sub btnChap11_Click(): SelectChapter 11: End Sub

Private Sub SelectChapter(ByVal chapterNum As Long)
    ShowChapter chapterNum
    modBerichtUI.UpdateChapterButtons Me, currentChapter
    modBerichtUI.UpdatePageButtons Me, currentPage, totalPages
End Sub

'==================== PAGE BUTTON HANDLERS ====================

Private Sub btnPage1_Click(): SelectPage 1: End Sub
Private Sub btnPage2_Click(): SelectPage 2: End Sub
Private Sub btnPage3_Click(): SelectPage 3: End Sub
Private Sub btnPage4_Click(): SelectPage 4: End Sub
Private Sub btnPage5_Click(): SelectPage 5: End Sub
Private Sub btnPage6_Click(): SelectPage 6: End Sub
Private Sub btnPage7_Click(): SelectPage 7: End Sub
Private Sub btnPage8_Click(): SelectPage 8: End Sub
Private Sub btnPage9_Click(): SelectPage 9: End Sub
Private Sub btnPage10_Click(): SelectPage 10: End Sub
Private Sub btnPage11_Click(): SelectPage 11: End Sub

Private Sub SelectPage(ByVal pageNum As Long)
    ShowPage pageNum
    modBerichtUI.UpdatePageButtons Me, currentPage, totalPages
End Sub

'==================== PUBLIC NAVIGATION ====================

Public Sub ShowChapter(ByVal chapterNum As Long)
    If chapterNum < 1 Or chapterNum > 11 Then Exit Sub
    currentChapter = chapterNum
    currentPage = 1
    LoadChapterData currentChapter
    RenderCurrentPage
End Sub

Public Sub ShowNextPage()
    If currentPage < totalPages Then
        currentPage = currentPage + 1
        RenderCurrentPage
    End If
End Sub

Public Sub ShowPreviousPage()
    If currentPage > 1 Then
        currentPage = currentPage - 1
        RenderCurrentPage
    End If
End Sub

Public Sub ShowPage(ByVal pageNum As Long)
    If pageNum >= 1 And pageNum <= totalPages Then
        currentPage = pageNum
        RenderCurrentPage
    End If
End Sub

'==================== DATA LOADING ====================

Private Sub LoadChapterData(ByVal chapterNumber As Long)
    Set chapterItems = New Collection
    
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets("SelbstbeurteilungKunde").ListObjects("TableSelbstbeurteilung")
    Dim prefix As String: prefix = CStr(chapterNumber) & "."
    
    Dim r As ListRow, id As String
    For Each r In lo.ListRows
        id = modReportData.CleanID(CStr(r.Range.Cells(1, 1).Value))
        If Len(id) = 0 Then GoTo NextRow
        
        If Left$(id, Len(prefix)) <> prefix Then GoTo NextRow
        If Not IsLevel3Leaf(id) Then GoTo NextRow
        
        chapterItems.Add id
        
NextRow:
    Next r
    
    If chapterItems.Count = 0 Then
        totalPages = 1
    Else
        totalPages = Int((chapterItems.Count - 1) / modBerichtUI.ROWS_PER_PAGE) + 1
    End If
    
    On Error Resume Next
    Me.caption = "Bericht - Kapitel " & chapterNumber & " (Seite " & currentPage & "/" & totalPages & ")"
    On Error GoTo 0
End Sub

'==================== RENDERING ====================

Private Sub RenderCurrentPage()
    On Error GoTo FAIL
    
    ClearAllRowsOnForm
    Set RowUIs = New Collection
    Set RowFrames = New Collection
    rowSeq = 0
    lastParentID = ""  ' Reset parent tracking
    
    On Error Resume Next
    Me.Controls("RowTemplate").Visible = False
    On Error GoTo 0
    
    Dim startIdx As Long, endIdx As Long
    startIdx = (currentPage - 1) * modBerichtUI.ROWS_PER_PAGE + 1
    endIdx = startIdx + modBerichtUI.ROWS_PER_PAGE - 1
    If endIdx > chapterItems.Count Then endIdx = chapterItems.Count
    
    Dim y As Single: y = modBerichtUI.START_TOP
    Dim lineNo As Long: lineNo = startIdx
    
    Dim i As Long, id As String, s As modReportData.SBRow
    For i = startIdx To endIdx
        id = chapterItems(i)
        s = modReportData.GetSB(id)
        
        Dim fra As MSForms.Frame
        Set fra = CloneRowTemplate(Me, y, "Row" & CStr(rowSeq))
        
        WireRowControls fra, lineNo, id, s
        
        y = y + modBerichtUI.ROW_HEIGHT + modBerichtUI.ROW_GAP
        lineNo = lineNo + 1
    Next i
    
    Exit Sub
    
FAIL:
    MsgBox "RenderCurrentPage error: " & Err.description, vbCritical
End Sub

Private Sub WireRowControls(ByVal fra As MSForms.Frame, ByVal lineNo As Long, _
                             ByVal id As String, ByRef s As modReportData.SBRow)
    
    Dim lbl As MSForms.Label, txt As MSForms.TextBox, chk As MSForms.CheckBox
    
    ' Line number
    Set lbl = FindLabel(fra, modBerichtUI.CN_LineNumber)
    If Not lbl Is Nothing Then
        lbl.caption = CStr(lineNo)
        lbl.ZOrder 0
    End If
    
    ' Parent context (LblID2) - show when parent changes
    Dim parentID As String: parentID = modBerichtUI.GetParentID(id)
    Dim showParent As Boolean: showParent = (parentID <> lastParentID And Len(parentID) > 0)
    
    If showParent Then
        lastParentID = parentID
        Dim parentFest As String: parentFest = modReportData.GetMasterFeststellungByID(parentID)
        
        Set lbl = FindLabel(fra, modBerichtUI.CN_ReportItemID2)
        If Not lbl Is Nothing Then
            lbl.caption = modBerichtUI.FormatParentContext(parentID, parentFest)
            lbl.Visible = True
            lbl.ZOrder 0
        End If
    Else
        ' Hide parent label if same as previous row
        Set lbl = FindLabel(fra, modBerichtUI.CN_ReportItemID2)
        If Not lbl Is Nothing Then lbl.Visible = False
    End If
    
    ' Current item ID (LblID3)
    Set lbl = FindLabel(fra, modBerichtUI.CN_ReportItemID3)
    If Not lbl Is Nothing Then
        lbl.caption = id
        lbl.ZOrder 0
    End If
    
    ' Customer answer
    Set lbl = FindLabel(fra, modBerichtUI.CN_AntwortKunde)
    If Not lbl Is Nothing Then
        lbl.caption = s.antwort
        lbl.ZOrder 0
    End If
    
    ' Left: Finding
    Set txt = FindTextBox(fra, modBerichtUI.CN_FestMaster)
    If Not txt Is Nothing Then
        txt.Locked = True
        txt.Text = modReportData.GetMasterFinding(id)
    End If
    
    ' Right: Level recommendation
    Dim initLevel As Long: initLevel = IIf(s.SelectedLevel = 0, 1, s.SelectedLevel)
    Dim levelTxt As String: levelTxt = modReportData.GetMasterLevelText(id, initLevel)
    
    If Len(levelTxt) > 0 Then
        Set txt = FindTextBox(fra, modBerichtUI.CN_LevelMaster)
        If Not txt Is Nothing Then
            txt.Locked = True
            txt.Text = levelTxt
        End If
        
        Set chk = FindCheckBox(fra, modBerichtUI.CN_Bericht)
        If Not chk Is Nothing Then chk.Value = s.selected
        
        ' Wire events
        Dim row As New CRowUI
        Set row.LblLine = FindLabel(fra, modBerichtUI.CN_LineNumber)
        Set row.LblID = FindLabel(fra, modBerichtUI.CN_ReportItemID3)
        Set row.LblAntwort = FindLabel(fra, modBerichtUI.CN_AntwortKunde)
        
        Set row.ChkInclude = FindCheckBox(fra, modBerichtUI.CN_Bericht)
        Set row.Opt1 = FindOption(fra, modBerichtUI.CN_Opt1)
        Set row.Opt2 = FindOption(fra, modBerichtUI.CN_Opt2)
        Set row.Opt3 = FindOption(fra, modBerichtUI.CN_Opt3)
        Set row.Opt4 = FindOption(fra, modBerichtUI.CN_Opt4)
        
        Set row.txtFestMaster = FindTextBox(fra, modBerichtUI.CN_FestMaster)
        Set row.TxtLevelMaster = FindTextBox(fra, modBerichtUI.CN_LevelMaster)
        Set row.TxtFestEdit = FindTextBox(fra, modBerichtUI.CN_FestEdit)
        Set row.TxtLevelEdit = FindTextBox(fra, modBerichtUI.CN_LevelEdit)
        
        Set row.ChkUseFestOverride = FindCheckBox(fra, modBerichtUI.CN_UseFestOverride)
        Set row.ChkUseLevelOverride = FindCheckBox(fra, modBerichtUI.CN_UseLevelOverride)
        
        On Error Resume Next
        Set row.BtnPromoteMaster = FindButton(fra, modBerichtUI.CN_BtnPromote)
        On Error GoTo 0
        
        row.Init id, s.antwort, s.selected, initLevel, _
                 modReportData.GetMasterFinding(id), levelTxt
        
        RowUIs.Add row
    Else
        HideIfExists fra, modBerichtUI.CN_LevelMaster
        HideIfExists fra, modBerichtUI.CN_LevelEdit
        HideIfExists fra, modBerichtUI.CN_UseLevelOverride
        HideIfExists fra, modBerichtUI.CN_Bericht
        HideIfExists fra, modBerichtUI.CN_Opt1
        HideIfExists fra, modBerichtUI.CN_Opt2
        HideIfExists fra, modBerichtUI.CN_Opt3
        HideIfExists fra, modBerichtUI.CN_Opt4
    End If
End Sub

'==================== TEMPLATE CLONING ====================

Private Function CloneRowTemplate(ByVal hostForm As Object, ByVal topY As Single, _
                                   ByVal groupName As String) As MSForms.Frame
    On Error GoTo FAIL
    
    Dim tpl As MSForms.Frame
    Set tpl = Me.Controls("RowTemplate")
    
    rowSeq = rowSeq + 1
    Dim fraName As String: fraName = "Row_" & CStr(rowSeq)
    Dim fra As MSForms.Frame
    Set fra = hostForm.Controls.Add("Forms.Frame.1", fraName, True)
    
    With fra
        .Left = tpl.Left
        .Top = topY
        .Width = tpl.Width
        .Height = tpl.Height
        .caption = tpl.caption
        .BackColor = tpl.BackColor
        .BorderStyle = tpl.BorderStyle
        .Visible = True
        .Enabled = True
    End With
    
    RowFrames.Add fraName
    
    Dim ctl As MSForms.Control
    For Each ctl In tpl.Controls
        If Not IsControlDirectChild(ctl, tpl) Then GoTo NextControl
        
        If typeName(ctl) = "MultiPage" Then
            CloneMultiPage ctl, fra
        Else
            CloneControl ctl, fra, groupName
        End If
        
NextControl:
    Next ctl
    
    Set CloneRowTemplate = fra
    Exit Function
    
FAIL:
    MsgBox "CloneRowTemplate error: " & Err.description, vbCritical
End Function

Private Sub CloneMultiPage(ByVal srcMP As MSForms.MultiPage, ByVal parentFrame As MSForms.Frame)
    On Error GoTo FAIL
    
    Dim dstMP As MSForms.MultiPage
    Set dstMP = parentFrame.Controls.Add("Forms.MultiPage.1")
    
    With dstMP
        .Left = srcMP.Left
        .Top = srcMP.Top
        .Width = srcMP.Width
        .Height = srcMP.Height
        .Visible = srcMP.Visible
        .Enabled = srcMP.Enabled
    End With
    
    Do While dstMP.Pages.Count > 0
        dstMP.Pages.Remove 0
    Loop
    
    Dim i As Long, srcPage As MSForms.Page, dstPage As MSForms.Page
    For i = 0 To srcMP.Pages.Count - 1
        Set srcPage = srcMP.Pages(i)
        Set dstPage = dstMP.Pages.Add
        
        dstPage.caption = srcPage.caption
        dstPage.Visible = srcPage.Visible
        
        Dim ctl As MSForms.Control
        For Each ctl In srcPage.Controls
            CloneControlToPage ctl, dstPage, ""
        Next ctl
    Next i
    
    dstMP.Value = srcMP.Value
    Exit Sub
    
FAIL:
    MsgBox "CloneMultiPage error: " & Err.description, vbCritical
End Sub

Private Sub CloneControl(ByVal src As MSForms.Control, ByVal parentFrame As MSForms.Frame, _
                         ByVal groupName As String)
    On Error Resume Next
    
    Dim progID As String: progID = ProgIDFor(typeName(src))
    If progID = "" Then Exit Sub
    
    Dim dst As MSForms.Control
    Set dst = parentFrame.Controls.Add(progID)
    If dst Is Nothing Then Exit Sub
    
    CopyControlProperties src, dst
    
    If typeName(src) = "OptionButton" Then
        Dim opt As MSForms.OptionButton
        Set opt = dst
        opt.groupName = groupName
    End If
    
    On Error GoTo 0
End Sub

Private Sub CloneControlToPage(ByVal src As MSForms.Control, ByVal dstPage As MSForms.Page, _
                                ByVal groupName As String)
    On Error Resume Next
    
    Dim progID As String: progID = ProgIDFor(typeName(src))
    If progID = "" Then Exit Sub
    
    Dim dst As MSForms.Control
    Set dst = dstPage.Controls.Add(progID)
    If dst Is Nothing Then Exit Sub
    
    CopyControlProperties src, dst
    
    If typeName(src) = "OptionButton" Then
        Dim opt As MSForms.OptionButton
        Set opt = dst
        opt.groupName = groupName
    End If
    
    On Error GoTo 0
End Sub

Private Sub CopyControlProperties(ByVal src As MSForms.Control, ByVal dst As MSForms.Control)
    On Error Resume Next
    
    dst.Left = src.Left
    dst.Top = src.Top
    dst.Width = src.Width
    dst.Height = src.Height
    
    dst.Name = src.Name
    dst.Tag = src.Tag
    dst.Visible = src.Visible
    dst.Enabled = src.Enabled
    
    Select Case typeName(src)
        Case "Label"
            dst.caption = src.caption
            dst.BackColor = src.BackColor
            dst.ForeColor = src.ForeColor
            dst.Font.Name = src.Font.Name
            dst.Font.Size = src.Font.Size
            dst.Font.Bold = src.Font.Bold
            
        Case "TextBox"
            dst.Text = src.Text
            dst.Locked = src.Locked
            dst.MultiLine = src.MultiLine
            dst.ScrollBars = src.ScrollBars
            dst.BackColor = src.BackColor
            dst.ForeColor = src.ForeColor
            dst.Font.Name = src.Font.Name
            dst.Font.Size = src.Font.Size
            
        Case "CheckBox", "OptionButton"
            dst.caption = src.caption
            dst.Value = src.Value
            dst.BackColor = src.BackColor
            dst.ForeColor = src.ForeColor
            dst.Font.Name = src.Font.Name
            dst.Font.Size = src.Font.Size
            
        Case "CommandButton"
            dst.caption = src.caption
            dst.BackColor = src.BackColor
            dst.Font.Name = src.Font.Name
            dst.Font.Size = src.Font.Size
    End Select
    
    On Error GoTo 0
End Sub

Private Function ProgIDFor(ByVal typeName As String) As String
    Select Case typeName
        Case "Label":         ProgIDFor = "Forms.Label.1"
        Case "TextBox":       ProgIDFor = "Forms.TextBox.1"
        Case "CheckBox":      ProgIDFor = "Forms.CheckBox.1"
        Case "OptionButton":  ProgIDFor = "Forms.OptionButton.1"
        Case "CommandButton": ProgIDFor = "Forms.CommandButton.1"
        Case "Frame":         ProgIDFor = "Forms.Frame.1"
        Case "MultiPage":     ProgIDFor = "Forms.MultiPage.1"
        Case Else:            ProgIDFor = ""
    End Select
End Function

Private Function IsControlDirectChild(ByVal ctl As MSForms.Control, _
                                      ByVal parentFrame As MSForms.Frame) As Boolean
    On Error Resume Next
    IsControlDirectChild = (ctl.parent Is parentFrame)
    On Error GoTo 0
End Function

'==================== FIND HELPERS ====================

Private Function FindLabel(ByVal parent As MSForms.Frame, ByVal key As String) As MSForms.Label
    Dim c As MSForms.Control
    For Each c In parent.Controls
        If typeName(c) = "Label" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindLabel = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindLabel = FindLabel(c, key)
            If Not FindLabel Is Nothing Then Exit Function
        ElseIf typeName(c) = "MultiPage" Then
            Dim mp As MSForms.MultiPage, pg As MSForms.Page
            Set mp = c
            For Each pg In mp.Pages
                Set FindLabel = FindLabelInPage(pg, key)
                If Not FindLabel Is Nothing Then Exit Function
            Next
        End If
    Next
End Function

Private Function FindLabelInPage(ByVal Page As MSForms.Page, ByVal key As String) As MSForms.Label
    Dim c As MSForms.Control
    For Each c In Page.Controls
        If typeName(c) = "Label" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindLabelInPage = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindLabelInPage = FindLabel(c, key)
            If Not FindLabelInPage Is Nothing Then Exit Function
        End If
    Next
End Function

Private Function FindTextBox(ByVal parent As MSForms.Frame, ByVal key As String) As MSForms.TextBox
    Dim c As MSForms.Control
    For Each c In parent.Controls
        If typeName(c) = "TextBox" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindTextBox = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindTextBox = FindTextBox(c, key)
            If Not FindTextBox Is Nothing Then Exit Function
        ElseIf typeName(c) = "MultiPage" Then
            Dim mp As MSForms.MultiPage, pg As MSForms.Page
            Set mp = c
            For Each pg In mp.Pages
                Set FindTextBox = FindTextBoxInPage(pg, key)
                If Not FindTextBox Is Nothing Then Exit Function
            Next
        End If
    Next
End Function

Private Function FindTextBoxInPage(ByVal Page As MSForms.Page, ByVal key As String) As MSForms.TextBox
    Dim c As MSForms.Control
    For Each c In Page.Controls
        If typeName(c) = "TextBox" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindTextBoxInPage = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindTextBoxInPage = FindTextBox(c, key)
            If Not FindTextBoxInPage Is Nothing Then Exit Function
        End If
    Next
End Function

Private Function FindCheckBox(ByVal parent As MSForms.Frame, ByVal key As String) As MSForms.CheckBox
    Dim c As MSForms.Control
    For Each c In parent.Controls
        If typeName(c) = "CheckBox" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindCheckBox = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindCheckBox = FindCheckBox(c, key)
            If Not FindCheckBox Is Nothing Then Exit Function
        ElseIf typeName(c) = "MultiPage" Then
            Dim mp As MSForms.MultiPage, pg As MSForms.Page
            Set mp = c
            For Each pg In mp.Pages
                Set FindCheckBox = FindCheckBoxInPage(pg, key)
                If Not FindCheckBox Is Nothing Then Exit Function
            Next
        End If
    Next
End Function

Private Function FindCheckBoxInPage(ByVal Page As MSForms.Page, ByVal key As String) As MSForms.CheckBox
    Dim c As MSForms.Control
    For Each c In Page.Controls
        If typeName(c) = "CheckBox" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindCheckBoxInPage = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindCheckBoxInPage = FindCheckBox(c, key)
            If Not FindCheckBoxInPage Is Nothing Then Exit Function
        End If
    Next
End Function

Private Function FindOption(ByVal parent As MSForms.Frame, ByVal key As String) As MSForms.OptionButton
    Dim c As MSForms.Control
    For Each c In parent.Controls
        If typeName(c) = "OptionButton" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindOption = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindOption = FindOption(c, key)
            If Not FindOption Is Nothing Then Exit Function
        ElseIf typeName(c) = "MultiPage" Then
            Dim mp As MSForms.MultiPage, pg As MSForms.Page
            Set mp = c
            For Each pg In mp.Pages
                Set FindOption = FindOptionInPage(pg, key)
                If Not FindOption Is Nothing Then Exit Function
            Next
        End If
    Next
End Function

Private Function FindOptionInPage(ByVal Page As MSForms.Page, ByVal key As String) As MSForms.OptionButton
    Dim c As MSForms.Control
    For Each c In Page.Controls
        If typeName(c) = "OptionButton" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindOptionInPage = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindOptionInPage = FindOption(c, key)
            If Not FindOptionInPage Is Nothing Then Exit Function
        End If
    Next
End Function

Private Function FindButton(ByVal parent As MSForms.Frame, ByVal key As String) As MSForms.CommandButton
    Dim c As MSForms.Control
    For Each c In parent.Controls
        If typeName(c) = "CommandButton" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindButton = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindButton = FindButton(c, key)
            If Not FindButton Is Nothing Then Exit Function
        ElseIf typeName(c) = "MultiPage" Then
            Dim mp As MSForms.MultiPage, pg As MSForms.Page
            Set mp = c
            For Each pg In mp.Pages
                Set FindButton = FindButtonInPage(pg, key)
                If Not FindButton Is Nothing Then Exit Function
            Next
        End If
    Next
End Function

Private Function FindButtonInPage(ByVal Page As MSForms.Page, ByVal key As String) As MSForms.CommandButton
    Dim c As MSForms.Control
    For Each c In Page.Controls
        If typeName(c) = "CommandButton" Then
            If StrComp(c.Tag, key, vbTextCompare) = 0 Or StrComp(c.Name, key, vbTextCompare) = 0 Then
                Set FindButtonInPage = c
                Exit Function
            End If
        ElseIf typeName(c) = "Frame" Then
            Set FindButtonInPage = FindButton(c, key)
            If Not FindButtonInPage Is Nothing Then Exit Function
        End If
    Next
End Function

'==================== UTILITIES ====================

Private Function IsLevel3Leaf(ByVal id As String) As Boolean
    Dim parts() As String: parts = Split(id, ".")
    If UBound(parts) <> 2 Then Exit Function
    If parts(0) Like "#*" And parts(1) Like "#*" And parts(2) Like "#*" Then
        If parts(2) Like "*[A-Za-z]*" Then
            IsLevel3Leaf = False
        Else
            IsLevel3Leaf = True
        End If
    End If
End Function

Private Sub HideIfExists(ByVal fra As MSForms.Frame, ByVal key As String)
    Dim c As Control
    Set c = FindCheckBox(fra, key): If Not c Is Nothing Then c.Visible = False: Exit Sub
    Set c = FindOption(fra, key):   If Not c Is Nothing Then c.Visible = False: Exit Sub
    Set c = FindTextBox(fra, key):  If Not c Is Nothing Then c.Visible = False: Exit Sub
    Dim lb As MSForms.Label: Set lb = FindLabel(fra, key)
    If Not lb Is Nothing Then lb.Visible = False
End Sub

'==================== CLEANUP ====================

Private Sub ClearAllRowsOnForm()
    On Error Resume Next
    Dim i As Long, nm As String
    If Not RowFrames Is Nothing Then
        For i = RowFrames.Count To 1 Step -1
            nm = RowFrames(i)
            If ControlExistsOnForm(nm) Then Me.Controls.Remove nm
        Next
        Set RowFrames = Nothing
    End If
    On Error GoTo 0
End Sub

Private Function ControlExistsOnForm(ByVal ctlName As String) As Boolean
    Dim c As MSForms.Control
    For Each c In Me.Controls
        If StrComp(c.Name, ctlName, vbTextCompare) = 0 Then
            ControlExistsOnForm = True
            Exit Function
        End If
    Next
End Function

