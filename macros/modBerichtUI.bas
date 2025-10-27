Attribute VB_Name = "modBerichtUI"

'=================  modBerichtUI.bas  =================
Option Explicit

'==================== UI CONSTANTS ====================

' Control names in RowTemplate
Public Const CN_LineNumber        As String = "LblLine"
Public Const CN_ReportItemID2     As String = "LblID2"      ' Parent context (1.2)
Public Const CN_ReportItemID3     As String = "LblID3"      ' Current item (1.2.1)
Public Const CN_AntwortKunde      As String = "LblAntwort"

Public Const CN_Bericht           As String = "ChkInclude"
Public Const CN_Opt1              As String = "Opt1"
Public Const CN_Opt2              As String = "Opt2"
Public Const CN_Opt3              As String = "Opt3"
Public Const CN_Opt4              As String = "Opt4"

Public Const CN_FestMaster        As String = "TxtFestMaster"
Public Const CN_LevelMaster       As String = "TxtLevelMaster"

Public Const CN_FestEdit          As String = "TxtFestEdit"
Public Const CN_LevelEdit         As String = "TxtLevelEdit"

Public Const CN_UseFestOverride   As String = "ChkUseFestOverride"
Public Const CN_UseLevelOverride  As String = "ChkUseLevelOverride"
Public Const CN_BtnPromote        As String = "BtnPromoteMaster"

' Layout constants
Public Const ROW_HEIGHT As Single = 180
Public Const ROW_GAP    As Single = 8
Public Const START_TOP  As Single = 50

' Pagination
Public Const ROWS_PER_PAGE As Long = 5

' Button colors
Public Const COLOR_SELECTED   As Long = 16737792  ' Orange RGB(255, 165, 0)
Public Const COLOR_DEFAULT    As Long = 15790320  ' Light gray RGB(240, 240, 240)

'==================== BUTTON MANAGEMENT ====================

Public Sub UpdateChapterButtons(ByVal frm As Object, ByVal currentChapter As Long)
    Dim i As Long
    Dim btn As MSForms.CommandButton
    
    For i = 1 To 11
        On Error Resume Next
        Set btn = frm.Controls("btnChap" & i)
        If Not btn Is Nothing Then
            If i = currentChapter Then
                btn.BackColor = COLOR_SELECTED
            Else
                btn.BackColor = COLOR_DEFAULT
            End If
        End If
        On Error GoTo 0
    Next i
End Sub

Public Sub UpdatePageButtons(ByVal frm As Object, ByVal currentPage As Long, ByVal totalPages As Long)
    Dim i As Long
    Dim btn As MSForms.CommandButton
    
    For i = 1 To 11
        On Error Resume Next
        Set btn = frm.Controls("btnPage" & i)
        If Not btn Is Nothing Then
            If i <= totalPages Then
                btn.Visible = True
                If i = currentPage Then
                    btn.BackColor = COLOR_SELECTED
                Else
                    btn.BackColor = COLOR_DEFAULT
                End If
            Else
                btn.Visible = False
            End If
        End If
        On Error GoTo 0
    Next i
End Sub

'==================== PARENT CONTEXT ====================

Public Function GetParentID(ByVal id As String) As String
    ' Extract parent ID from 1.2.3 -> 1.2
    Dim parts() As String
    parts = Split(id, ".")
    
    If UBound(parts) >= 1 Then
        GetParentID = parts(0) & "." & parts(1)
    Else
        GetParentID = ""
    End If
End Function

Public Function FormatParentContext(ByVal parentID As String, ByVal feststellung As String) As String
    ' Format: "1.2 - Führungsgrundsätze"
    If Len(parentID) > 0 And Len(feststellung) > 0 Then
        FormatParentContext = parentID & " - " & feststellung
    Else
        FormatParentContext = ""
    End If
End Function
