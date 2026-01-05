VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRowUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================  CRowUI.cls  =================
Option Explicit

' Intent: encapsulate one row's controls + events.
' Public members are bound by the form after cloning the template.

' References to controls (late-bound to MSForms)
Public WithEvents ChkInclude          As MSForms.CheckBox
Attribute ChkInclude.VB_VarHelpID = -1
Public WithEvents Opt1                As MSForms.OptionButton
Attribute Opt1.VB_VarHelpID = -1
Public WithEvents Opt2                As MSForms.OptionButton
Attribute Opt2.VB_VarHelpID = -1
Public WithEvents Opt3                As MSForms.OptionButton
Attribute Opt3.VB_VarHelpID = -1
Public WithEvents Opt4                As MSForms.OptionButton
Attribute Opt4.VB_VarHelpID = -1

Public WithEvents TxtFestEdit         As MSForms.TextBox
Attribute TxtFestEdit.VB_VarHelpID = -1
Public WithEvents TxtLevelEdit        As MSForms.TextBox
Attribute TxtLevelEdit.VB_VarHelpID = -1
Public WithEvents ChkUseFestOverride  As MSForms.CheckBox
Attribute ChkUseFestOverride.VB_VarHelpID = -1
Public WithEvents ChkUseLevelOverride As MSForms.CheckBox
Attribute ChkUseLevelOverride.VB_VarHelpID = -1
Public WithEvents BtnPromoteMaster    As MSForms.CommandButton ' optional
Attribute BtnPromoteMaster.VB_VarHelpID = -1

' Read-only previews
Public LblLine        As MSForms.Label
Public LblID          As MSForms.Label
Public LblAntwort     As MSForms.Label
Public txtFestMaster  As MSForms.TextBox   ' note: your original casing preserved
Public TxtLevelMaster As MSForms.TextBox

' Data
Public reportItemID   As String
Private mLevel        As Long              ' 1..4 current level
Private mFindingOverride As String
Private mFindingOverrideEnabled As Boolean
Private mLevelOverrideText(1 To 4) As String
Private mLevelOverrideEnabled(1 To 4) As Boolean

' ---------- Init ----------
Public Sub Init(ByVal id As String, ByVal antwort As String, ByVal selected As Boolean, _
                ByVal levelNum As Long, ByVal festMaster As String, ByVal levelMaster As String)

    reportItemID = id

    ' Labels
    If Not LblID Is Nothing Then LblID.caption = id
    If Not LblLine Is Nothing Then LblLine.caption = LblLine.caption ' caller sets it
    If Not LblAntwort Is Nothing Then LblAntwort.caption = antwort

    ' Include + level
    If Not ChkInclude Is Nothing Then ChkInclude.Value = selected
    mLevel = IIf(levelNum < 1 Or levelNum > 4, 1, levelNum)

    If Not Opt1 Is Nothing Then Opt1.Value = (mLevel = 1)
    If Not Opt2 Is Nothing Then Opt2.Value = (mLevel = 2)
    If Not Opt3 Is Nothing Then Opt3.Value = (mLevel = 3)
    If Not Opt4 Is Nothing Then Opt4.Value = (mLevel = 4)

    ' Master previews (left = Finding, right = Level preview)
    If Not txtFestMaster Is Nothing Then
        txtFestMaster.Locked = True
        txtFestMaster.Text = festMaster
    End If

    If Not TxtLevelMaster Is Nothing Then
        TxtLevelMaster.Locked = True
       TxtLevelMaster.Text = levelMaster
    End If

    ' Ensure right preview is consistent with current level/override
    RefreshLevelPreview
End Sub

Public Sub ConfigureOverrides(ByVal findingText As String, ByVal findingEnabled As Boolean, _
                              ByVal levelTexts As Variant, ByVal levelEnabled As Variant)
    mFindingOverride = findingText
    mFindingOverrideEnabled = findingEnabled
    If Not TxtFestEdit Is Nothing Then TxtFestEdit.Text = mFindingOverride
    If Not ChkUseFestOverride Is Nothing Then ChkUseFestOverride.Value = findingEnabled

    Dim i As Long
    For i = 1 To 4
        mLevelOverrideText(i) = ""
        mLevelOverrideEnabled(i) = False
        If IsArray(levelTexts) Then
            If ArrayHasIndex(levelTexts, i) Then
                mLevelOverrideText(i) = modABTableUtils.NzString(levelTexts(i))
            ElseIf ArrayHasIndex(levelTexts, i - 1) Then
                mLevelOverrideText(i) = modABTableUtils.NzString(levelTexts(i - 1))
            End If
        End If
        If IsArray(levelEnabled) Then
            If ArrayHasIndex(levelEnabled, i) Then
                mLevelOverrideEnabled(i) = modABTableUtils.NzBool(levelEnabled(i))
            ElseIf ArrayHasIndex(levelEnabled, i - 1) Then
                mLevelOverrideEnabled(i) = modABTableUtils.NzBool(levelEnabled(i - 1))
            End If
        End If
    Next i

    ApplyLevelOverrideState
End Sub

' ---------- Level selection ----------
Private Sub Opt1_Click(): HandleLevelClick 1: End Sub
Private Sub Opt2_Click(): HandleLevelClick 2: End Sub
Private Sub Opt3_Click(): HandleLevelClick 3: End Sub
Private Sub Opt4_Click(): HandleLevelClick 4: End Sub

Private Sub HandleLevelClick(ByVal newLevel As Long)
    mLevel = newLevel
    ' Selecting a level implies "include"
    If Not ChkInclude Is Nothing Then ChkInclude.Value = True
    modReportData.SaveSBState reportItemID, True, mLevel
    ApplyLevelOverrideState
    RefreshLevelPreview
End Sub

' ---------- Include toggle ----------
Private Sub ChkInclude_Click()
    modReportData.SaveSBState reportItemID, ChkInclude.Value, mLevel
End Sub

' ---------- Overrides ----------
Private Sub ChkUseFestOverride_Click()
    ' Save/clear the Finding override
    If ChkUseFestOverride.Value Then
        mFindingOverrideEnabled = True
        mFindingOverride = NzTB(TxtFestEdit)
        modReportData.SaveOverride reportItemID, ctFinding, 0, mFindingOverride, True
    Else
        mFindingOverrideEnabled = False
        mFindingOverride = ""
        modReportData.SaveOverride reportItemID, ctFinding, 0, "", False
    End If
End Sub

Private Sub ChkUseLevelOverride_Click()
    ' Save/clear the Level-N override (and refresh preview)
    If ChkUseLevelOverride.Value Then
        mLevelOverrideEnabled(mLevel) = True
        mLevelOverrideText(mLevel) = NzTB(TxtLevelEdit)
        modReportData.SaveOverride reportItemID, ctRecommendation, mLevel, mLevelOverrideText(mLevel), True
    Else
        mLevelOverrideEnabled(mLevel) = False
        mLevelOverrideText(mLevel) = ""
        modReportData.SaveOverride reportItemID, ctRecommendation, mLevel, "", False
    End If
    RefreshLevelPreview
End Sub

Private Sub TxtLevelEdit_Change()
    ' Live preview while typing when override is ON
    If Not ChkUseLevelOverride Is Nothing And ChkUseLevelOverride.Value Then
        mLevelOverrideText(mLevel) = NzTB(TxtLevelEdit)
        RefreshLevelPreview
    End If
End Sub

Private Sub TxtFestEdit_Change()
    If Not ChkUseFestOverride Is Nothing And ChkUseFestOverride.Value Then
        mFindingOverride = NzTB(TxtFestEdit)
    End If
End Sub

Private Sub BtnPromoteMaster_Click()
    MsgBox "Promote-to-Master not wired yet.", vbInformation
End Sub

' ---------- Preview recomputation ----------
Private Sub RefreshLevelPreview()
    If TxtLevelMaster Is Nothing Then Exit Sub

    Dim body As String
    If Not ChkUseLevelOverride Is Nothing And ChkUseLevelOverride.Value Then
        body = NzTB(TxtLevelEdit)
        If Len(Trim$(body)) = 0 Then
            ' No override typed yet ? fall back to Master LevelN so the preview isn't blank
            body = modReportData.GetMasterLevelText(reportItemID, mLevel)
        End If
    Else
        body = modReportData.GetMasterLevelText(reportItemID, mLevel)
    End If

    TxtLevelMaster.Locked = True
    TxtLevelMaster.Text = body
End Sub

' ---------- Small utility ----------
Private Sub ApplyLevelOverrideState()
    If mLevel < 1 Or mLevel > 4 Then Exit Sub
    If Not TxtLevelEdit Is Nothing Then TxtLevelEdit.Text = mLevelOverrideText(mLevel)
    If Not ChkUseLevelOverride Is Nothing Then ChkUseLevelOverride.Value = mLevelOverrideEnabled(mLevel)
End Sub

Private Function NzTB(tb As MSForms.TextBox) As String
    If tb Is Nothing Then
        NzTB = ""
    Else
        NzTB = tb.Text
    End If
End Function

Private Function ArrayHasIndex(ByVal source As Variant, ByVal index As Long) As Boolean
    On Error GoTo FAIL
    Dim lb As Long, ub As Long
    lb = LBound(source)
    ub = UBound(source)
    ArrayHasIndex = (index >= lb And index <= ub)
    Exit Function
FAIL:
    ArrayHasIndex = False
End Function

