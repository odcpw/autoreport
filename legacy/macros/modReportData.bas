Attribute VB_Name = "modReportData"
Option Explicit

'=============================================================
' Cached accessors for structured Rows data (report editor)
'=============================================================

Public Enum ComponentType
    ctFinding = 1
    ctRecommendation = 2
End Enum

Public Type SBRow
    reportItemID As String
    antwort As String
    Bemerkung As String
    selected As Boolean
    SelectedLevel As Long
End Type

Private Type RowColumns
    rowId As Long
    chapterId As Long
    masterFinding As Long
    masterLevel(1 To 4) As Long
    overrideFinding As Long
    overrideLevel(1 To 4) As Long
    useOverrideFinding As Long
    useOverrideLevel(1 To 4) As Long
    customerAnswer As Long
    customerRemark As Long
    customerPriority As Long
    includeFinding As Long
    includeRecommendation As Long
    selectedLevel As Long
End Type

Private mCols As RowColumns
Private mColsReady As Boolean
Private dictRows As Scripting.Dictionary

Public Sub LoadAllCaches()
    Dim ws As Worksheet
    Set ws = modABRowsRepository.RowsSheet()
    If Not ResolveRowColumns(ws) Then Exit Sub

    Set dictRows = New Scripting.Dictionary
    dictRows.CompareMode = TextCompare

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, mCols.rowId).End(xlUp).Row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim id As String
        id = CleanID(ws.Cells(r, mCols.rowId).Value)
        If Len(id) = 0 Then GoTo ContinueRow
        dictRows(id) = ReadRowSnapshot(ws, r)
ContinueRow:
    Next r
End Sub

Public Function GetSB(reportItemID As String) As SBRow
    EnsureCache
    Dim id As String
    id = CleanID(reportItemID)

    Dim result As SBRow
    result.reportItemID = id

    If dictRows.Exists(id) Then
        Dim snap As Scripting.Dictionary
        Set snap = dictRows(id)
        result.antwort = CStr(NzVariant(snap("customerAnswer"), ""))
        result.Bemerkung = NzString(snap("customerRemark"))
        result.selected = NzBool(snap("includeFinding"))
        result.SelectedLevel = CLng(NzNumber(snap("selectedLevel"), 0))
    End If

    GetSB = result
End Function

Public Function GetMasterFinding(reportItemID As String) As String
    EnsureCache
    Dim id As String
    id = CleanID(reportItemID)
    If dictRows.Exists(id) Then
        GetMasterFinding = NzString(dictRows(id)("masterFinding"))
    Else
        GetMasterFinding = ""
    End If
End Function

Public Function GetMasterLevelText(reportItemID As String, levelNum As Long) As String
    EnsureCache
    Dim id As String
    id = CleanID(reportItemID)
    If levelNum < 1 Or levelNum > 4 Then Exit Function
    If dictRows.Exists(id) Then
        GetMasterLevelText = NzString(dictRows(id)("masterLevel" & CStr(levelNum)))
    Else
        GetMasterLevelText = ""
    End If
End Function

Public Function GetMasterFeststellungByID(reportItemID As String) As String
    GetMasterFeststellungByID = GetMasterFinding(reportItemID)
End Function

Public Function GetOverrideFindingText(reportItemID As String) As String
    EnsureCache
    Dim id As String
    id = CleanID(reportItemID)
    If dictRows.Exists(id) Then
        GetOverrideFindingText = NzString(dictRows(id)("overrideFinding"))
    Else
        GetOverrideFindingText = ""
    End If
End Function

Public Function IsOverrideFindingEnabled(reportItemID As String) As Boolean
    EnsureCache
    Dim id As String
    id = CleanID(reportItemID)
    If dictRows.Exists(id) Then
        IsOverrideFindingEnabled = NzBool(dictRows(id)("useOverrideFinding"))
    Else
        IsOverrideFindingEnabled = False
    End If
End Function

Public Function GetOverrideLevelText(reportItemID As String, levelNum As Long) As String
    EnsureCache
    If levelNum < 1 Or levelNum > 4 Then Exit Function
    Dim id As String
    id = CleanID(reportItemID)
    If dictRows.Exists(id) Then
        GetOverrideLevelText = NzString(dictRows(id)("overrideLevel" & CStr(levelNum)))
    Else
        GetOverrideLevelText = ""
    End If
End Function

Public Function IsOverrideLevelEnabled(reportItemID As String, levelNum As Long) As Boolean
    EnsureCache
    If levelNum < 1 Or levelNum > 4 Then Exit Function
    Dim id As String
    id = CleanID(reportItemID)
    If dictRows.Exists(id) Then
        IsOverrideLevelEnabled = NzBool(dictRows(id)("useOverrideLevel" & CStr(levelNum)))
    Else
        IsOverrideLevelEnabled = False
    End If
End Function

Public Sub SaveSBState(reportItemID As String, selected As Boolean, levelNum As Long)
    Dim id As String
    id = CleanID(reportItemID)
    If Len(id) = 0 Then Exit Sub

    modABRowsRepository.EnsureRowRecord id
    modABRowsRepository.SetIncludeFlags id, selected, selected
    modABRowsRepository.UpdateRowField id, "selectedLevel", levelNum
    RefreshRowSnapshot id
End Sub

Public Sub SaveOverride(reportItemID As String, comp As ComponentType, levelNum As Long, textBody As String, Optional enableOverride As Variant)
    Dim id As String
    id = CleanID(reportItemID)
    If Len(id) = 0 Then Exit Sub

    Dim enable As Boolean
    If IsMissing(enableOverride) Then
        enable = (Len(Trim$(textBody)) > 0)
    Else
        enable = NzBool(enableOverride)
    End If

    Select Case comp
        Case ctFinding
            If textBody = GetOverrideFindingText(id) And enable = IsOverrideFindingEnabled(id) Then Exit Sub
            modABRowsRepository.SetFindingOverride id, textBody, enable
        Case ctRecommendation
            If levelNum < 1 Or levelNum > 4 Then Exit Sub
            If textBody = GetOverrideLevelText(id, levelNum) And enable = IsOverrideLevelEnabled(id, levelNum) Then Exit Sub
            modABRowsRepository.SetRecommendationOverride id, levelNum, textBody, enable
        Case Else
            Exit Sub
    End Select

    RefreshRowSnapshot id
End Sub

Public Function CleanID(ByVal value As String) As String
    CleanID = NormalizeReportItemId(value)
End Function

' ---------- Internal helpers ----------

Private Sub EnsureCache()
    If dictRows Is Nothing Then LoadAllCaches
End Sub

Private Function ResolveRowColumns(ws As Worksheet) As Boolean
    mCols.rowId = HeaderIndex(ws, "rowId")
    mCols.chapterId = HeaderIndex(ws, "chapterId")
    mCols.masterFinding = HeaderIndex(ws, "masterFinding")
    mCols.masterLevel(1) = HeaderIndex(ws, "masterLevel1")
    mCols.masterLevel(2) = HeaderIndex(ws, "masterLevel2")
    mCols.masterLevel(3) = HeaderIndex(ws, "masterLevel3")
    mCols.masterLevel(4) = HeaderIndex(ws, "masterLevel4")
    mCols.overrideFinding = HeaderIndex(ws, "overrideFinding")
    mCols.useOverrideFinding = HeaderIndex(ws, "useOverrideFinding")
    mCols.overrideLevel(1) = HeaderIndex(ws, "overrideLevel1")
    mCols.overrideLevel(2) = HeaderIndex(ws, "overrideLevel2")
    mCols.overrideLevel(3) = HeaderIndex(ws, "overrideLevel3")
    mCols.overrideLevel(4) = HeaderIndex(ws, "overrideLevel4")
    mCols.useOverrideLevel(1) = HeaderIndex(ws, "useOverrideLevel1")
    mCols.useOverrideLevel(2) = HeaderIndex(ws, "useOverrideLevel2")
    mCols.useOverrideLevel(3) = HeaderIndex(ws, "useOverrideLevel3")
    mCols.useOverrideLevel(4) = HeaderIndex(ws, "useOverrideLevel4")
    mCols.customerAnswer = HeaderIndex(ws, "customerAnswer")
    mCols.customerRemark = HeaderIndex(ws, "customerRemark")
    mCols.customerPriority = HeaderIndex(ws, "customerPriority")
    mCols.includeFinding = HeaderIndex(ws, "includeFinding")
    mCols.includeRecommendation = HeaderIndex(ws, "includeRecommendation")
    mCols.selectedLevel = HeaderIndex(ws, "selectedLevel")

    ResolveRowColumns = ValidateColumns()
    mColsReady = ResolveRowColumns
End Function

Private Function ValidateColumns() As Boolean
    Dim missing As String
    missing = ""
    If mCols.rowId = 0 Then missing = missing & "rowId" & vbCrLf
    If mCols.masterFinding = 0 Then missing = missing & "masterFinding" & vbCrLf
    Dim i As Long
    For i = 1 To 4
        If mCols.masterLevel(i) = 0 Then missing = missing & "masterLevel" & CStr(i) & vbCrLf
        If mCols.overrideLevel(i) = 0 Then missing = missing & "overrideLevel" & CStr(i) & vbCrLf
        If mCols.useOverrideLevel(i) = 0 Then missing = missing & "useOverrideLevel" & CStr(i) & vbCrLf
    Next i
    If mCols.overrideFinding = 0 Then missing = missing & "overrideFinding" & vbCrLf
    If mCols.useOverrideFinding = 0 Then missing = missing & "useOverrideFinding" & vbCrLf
    If mCols.customerAnswer = 0 Then missing = missing & "customerAnswer" & vbCrLf
    If mCols.customerRemark = 0 Then missing = missing & "customerRemark" & vbCrLf
    If mCols.customerPriority = 0 Then missing = missing & "customerPriority" & vbCrLf
    If mCols.includeFinding = 0 Then missing = missing & "includeFinding" & vbCrLf
    If mCols.includeRecommendation = 0 Then missing = missing & "includeRecommendation" & vbCrLf
    If mCols.selectedLevel = 0 Then missing = missing & "selectedLevel" & vbCrLf

    If Len(missing) > 0 Then
        MsgBox "Rows sheet is missing required columns:" & vbCrLf & missing, vbCritical
        ValidateColumns = False
    Else
        ValidateColumns = True
    End If
End Function

Private Function ReadRowSnapshot(ws As Worksheet, rowIndex As Long) As Scripting.Dictionary
    Dim snap As New Scripting.Dictionary
    snap.CompareMode = TextCompare

    snap("rowId") = CleanID(ws.Cells(rowIndex, mCols.rowId).Value)
    snap("chapterId") = NzString(ws.Cells(rowIndex, mCols.chapterId).Value)
    snap("masterFinding") = NzString(ws.Cells(rowIndex, mCols.masterFinding).Value)

    Dim i As Long
    For i = 1 To 4
        snap("masterLevel" & CStr(i)) = NzString(ws.Cells(rowIndex, mCols.masterLevel(i)).Value)
        snap("overrideLevel" & CStr(i)) = NzString(ws.Cells(rowIndex, mCols.overrideLevel(i)).Value)
        snap("useOverrideLevel" & CStr(i)) = NzBool(ws.Cells(rowIndex, mCols.useOverrideLevel(i)).Value)
    Next i

    snap("overrideFinding") = NzString(ws.Cells(rowIndex, mCols.overrideFinding).Value)
    snap("useOverrideFinding") = NzBool(ws.Cells(rowIndex, mCols.useOverrideFinding).Value)
    snap("customerAnswer") = ws.Cells(rowIndex, mCols.customerAnswer).Value
    snap("customerRemark") = NzString(ws.Cells(rowIndex, mCols.customerRemark).Value)
    snap("customerPriority") = ws.Cells(rowIndex, mCols.customerPriority).Value
    snap("includeFinding") = NzBool(ws.Cells(rowIndex, mCols.includeFinding).Value)
    snap("includeRecommendation") = NzBool(ws.Cells(rowIndex, mCols.includeRecommendation).Value)
    snap("selectedLevel") = NzNumber(ws.Cells(rowIndex, mCols.selectedLevel).Value)

    Set ReadRowSnapshot = snap
End Function

Private Sub RefreshRowSnapshot(rowId As String)
    EnsureCache
    Dim ws As Worksheet
    Set ws = modABRowsRepository.RowsSheet()
    If Not mColsReady Then
        If Not ResolveRowColumns(ws) Then Exit Sub
    End If

    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, "rowId", rowId)
    If rowIndex = 0 Then
        If dictRows.Exists(rowId) Then dictRows.Remove rowId
    Else
        dictRows(rowId) = ReadRowSnapshot(ws, rowIndex)
    End If
End Sub

Private Function NzVariant(value As Variant, defaultValue As Variant) As Variant
    If IsMissing(value) Or IsEmpty(value) Then
        NzVariant = defaultValue
    ElseIf IsNull(value) Then
        NzVariant = defaultValue
    Else
        NzVariant = value
    End If
End Function
