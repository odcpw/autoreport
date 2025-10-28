Attribute VB_Name = "ImportMyMaster"
Option Explicit

' ================================================================
' Import MyMaster from DOCX into structured AutoBericht sheets
'   Required DOCX headers (exact tokens, case-insensitive):
'       ReportItemID | Feststellung | Level1 | Level2 | Level3 | Level4
'   Writes into sheet "Rows" (JSON-aligned layout)
' ================================================================

Private Type RowColumns
    rowId As Long
    chapterId As Long
    masterFinding As Long
    masterLevel(1 To 4) As Long
    overrideFinding As Long
    overrideLevel(1 To 4) As Long
    useOverrideFinding As Long
    useOverrideLevel(1 To 4) As Long
    includeFinding As Long
    includeRecommendation As Long
    selectedLevel As Long
    overwriteMode As Long
    done As Long
    notes As Long
    customerAnswer As Long
    customerRemark As Long
    customerPriority As Long
    lastEditedBy As Long
    lastEditedAt As Long
End Type

Public Sub ImportMyMaster()
    Const REMOVE_MISSING_ROWS As Boolean = False ' True removes Rows entries not present in DOCX

    Dim rowsWs As Worksheet
    Set rowsWs = modABRowsRepository.RowsSheet()

    Dim cols As RowColumns
    If Not ResolveRowColumns(rowsWs, cols) Then Exit Sub

    Dim rowIndexMap As Scripting.Dictionary
    Set rowIndexMap = BuildExistingRowIndex(rowsWs, cols)

    Dim importedIds As Scripting.Dictionary
    Set importedIds = New Scripting.Dictionary
    importedIds.CompareMode = TextCompare

    ' Pick DOCX
    Dim fd As FileDialog, docPath As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Master DOCX (MyMaster source)"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx;*.docm;*.doc"
        If .Show <> -1 Then Exit Sub
        docPath = .SelectedItems(1)
    End With

    ' Start Word (late binding)
    Dim wApp As Object, wDoc As Object
    On Error Resume Next
    Set wApp = CreateObject("Word.Application")
    On Error GoTo 0
    If wApp Is Nothing Then
        MsgBox "Could not start Microsoft Word.", vbCritical
        Exit Sub
    End If
    wApp.Visible = False

    Dim imported As Long, updated As Long
    imported = 0: updated = 0

    On Error GoTo CLEANUP
    Set wDoc = wApp.Documents.Open(FileName:=docPath, ReadOnly:=True, AddToRecentFiles:=False)

    Dim wTbl As Object
    For Each wTbl In wDoc.Tables
        Dim idxID&, idxFest&, idxL1&, idxL2&, idxL3&, idxL4&
        If DetectExactHeadersStrict(wTbl, idxID, idxFest, idxL1, idxL2, idxL3, idxL4) Then
            ImportTableStrict wTbl, idxID, idxFest, idxL1, idxL2, idxL3, idxL4, _
                              rowsWs, cols, rowIndexMap, importedIds, imported, updated
        End If
    Next wTbl

    If REMOVE_MISSING_ROWS Then
        DeleteRowsNotInSet rowsWs, cols, importedIds
    End If

    If imported = 0 And updated = 0 Then
        MsgBox "No compatible tables found in Word with exact headers on row 1." & vbCrLf & _
               "Open the Immediate Window (Ctrl+G) to inspect detected headers.", vbExclamation
    Else
        MsgBox "MyMaster import complete." & vbCrLf & _
               "Inserted: " & imported & "   Updated: " & updated, vbInformation
    End If

CLEANUP:
    On Error Resume Next
    If Not wDoc Is Nothing Then wDoc.Close SaveChanges:=False
    If Not wApp Is Nothing Then wApp.Quit
    Set wDoc = Nothing: Set wApp = Nothing
End Sub

' ---------- Structured sheet helpers ----------

Private Function ResolveRowColumns(ws As Worksheet, cols As RowColumns) As Boolean
    cols.rowId = HeaderIndex(ws, "rowId")
    cols.chapterId = HeaderIndex(ws, "chapterId")
    cols.masterFinding = HeaderIndex(ws, "masterFinding")
    cols.masterLevel(1) = HeaderIndex(ws, "masterLevel1")
    cols.masterLevel(2) = HeaderIndex(ws, "masterLevel2")
    cols.masterLevel(3) = HeaderIndex(ws, "masterLevel3")
    cols.masterLevel(4) = HeaderIndex(ws, "masterLevel4")
    cols.overrideFinding = HeaderIndex(ws, "overrideFinding")
    cols.useOverrideFinding = HeaderIndex(ws, "useOverrideFinding")
    cols.overrideLevel(1) = HeaderIndex(ws, "overrideLevel1")
    cols.overrideLevel(2) = HeaderIndex(ws, "overrideLevel2")
    cols.overrideLevel(3) = HeaderIndex(ws, "overrideLevel3")
    cols.overrideLevel(4) = HeaderIndex(ws, "overrideLevel4")
    cols.useOverrideLevel(1) = HeaderIndex(ws, "useOverrideLevel1")
    cols.useOverrideLevel(2) = HeaderIndex(ws, "useOverrideLevel2")
    cols.useOverrideLevel(3) = HeaderIndex(ws, "useOverrideLevel3")
    cols.useOverrideLevel(4) = HeaderIndex(ws, "useOverrideLevel4")
    cols.includeFinding = HeaderIndex(ws, "includeFinding")
    cols.includeRecommendation = HeaderIndex(ws, "includeRecommendation")
    cols.selectedLevel = HeaderIndex(ws, "selectedLevel")
    cols.overwriteMode = HeaderIndex(ws, "overwriteMode")
    cols.done = HeaderIndex(ws, "done")
    cols.notes = HeaderIndex(ws, "notes")
    cols.customerAnswer = HeaderIndex(ws, "customerAnswer")
    cols.customerRemark = HeaderIndex(ws, "customerRemark")
    cols.customerPriority = HeaderIndex(ws, "customerPriority")
    cols.lastEditedBy = HeaderIndex(ws, "lastEditedBy")
    cols.lastEditedAt = HeaderIndex(ws, "lastEditedAt")

    ResolveRowColumns = ValidateColumns(cols)
End Function

Private Function ValidateColumns(cols As RowColumns) As Boolean
    Dim missing As String
    missing = ""

    If cols.rowId = 0 Then missing = missing & "rowId" & vbCrLf
    If cols.chapterId = 0 Then missing = missing & "chapterId" & vbCrLf
    If cols.masterFinding = 0 Then missing = missing & "masterFinding" & vbCrLf

    Dim i As Long
    For i = 1 To 4
        If cols.masterLevel(i) = 0 Then missing = missing & "masterLevel" & CStr(i) & vbCrLf
        If cols.overrideLevel(i) = 0 Then missing = missing & "overrideLevel" & CStr(i) & vbCrLf
        If cols.useOverrideLevel(i) = 0 Then missing = missing & "useOverrideLevel" & CStr(i) & vbCrLf
    Next i

    If cols.overrideFinding = 0 Then missing = missing & "overrideFinding" & vbCrLf
    If cols.useOverrideFinding = 0 Then missing = missing & "useOverrideFinding" & vbCrLf
    If cols.includeFinding = 0 Then missing = missing & "includeFinding" & vbCrLf
    If cols.includeRecommendation = 0 Then missing = missing & "includeRecommendation" & vbCrLf
    If cols.selectedLevel = 0 Then missing = missing & "selectedLevel" & vbCrLf
    If cols.overwriteMode = 0 Then missing = missing & "overwriteMode" & vbCrLf
    If cols.done = 0 Then missing = missing & "done" & vbCrLf
    If cols.notes = 0 Then missing = missing & "notes" & vbCrLf
    If cols.customerAnswer = 0 Then missing = missing & "customerAnswer" & vbCrLf
    If cols.customerRemark = 0 Then missing = missing & "customerRemark" & vbCrLf
    If cols.customerPriority = 0 Then missing = missing & "customerPriority" & vbCrLf
    If cols.lastEditedBy = 0 Then missing = missing & "lastEditedBy" & vbCrLf
    If cols.lastEditedAt = 0 Then missing = missing & "lastEditedAt" & vbCrLf

    If Len(missing) > 0 Then
        MsgBox "Rows sheet is missing required columns:" & vbCrLf & missing, vbCritical
        ValidateColumns = False
    Else
        ValidateColumns = True
    End If
End Function

Private Function BuildExistingRowIndex(ws As Worksheet, cols As RowColumns) As Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    map.CompareMode = TextCompare

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, cols.rowId).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim key As String
        key = NormalizeReportItemId(ws.Cells(r, cols.rowId).Value)
        If Len(key) > 0 Then map(key) = r
    Next r
    Set BuildExistingRowIndex = map
End Function

Private Sub ImportTableStrict(ByVal tbl As Object, _
    ByVal idxID As Long, ByVal idxFest As Long, _
    ByVal idxL1 As Long, ByVal idxL2 As Long, ByVal idxL3 As Long, ByVal idxL4 As Long, _
    ByVal rowsWs As Worksheet, ByRef cols As RowColumns, _
    ByRef rowIndexMap As Scripting.Dictionary, ByRef importedIds As Scripting.Dictionary, _
    ByRef insertedOut As Long, ByRef updatedOut As Long)

    Dim lastR As Long
    lastR = tbl.Rows.Count
    If lastR < 2 Then Exit Sub

    Dim r As Long
    For r = 2 To lastR
        Dim idRaw As String
        idRaw = CleanHeader(GetCellText(tbl.Cell(r, idxID)))
        If Len(idRaw) = 0 Then GoTo ContinueRow

        Dim id As String
        id = NormalizeReportItemId(idRaw)
        If Not IsValidReportItemId(id) Then GoTo ContinueRow

        Dim fest As String
        Dim l1 As String, l2 As String, l3 As String, l4 As String
        fest = CleanBody(GetCellText(tbl.Cell(r, idxFest)))
        l1 = CleanBody(GetCellText(tbl.Cell(r, idxL1)))
        l2 = CleanBody(GetCellText(tbl.Cell(r, idxL2)))
        l3 = CleanBody(GetCellText(tbl.Cell(r, idxL3)))
        l4 = CleanBody(GetCellText(tbl.Cell(r, idxL4)))

        UpsertMasterRow rowsWs, cols, rowIndexMap, importedIds, id, fest, Array(l1, l2, l3, l4), insertedOut, updatedOut

ContinueRow:
    Next r
End Sub

Private Sub UpsertMasterRow(ByVal rowsWs As Worksheet, ByRef cols As RowColumns, _
    ByRef rowIndexMap As Scripting.Dictionary, ByRef importedIds As Scripting.Dictionary, _
    ByVal rowId As String, ByVal fest As String, ByVal levels As Variant, _
    ByRef insertedOut As Long, ByRef updatedOut As Long)

    Dim existed As Boolean
    existed = rowIndexMap.Exists(rowId)

    Dim rowIndex As Long
    rowIndex = modABRowsRepository.EnsureRowRecord(rowId, ParentChapterId(rowId))
    If rowIndex = 0 Then Exit Sub

    rowsWs.Cells(rowIndex, cols.masterFinding).Value = fest

    Dim i As Long
    For i = 1 To 4
        rowsWs.Cells(rowIndex, cols.masterLevel(i)).Value = NzLevel(levels, i)
    Next i

    importedIds(rowId) = True
    If existed Then
        updatedOut = updatedOut + 1
    Else
        rowIndexMap(rowId) = rowIndex
        insertedOut = insertedOut + 1
    End If
End Sub

Private Sub DeleteRowsNotInSet(ByVal ws As Worksheet, ByRef cols As RowColumns, _
    ByRef importedIds As Scripting.Dictionary)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, cols.rowId).End(xlUp).Row
    Dim r As Long
    For r = lastRow To ROW_HEADER_ROW + 1 Step -1
        Dim id As String
        id = NormalizeReportItemId(ws.Cells(r, cols.rowId).Value)
        If Len(id) = 0 Then
            ws.Rows(r).Delete
        ElseIf Not importedIds.Exists(id) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Private Function NzLevel(levels As Variant, position As Long) As String
    On Error GoTo CLEANUP
    NzLevel = ""
    If IsArray(levels) Then
        If LBound(levels) <= position - 1 And UBound(levels) >= position - 1 Then
            NzLevel = Trim$(CStr(levels(position - 1)))
        End If
    End If
    Exit Function
CLEANUP:
    NzLevel = ""
End Function

' ---------- Helpers (STRICT) ----------

Private Function DetectExactHeadersStrict(ByVal tbl As Object, _
    ByRef idxID As Long, ByRef idxFest As Long, _
    ByRef idxL1 As Long, ByRef idxL2 As Long, ByRef idxL3 As Long, ByRef idxL4 As Long) As Boolean

    Dim c As Long, hdr As String
    idxID = 0: idxFest = 0: idxL1 = 0: idxL2 = 0: idxL3 = 0: idxL4 = 0
    If tbl.Rows.Count = 0 Then Exit Function

    Dim hdrDebug As String: hdrDebug = ""

    For c = 1 To tbl.Columns.Count
        hdr = CleanHeader(GetCellText(tbl.Cell(1, c)))
        hdrDebug = hdrDebug & IIf(hdrDebug = "", "", " | ") & hdr
        Select Case True
            Case SameToken(hdr, "ReportItemID"): idxID = c
            Case SameToken(hdr, "Feststellung"): idxFest = c
            Case SameToken(hdr, "Level1"): idxL1 = c
            Case SameToken(hdr, "Level2"): idxL2 = c
            Case SameToken(hdr, "Level3"): idxL3 = c
            Case SameToken(hdr, "Level4"): idxL4 = c
        End Select
    Next c

    Debug.Print "Table headers (row1): "; hdrDebug

    DetectExactHeadersStrict = (idxID > 0 And idxFest > 0 And idxL1 > 0 And idxL2 > 0 And idxL3 > 0 And idxL4 > 0)
End Function

Private Function GetCellText(ByVal cellObj As Object) As String
    Dim s As String
    s = cellObj.Range.Text
    s = Replace$(s, Chr$(13), vbLf)
    s = Replace$(s, Chr$(7), "")
    GetCellText = s
End Function

Private Function CleanHeader(ByVal s As String) As String
    s = Replace$(s, vbCr, vbLf)
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, Chr$(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    CleanHeader = Trim$(s)
End Function

Private Function CleanBody(ByVal s As String) As String
    s = Replace$(s, vbCr, vbLf)
    CleanBody = Trim$(s)
End Function

Private Function SameToken(ByVal a As String, ByVal b As String) As Boolean
    SameToken = (LCase$(CleanHeader(a)) = LCase$(b))
End Function
