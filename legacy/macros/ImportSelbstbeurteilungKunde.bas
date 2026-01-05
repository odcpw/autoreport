Attribute VB_Name = "ImportSelbstbeurteilungKunde"
Option Explicit

Public Sub btnSelbsbeurteilung_Click()
    Const SRC_SHEET_NAME As String = "Selbstbeurteilung Kunde"
    Dim filePath As String

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Waehle die Selbstbeurteilung-Datei"
        .Filters.Clear
        .Filters.Add "Excel Dateien", "*.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    On Error GoTo CleanFail
    Set sourceWb = Workbooks.Open(FileName:=filePath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(SRC_SHEET_NAME)

    Dim rowsWs As Worksheet
    Set rowsWs = modABRowsRepository.RowsSheet()

    Dim colRowId As Long
    Dim colAnswer As Long
    Dim colRemark As Long
    Dim colPriority As Long
    colRowId = HeaderIndex(rowsWs, "rowId")
    colAnswer = HeaderIndex(rowsWs, "customerAnswer")
    colRemark = HeaderIndex(rowsWs, "customerRemark")
    colPriority = HeaderIndex(rowsWs, "customerPriority")

    If colRowId = 0 Or colAnswer = 0 Or colRemark = 0 Or colPriority = 0 Then
        MsgBox "Rows sheet is missing required customer columns (rowId, customerAnswer, customerRemark, customerPriority).", vbCritical
        GoTo CleanFail
    End If

    Dim existingIndex As Scripting.Dictionary
    Set existingIndex = BuildRowIndex(rowsWs, colRowId)

    Dim inserted As Long, updated As Long
    inserted = 0: updated = 0

    Dim lastRow As Long
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then lastRow = 3

    Dim r As Long
    For r = 3 To lastRow
        Dim idRaw As String
        idRaw = Trim$(CStr(sourceWs.Cells(r, 1).Value))
        If Len(idRaw) = 0 Then GoTo ContinueRow

        Dim rowId As String
        rowId = NormalizeReportItemId(idRaw)
        If Not IsValidReportItemId(rowId) Then GoTo ContinueRow

        Dim existed As Boolean
        existed = existingIndex.Exists(rowId)

        Dim rowIndex As Long
        rowIndex = modABRowsRepository.EnsureRowRecord(rowId, ParentChapterId(rowId))
        If rowIndex = 0 Then GoTo ContinueRow

        rowsWs.Cells(rowIndex, colAnswer).Value = sourceWs.Cells(r, 4).Value   ' Antwort
        rowsWs.Cells(rowIndex, colRemark).Value = sourceWs.Cells(r, 5).Value   ' Bemerkung
        rowsWs.Cells(rowIndex, colPriority).Value = sourceWs.Cells(r, 6).Value ' Prioritaet

        If existed Then
            updated = updated + 1
        Else
            inserted = inserted + 1
            existingIndex(rowId) = rowIndex
        End If

ContinueRow:
    Next r

    MsgBox "Selbstbeurteilung import abgeschlossen." & vbCrLf & _
           "Neu angelegt: " & inserted & "   Aktualisiert: " & updated, vbInformation

CleanExit:
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

CleanFail:
    MsgBox "Fehler beim Import: " & Err.Description, vbCritical
    GoTo CleanExit
End Sub

Private Function BuildRowIndex(ws As Worksheet, colRowId As Long) As Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    map.CompareMode = TextCompare

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colRowId).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim id As String
        id = NormalizeReportItemId(ws.Cells(r, colRowId).Value)
        If Len(id) > 0 Then map(id) = r
    Next r

    Set BuildRowIndex = map
End Function
