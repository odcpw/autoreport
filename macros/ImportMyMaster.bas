Attribute VB_Name = "ImportMyMaster"
Option Explicit

' ================================================================
' Import MyMaster from DOCX with STRICT headers (first row only)
'   Required headers (exact tokens, case-insensitive):
'     ReportItemID | Feststellung | Level1 | Level2 | Level3 | Level4
'   Sheet/Table required:
'     Sheet "MyMaster" with Table "TableMyMaster"
'     Columns: ReportItemID, Feststellung, Level1, Level2, Level3, Level4
' ================================================================
Public Sub ImportMyMaster()
    Const CLEAR_EXISTING As Boolean = True ' set False to append/update

    Dim ws As Worksheet, lo As ListObject
    Dim fd As fileDialog, docPath As String
    Dim wApp As Object, wDoc As Object, wTbl As Object
    Dim imported As Long, updated As Long

    ' Locate Excel table
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("MyMaster")
    Set lo = ws.ListObjects("TableMyMaster")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Excel table 'TableMyMaster' not found on sheet 'MyMaster'." & vbCrLf & _
               "Create it with columns: ReportItemID, Feststellung, Level1, Level2, Level3, Level4.", vbCritical
        Exit Sub
    End If
    If Not HeadersMatchExcel(lo) Then
        MsgBox "TableMyMaster columns must be exactly:" & vbCrLf & _
               "ReportItemID, Feststellung, Level1, Level2, Level3, Level4", vbCritical
        Exit Sub
    End If

    ' Pick DOCX
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Master DOCX (MyMaster source)"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx;*.docm;*.doc"
        If .Show <> -1 Then Exit Sub
        docPath = .SelectedItems(1)
    End With

    ' Start Word (late binding)
    On Error Resume Next
    Set wApp = CreateObject("Word.Application")
    On Error GoTo 0
    If wApp Is Nothing Then
        MsgBox "Could not start Microsoft Word.", vbCritical
        Exit Sub
    End If
    wApp.Visible = False

    On Error GoTo CLEANUP
    Set wDoc = wApp.Documents.Open(FileName:=docPath, ReadOnly:=True, AddToRecentFiles:=False)

    ' Build index of existing IDs if appending/updating
    Dim idRowIndex As Object: Set idRowIndex = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long, last As Long: last = lo.DataBodyRange.Rows.Count
        Dim k$: For r = 1 To last
            k = CStr(lo.DataBodyRange.Cells(r, 1).Value)
            If Len(k) > 0 Then idRowIndex(k) = r
        Next r
    End If

    ' Clear existing data if desired
    If CLEAR_EXISTING Then
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
        idRowIndex.RemoveAll
    End If

    ' Scan Word tables strictly
    Dim idxID&, idxFest&, idxL1&, idxL2&, idxL3&, idxL4&
    Dim insThis As Long, updThis As Long
    imported = 0: updated = 0

    For Each wTbl In wDoc.Tables
        If DetectExactHeadersStrict(wTbl, idxID, idxFest, idxL1, idxL2, idxL3, idxL4) Then
            ImportOneStrict wTbl, idxID, idxFest, idxL1, idxL2, idxL3, idxL4, lo, idRowIndex, insThis, updThis
            imported = imported + insThis
            updated = updated + updThis
        End If
    Next

    If imported = 0 And updated = 0 Then
        MsgBox "No compatible tables found in Word with exact headers on row 1." & vbCrLf & _
               "Open the Immediate Window (Ctrl+G) to see detected headers per table.", vbExclamation
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

' ---------- Helpers (STRICT) ----------

Private Function HeadersMatchExcel(lo As ListObject) As Boolean
    ' Expect exactly 6 columns in order:
    ' ReportItemID, Feststellung, Level1, Level2, Level3, Level4
    If lo.ListColumns.Count <> 6 Then Exit Function
    HeadersMatchExcel = ( _
        SameToken(lo.ListColumns(1).Name, "ReportItemID") And _
        SameToken(lo.ListColumns(2).Name, "Feststellung") And _
        SameToken(lo.ListColumns(3).Name, "Level1") And _
        SameToken(lo.ListColumns(4).Name, "Level2") And _
        SameToken(lo.ListColumns(5).Name, "Level3") And _
        SameToken(lo.ListColumns(6).Name, "Level4"))
End Function

Private Function DetectExactHeadersStrict(ByVal tbl As Object, _
    ByRef idxID As Long, ByRef idxFest As Long, _
    ByRef idxL1 As Long, ByRef idxL2 As Long, ByRef idxL3 As Long, ByRef idxL4 As Long) As Boolean

    Dim c As Long, hdr As String
    idxID = 0: idxFest = 0: idxL1 = 0: idxL2 = 0: idxL3 = 0: idxL4 = 0
    If tbl.Rows.Count = 0 Then Exit Function

    ' Print the headers we see (helps debugging if it fails)
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

Private Sub ImportOneStrict(ByVal tbl As Object, _
    ByVal idxID As Long, ByVal idxFest As Long, _
    ByVal idxL1 As Long, ByVal idxL2 As Long, ByVal idxL3 As Long, ByVal idxL4 As Long, _
    ByVal lo As ListObject, ByVal idRowIndex As Object, _
    ByRef insertedOut As Long, ByRef updatedOut As Long)

    Dim r As Long, lastR As Long
    Dim idRaw$, id$, fest$, l1$, l2$, l3$, l4$
    insertedOut = 0: updatedOut = 0
    lastR = tbl.Rows.Count
    If lastR < 2 Then Exit Sub ' no data

    For r = 2 To lastR
        idRaw = CleanHeader(GetCellText(tbl.Cell(r, idxID))) ' reuse CleanHeader to strip NBSP/CR/chr(7)
        If Len(idRaw) = 0 Then
            ' Layout-only row ? skip
        Else
            id = NormalizeReportItemID(idRaw)  ' strip trailing dots/spaces, collapse ..
            If IsValidReportItemID_Lax(id) Then
                ' Keep row even if texts are blank
                fest = CleanBody(GetCellText(tbl.Cell(r, idxFest)))
                l1 = CleanBody(GetCellText(tbl.Cell(r, idxL1)))
                l2 = CleanBody(GetCellText(tbl.Cell(r, idxL2)))
                l3 = CleanBody(GetCellText(tbl.Cell(r, idxL3)))
                l4 = CleanBody(GetCellText(tbl.Cell(r, idxL4)))

                If idRowIndex.Exists(id) Then
                    Dim erow&: erow = idRowIndex(id)
                    With lo.DataBodyRange.Rows(erow)
                        .Cells(1, 1).Value = id
                        .Cells(1, 2).Value = fest
                        .Cells(1, 3).Value = l1
                        .Cells(1, 4).Value = l2
                        .Cells(1, 5).Value = l3
                        .Cells(1, 6).Value = l4
                    End With
                    updatedOut = updatedOut + 1
                Else
                    Dim newRow As ListRow
                    Set newRow = lo.ListRows.Add
                    With newRow.Range
                        .Cells(1, 1).Value = id
                        .Cells(1, 2).Value = fest
                        .Cells(1, 3).Value = l1
                        .Cells(1, 4).Value = l2
                        .Cells(1, 5).Value = l3
                        .Cells(1, 6).Value = l4
                    End With
                    idRowIndex(id) = newRow.Index
                    insertedOut = insertedOut + 1
                End If
            Else
                ' Not a valid ID (e.g., "Kapitel") ? skip
            End If
        End If
    Next r
End Sub

' -------- text/ID helpers --------

Private Function GetCellText(ByVal cellObj As Object) As String
    ' Word table cell text ends with Chr(13) & Chr(7)
    Dim s$: s = cellObj.Range.Text
    s = Replace$(s, Chr$(13), vbLf)
    s = Replace$(s, Chr$(7), "")
    GetCellText = s
End Function

Private Function CleanHeader(ByVal s As String) As String
    ' Strict header cleaner: remove control chars and NBSP, trim spaces; do NOT alter words
    s = Replace$(s, vbCr, vbLf)
    s = Replace$(s, vbLf, " ")
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, Chr$(160), " ") ' NBSP
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    CleanHeader = Trim$(s)
End Function

Private Function SameToken(ByVal a As String, ByVal b As String) As Boolean
    SameToken = (LCase$(CleanHeader(a)) = LCase$(b))
End Function

Private Function CleanBody(ByVal s As String) As String
    ' Keep content; normalize line breaks; trim
    s = Replace$(s, vbCr, vbLf)
    s = Trim$(s)
    CleanBody = s
End Function

Private Function NormalizeReportItemID(ByVal s As String) As String
    ' Canonicalize: strip trailing dots/spaces, collapse double dots
    Dim t$: t = Trim$(s)
    Do While Len(t) > 0 And (Right$(t, 1) = "." Or Right$(t, 1) = " ")
        t = Left$(t, Len(t) - 1)
    Loop
    Do While InStr(t, "..") > 0
        t = Replace$(t, "..", ".")
    Loop
    NormalizeReportItemID = t
End Function

Private Function IsValidReportItemID_Lax(ByVal s As String) As Boolean
    ' Accept IDs with >=1 segments: 1, 1.1, 1.1.1, 2.3.4.5, and with optional a/b suffixes
    Dim parts() As String, i&, seg$
    If Len(s) = 0 Then Exit Function
    parts = Split(s, ".")
    If UBound(parts) < 0 Then Exit Function ' need at least "1"
    For i = LBound(parts) To UBound(parts)
        seg = parts(i)
        If Len(seg) = 0 Then Exit Function
        ' allow digits with optional trailing letters (e.g., "1", "1a")
        ' reject blatant non-IDs like "Kapitel"
        If Not seg Like "*[0-9]*" Then Exit Function
    Next i
    IsValidReportItemID_Lax = True
End Function


