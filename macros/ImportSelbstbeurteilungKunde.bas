Attribute VB_Name = "ImportSelbstbeurteilungKunde"
Sub btnSelbsbeurteilung_Click()
    Dim filePath As String
    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim i As Long
    Dim itemCode As String

    ' Let user pick the source file
    With Application.fileDialog(msoFileDialogFilePicker)
        .Title = "Wähle die Selbstbeurteilung-Datei"
        .Filters.Clear
        .Filters.Add "Excel Dateien", "*.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    ' Disable screen updating and alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set targetWb = ThisWorkbook
    Set targetWs = targetWb.Sheets("SelbstbeurteilungKunde")

    Set sourceWb = Workbooks.Open(FileName:=filePath, ReadOnly:=True)
    On Error GoTo CleanFail

    Set sourceWs = sourceWb.Sheets("Selbstbeurteilung Kunde")

    ' Copy columns D,E,F rows 3 to 706 if column A is not blank
    For i = 3 To 706
        itemCode = Trim(CStr(sourceWs.Cells(i, 1).Value)) ' Column A

        If itemCode <> "" Then
            targetWs.Cells(i, 4).Value = sourceWs.Cells(i, 4).Value  ' Column D
            targetWs.Cells(i, 5).Value = sourceWs.Cells(i, 5).Value  ' Column E
            targetWs.Cells(i, 6).Value = sourceWs.Cells(i, 6).Value  ' Column F
        End If
    Next i

    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Import abgeschlossen.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "Fehler beim Import: " & Err.description, vbCritical
    On Error Resume Next
    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub




