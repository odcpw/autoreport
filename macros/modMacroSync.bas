Attribute VB_Name = "modMacroSync"
Option Explicit

'=============================================================
' Centralised macro importer
'-------------------------------------------------------------
' • Requires: Tools ▸ References ▸ Microsoft Visual Basic for
'   Applications Extensibility 5.3
' • Trust Center ▸ Macro Settings ▸ Trust access to the VBA
'   project object model must be enabled.
' • Source folder must contain the exported .bas/.cls modules and
'   .frm/.frx userform pairs.
'=============================================================

Private Const DEFAULT_SOURCE_FOLDER As String = "C:\\Autobericht\\macros"

Public Sub RunRefreshMacros()
    RefreshProjectMacros
End Sub

Public Sub RefreshProjectMacros(Optional ByVal sourceFolder As String = "")
    Dim folderPath As String
    folderPath = ResolveSourceFolder(sourceFolder)
    If Len(folderPath) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisWorkbook.VBProject

    Dim preserved As Object
    Set preserved = BuildPreserveDictionary()

    RemoveExistingComponents vbProj, preserved, folderPath
    ImportFolderComponents vbProj, folderPath, "*.bas"
    ImportFolderComponents vbProj, folderPath, "*.cls"
    ImportFolderComponents vbProj, folderPath, "*.frm" ' imports any .frm/.frx pairs

    Application.ScreenUpdating = True
    MsgBox "Macros refreshed from " & folderPath, vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Macro refresh failed: " & Err.Description, vbCritical
End Sub

Private Function ResolveSourceFolder(ByVal overrideFolder As String) As String
    Dim candidate As String
    candidate = Trim$(overrideFolder)
    If Len(candidate) = 0 Then candidate = DEFAULT_SOURCE_FOLDER

    If Len(candidate) = 0 Then
        MsgBox "No source folder defined.", vbExclamation
        ResolveSourceFolder = ""
        Exit Function
    End If

    If Right$(candidate, 1) = "\" Or Right$(candidate, 1) = "/" Then
        candidate = Left$(candidate, Len(candidate) - 1)
    End If

    If Dir(candidate, vbDirectory) = vbNullString Then
        MsgBox "Source folder not found: " & candidate, vbCritical
        ResolveSourceFolder = ""
    Else
        ResolveSourceFolder = candidate
    End If
End Function

Private Function BuildPreserveDictionary() As Object
    Dim dict As Object:
    Set dict = CreateObject("Scripting.Dictionary")
    Dim items() As String
    items = Split(PRESERVE_COMPONENTS_LIST, "|")
    Dim entry As Variant
    For Each entry In items
        entry = Trim$(entry)
        If Len(entry) > 0 Then dict(entry) = True
    Next entry
    Set BuildPreserveDictionary = dict
End Function

Private Sub RemoveExistingComponents(ByVal vbProj As VBIDE.VBProject, ByVal preserved As Object, ByVal exportFolder As String)
    Dim components As New Collection
    Dim vbComp As VBIDE.VBComponent

    For Each vbComp In vbProj.VBComponents
        If ShouldRemoveComponent(vbComp, preserved) Then
            If exportFolder <> "" Then
                ExportComponent vbComp, exportFolder
            End If
            components.Add vbComp
        End If
    Next vbComp

    For Each vbComp In components
        vbProj.VBComponents.Remove vbComp
    Next vbComp
End Sub

Private Sub ExportComponent(ByVal vbComp As VBIDE.VBComponent, ByVal exportFolder As String)
    On Error GoTo CleanExit
    If vbComp.Type <> vbext_ct_MSForm Then Exit Sub
    If Len(exportFolder) = 0 Then Exit Sub
    If Dir(exportFolder, vbDirectory) = vbNullString Then
        On Error Resume Next
        MkDir exportFolder
        On Error GoTo CleanExit
    End If
    Dim targetPath As String
    targetPath = exportFolder & "\" & vbComp.Name & ".frm"
    vbComp.Export targetPath
CleanExit:
End Sub

Private Function ShouldRemoveComponent(ByVal vbComp As VBIDE.VBComponent, ByVal preserved As Object) As Boolean
    Select Case vbComp.Type
        Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
            If preserved.Exists(vbComp.Name) Then
                ShouldRemoveComponent = False
            Else
                ShouldRemoveComponent = True
            End If
        Case Else
            ShouldRemoveComponent = False ' worksheets / ThisWorkbook handled via preserve list
    End Select
End Function

Private Sub ImportFolderComponents(ByVal vbProj As VBIDE.VBProject, ByVal folderPath As String, ByVal pattern As String)
    Dim file As String
    file = Dir(folderPath & "\" & pattern)
    Do While Len(file) > 0
        vbProj.VBComponents.Import folderPath & "\" & file
        file = Dir
    Loop
End Sub
