VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhotoSorterForm 
   Caption         =   "PhotoSorter"
   ClientHeight    =   12255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23055
   OleObjectBlob   =   "PhotoSorterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PhotoSorterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim imageFiles As Collection
Dim imagesLoadedFlag As Boolean
Private currentIndex As Long
Private currentImageName As String
Private buttonCollection As Collection
Public existingButtonCaptions As Scripting.Dictionary

Private Sub UserForm_Initialize()
    PhotoSorterForm_Initialize
End Sub
Public Sub PhotoSorterForm_Initialize()
    Set imageFiles = New Collection
    Me.Top = 0
    InitializeSubDirectoryNames
    
    ' Initialize lblImageCounts
    lblImageCounts.caption = "Total Images: 0, Unsorted Images: 0"
    
    imagesLoadedFlag = False
    ' Populate cmbFilter
    With cmbFilter
        .AddItem "Show All"
        .AddItem "Show Only Unsorted"
        .AddItem "Show Only Sorted"
        .ListIndex = 0 ' Set default to "Show All"
    End With
End Sub

Private Sub cmbFilter_Change()
    If imagesLoadedFlag Then
        FilterAndReloadImages
    End If
    InvisibleTextBox.SetFocus
End Sub

Private Sub FilterAndReloadImages()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim unsortedCol As Range
    Dim filterOption As String
    Dim rowIndex As Long
    Dim imagePath As String
    Dim FileName As String
    Dim categoryCol As Range
    Dim isSorted As Boolean

    Set ws = ThisWorkbook.Sheets("PSHelperSheet")
    filterOption = cmbFilter.value
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Set unsortedCol = ws.Rows(1).Find(What:="Unsorted", LookIn:=xlValues, LookAt:=xlWhole)

    ' Clear the current imageFiles collection
    Set imageFiles = New Collection

    ' Debugging message
    Debug.Print "Filter option selected: " & filterOption

    ' Loop through the helper sheet to load images based on the filter option
    For rowIndex = 2 To lastRow ' Assuming the first row is headers
        FileName = ws.Cells(rowIndex, 1).value
        imagePath = rootPath & "\" & FileName
        
        If filterOption = "Show All" Then
            imageFiles.Add imagePath
        ElseIf filterOption = "Show Only Unsorted" Then
            If Not unsortedCol Is Nothing And ws.Cells(rowIndex, unsortedCol.column).value = 1 Then
                imageFiles.Add imagePath
                ' Debugging message
                Debug.Print "Unsorted image added: " & imagePath
            End If
        ElseIf filterOption = "Show Only Sorted" Then
            isSorted = False
            For Each categoryCol In ws.Rows(1).Columns
                If categoryCol.value <> "Unsorted" And ws.Cells(rowIndex, categoryCol.column).value = 1 Then
                    isSorted = True
                    Exit For
                End If
            Next categoryCol
            If isSorted Then
                imageFiles.Add imagePath
                ' Debugging message
                Debug.Print "Sorted image added: " & imagePath
            End If
        End If
    Next rowIndex

    ' Display the first image if available
    If imageFiles.count > 0 Then
        currentIndex = 1
        UpdateImageDisplay
    Else
        ' Clear the image display if no images match the filter
        currentIndex = 0
        If Not Me.ImageControl Is Nothing Then
            Me.ImageControl.Picture = Nothing
        End If
        lblCurrentImageName.caption = ""
    End If

    ' Update the counts
    UpdateImageCounts

    ' Debugging message
    Debug.Print "Total images loaded: " & imageFiles.count
End Sub


Private Sub btnCloseAndQuit_Click()
    Unload Me
End Sub

Private Sub cmdCreateAllFolders_Click()
    createAllFolders
End Sub

Private Sub cmdRemoveEmptyFolders_Click()
    RemoveEmptyFolders
End Sub

' === helper: release image handle ===
Private Sub ReleaseImageHandle()
    On Error Resume Next
    Set Me.ImageControl.Picture = Nothing
    Me.ImageControl.Picture = LoadPicture("")
    DoEvents
End Sub

' === helper: safe move ===
Private Sub MoveFileSafe(ByVal src As String, ByVal dst As String)
    Dim i As Long
    On Error GoTo CopyFallback
    Name src As dst
    Exit Sub
CopyFallback:
    Err.Clear
    For i = 1 To 5
        On Error Resume Next
        FileCopy src, dst
        If Err.Number = 0 Then
            Kill src
            Exit Sub
        End If
        Err.Clear
        DoEvents
        SleepMs 100 * i
    Next i
    MsgBox "Could not move " & src & " ? " & dst, vbExclamation, "MoveFileSafe"
End Sub

Private Sub SleepMs(ms As Long)
    Dim t As Single: t = Timer + ms / 1000#
    Do While Timer < t: DoEvents: Loop
End Sub

Private Sub cmdChooseDirectory_Click()
    ' Use a function to get the selected directory path
    Dim selectedPath As String
    selectedPath = chooseDirectory()

    If selectedPath <> "" Then
        ' Update the directory path label on the form
        lblDirectoryPath.caption = selectedPath
        
        ' Store the selected path in the public variable
        rootPath = selectedPath
        
        ' Parse directories, load images into sheet, and initialize imageFiles collection
        ParseAndLoadImages selectedPath
        
        ' Update subdirectory names in the BerichtLabels sheet
        UpdateSubDirectoryNames
        
        ' Recreate user form buttons based on new data
        CreateUserFormButtons
        
        ' Optional: Display the first image if available
        If imageFiles.count > 0 Then
            UpdateImageDisplay ' Assuming this subroutine handles image display logic
        End If
    Else
        ' Update label if no directory was selected
        lblDirectoryPath.caption = "No directory selected"
    End If
    
        ' Set the flag to indicate that images are loaded
    imagesLoadedFlag = True
    
        ' Call FilterAndReloadImages after loading images
    FilterAndReloadImages

    ' Assuming there's an invisible TextBox to manage focus
    UpdateImageCounts
    InvisibleTextBox.SetFocus
End Sub

' Code for the "Previous" button
Private Sub btnPrevious_Click()
    If currentIndex > 1 Then
        currentIndex = currentIndex - 1
        UpdateImageDisplay
        UpdateImageCounts
    End If
    InvisibleTextBox.SetFocus ' Set focus to the invisible TextBox
End Sub

' Code for the "Next" button
Private Sub btnNext_Click()
    If currentIndex < imageFiles.count Then
        currentIndex = currentIndex + 1
        UpdateImageDisplay
        UpdateImageCounts
    End If
    InvisibleTextBox.SetFocus ' Set focus to the invisible TextBox
End Sub

Private Sub InvisibleTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyA ' ASCII value for "A"
            btnPrevious_Click ' Simulate clicking the "Previous" button
        Case vbKeyD ' ASCII value for "D"
            btnNext_Click ' Simulate clicking the "Next" button
    End Select
End Sub
Private Sub UpdateImageCounts()
    Dim totalImages As Long
    Dim unsortedImages As Long
    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim lastRow As Long
    Dim unsortedCol As Range
    Dim currentImageIndex As Long
    Dim totalImageCount As Long

    Set ws = ThisWorkbook.Sheets("PSHelperSheet")
    totalImages = ws.Cells(ws.Rows.count, 1).End(xlUp).row - 1 ' Assuming the first row is headers
    unsortedImages = 0
    
    ' Find the column for "Unsorted"
    Set unsortedCol = ws.Rows(1).Find(What:="Unsorted", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not unsortedCol Is Nothing Then
        ' Count the number of unsorted images
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        For rowIndex = 2 To lastRow ' Start from row 2 to skip headers
            If ws.Cells(rowIndex, unsortedCol.column).value = 1 Then
                unsortedImages = unsortedImages + 1
            End If
        Next rowIndex
    End If
    
    currentImageIndex = currentIndex
    totalImageCount = imageFiles.count
    
    lblImageCounts.caption = "Image " & currentImageIndex & " of " & totalImageCount & _
                             " - Total Images " & totalImages & _
                             " - Unsorted Images " & unsortedImages
End Sub

Public Sub ParseAndLoadImages(ByVal rootDirectoryPath As String)
    Dim ws As Worksheet
    Dim labelsSheet As Worksheet
    Dim fso As Object
    Dim directoriesQueue As Collection
    Dim currentPath As String
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object
    Dim imageFilesDict As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim fullPath As String
    Dim FileName As String
    Dim categoryName As String
    Dim catCol As Object
    Dim colIndex As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set directoriesQueue = New Collection
    Set ws = ThisWorkbook.Sheets("PSHelperSheet")
    Set labelsSheet = ThisWorkbook.Sheets("PSCategoryLabels")
    Set imageFilesDict = CreateObject("Scripting.Dictionary")
    Set catCol = CreateObject("Scripting.Dictionary")

    ' Initialize
    directoriesQueue.Add rootDirectoryPath
    If subDirectoryNames Is Nothing Then InitializeSubDirectoryNames
    If Not subDirectoryNames.Exists("Unsorted") Then subDirectoryNames.Add "Unsorted", "Unsorted"

    ' Clear existing sheet contents
    ws.Range("A1:ZZ" & ws.Rows.count).ClearContents

    ' Load category labels from PSCategoryLabels
    LoadDynamicButtonCategories ws, catCol, "A"
    LoadDynamicButtonCategories ws, catCol, "B"
    LoadDynamicButtonCategories ws, catCol, "C"

    ' Loop through directories and files
    While directoriesQueue.count > 0
        currentPath = directoriesQueue(1)
        directoriesQueue.Remove 1
        Set folder = fso.GetFolder(currentPath)

        For Each file In folder.Files
            If LCase(file.Type) Like "*jpeg*" Or LCase(file.Type) Like "*jpg*" Or LCase(file.Type) Like "*png*" Then
                fullPath = file.Path
                FileName = Mid(fullPath, InStrRev(fullPath, "\") + 1)
                
                ' Determine the category based on the path
                If currentPath = rootDirectoryPath Then
                    categoryName = "Unsorted"
                Else
                    Dim categoryStart As Integer
                    categoryStart = InStrRev(fullPath, "\", InStrRev(fullPath, "\") - 1)
                    Dim categoryEnd As Integer
                    categoryEnd = InStrRev(fullPath, "\")
                    categoryName = Mid(fullPath, categoryStart + 1, categoryEnd - categoryStart - 1)
                End If

                ' Add category column if it does not exist
                If Not catCol.Exists(categoryName) Then
                    catCol.Add categoryName, catCol.count + 2
                    ws.Cells(1, catCol(categoryName)).value = categoryName
                End If

                ' Add file to the helper sheet
                If imageFilesDict.Exists(FileName) Then
                    rowIndex = imageFilesDict(FileName)
                    ws.Cells(rowIndex, catCol(categoryName)).value = "1"
                Else
                    rowIndex = ws.Cells(ws.Rows.count, "A").End(xlUp).row + 1
                    ws.Cells(rowIndex, 1).value = FileName
                    ws.Cells(rowIndex, catCol(categoryName)).value = "1"
                    imageFilesDict.Add FileName, rowIndex
                End If
            End If
        Next file

        ' Add subdirectories to the queue
        For Each subFolder In folder.SubFolders
            If Not subDirectoryNames.Exists(subFolder.Name) Then
                subDirectoryNames.Add subFolder.Name, subFolder.Name
            End If
            directoriesQueue.Add subFolder.Path
        Next subFolder
    Wend

    ' Find the last row and last column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column

    ' Sort the helper sheet by the filenames
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Columns(1), Order:=xlAscending
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .header = xlYes
        .Apply
    End With

    ' Initialize the imageFiles collection in the correct order
    Set imageFiles = New Collection
    For rowIndex = 2 To lastRow ' Assuming the first row is headers
        FileName = ws.Cells(rowIndex, 1).value
        Dim imagePath As String
        imagePath = ""

        ' Find the first category where the image exists
        For colIndex = 2 To lastCol
            If ws.Cells(rowIndex, colIndex).value = 1 Then
                categoryName = ws.Cells(1, colIndex).value
                If categoryName = "Unsorted" Then
                    imagePath = rootPath & "\" & FileName
                Else
                    imagePath = rootPath & "\" & categoryName & "\" & FileName
                End If
                If Dir(imagePath) <> "" Then
                    Exit For
                End If
            End If
        Next colIndex

        ' Default to root directory if no path was found
        If imagePath = "" Then
            imagePath = rootPath & "\" & FileName
        End If

        If Dir(imagePath) <> "" Then
            imageFiles.Add imagePath
        End If
    Next rowIndex
End Sub

Private Sub LoadDynamicButtonCategories(ws As Worksheet, ByRef catCol As Object, columnRange As String)
    Dim labelsSheet As Worksheet
    Set labelsSheet = ThisWorkbook.Sheets("PSCategoryLabels")
    
    Dim lastRow As Long
    Dim catName As String
    Dim i As Long
    lastRow = labelsSheet.Cells(labelsSheet.Rows.count, columnRange).End(xlUp).row
    For i = 1 To lastRow
        catName = Trim(labelsSheet.Cells(i, columnRange).value)
        If Not catCol.Exists(catName) Then
            catCol.Add catName, catCol.count + 2
            ws.Cells(1, catCol(catName)).value = catName
        End If
    Next i
End Sub

Private Sub UpdateImageDisplay()
    If currentIndex = 0 Then
        currentIndex = 1
    End If
    
    If imageFiles.count >= currentIndex Then
        Dim imageName As String
        imageName = imageFiles.item(currentIndex)
        
        ' Extract the image file name only (excluding path)
        imageName = Mid(imageName, InStrRev(imageName, "\") + 1)
        
        ' Find the row in PSHelperSheet corresponding to this imageName
        With ThisWorkbook.Sheets("PSHelperSheet")
            Dim found As Range
            Set found = .Columns(1).Find(What:=imageName, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then
                currentImageRow = found.row
                
                ' Initialize imagePath
                Dim imagePath As String
                imagePath = ""
                
                ' Check the columns to find the first category where the image exists
                Dim colIndex As Integer
                For colIndex = 2 To .Cells(1, .Columns.count).End(xlToLeft).column
                    If .Cells(currentImageRow, colIndex).value = 1 Then
                        Dim category As String
                        category = .Cells(1, colIndex).value
                        If category = "Unsorted" Then
                            imagePath = rootPath & "\" & imageName
                        Else
                            imagePath = rootPath & "\" & category & "\" & imageName
                        End If
                        ' Check if the file exists in this path
                        If Dir(imagePath) <> "" Then
                            Exit For
                        End If
                    End If
                Next colIndex
                
                ' If no path was found, default to root directory
                If imagePath = "" Then
                    imagePath = rootPath & "\" & imageName
                End If

                ' Display the image if it exists
                If Dir(imagePath) <> "" Then
                    With ImageControl
                        .Picture = LoadPicture(imagePath)
                        .PictureSizeMode = fmPictureSizeModeZoom  ' Maintain aspect ratio
                    End With
                Else
                    MsgBox "Image not found: " & imagePath
                End If
            Else
                currentImageRow = 0 ' Reset or handle error
                MsgBox "Image row not found in PSHelperSheet for: " & imageName
            End If
        End With
    End If
    updateButtonStates
End Sub

Private Sub CreateUserFormButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PSCategoryLabels")

    Const btnWidth As Integer = 80
    Const btnHeight As Integer = 20
    Const spaceBetween As Integer = 5
    Const offsetX As Integer = 5
    Const offsetY As Integer = 5
    Const maxRows As Integer = 5

    Set buttonCollection = New Collection
    Set existingButtonCaptions = New Scripting.Dictionary

    CreateButtonBlock ws, "A", Me.ButtonsBericht, offsetX, offsetY, btnWidth, btnHeight, spaceBetween, maxRows
    CreateButtonBlock ws, "B", Me.ButtonsVGSeminar, offsetX, offsetY, btnWidth, btnHeight, spaceBetween, maxRows
    CreateButtonBlock ws, "C", Me.ButtonsSubfolders, offsetX, offsetY, btnWidth, btnHeight, spaceBetween, maxRows
End Sub

Private Sub CreateButtonBlock(ws As Worksheet, column As String, targetFrame As MSForms.Frame, _
                              offsetX As Integer, offsetY As Integer, btnWidth As Integer, _
                              btnHeight As Integer, spaceBetween As Integer, maxRows As Integer)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, column).End(xlUp).row

    Dim i As Integer
    Dim posIndex As Integer
    Dim topPos As Integer
    topPos = offsetY ' Initialize top position for buttons
    
     For i = 1 To lastRow
        Dim caption As String
        caption = ws.Cells(i, column).value

        If Not existingButtonCaptions.Exists(caption) Then
            Dim dynamicButton As CButton
            Set dynamicButton = New CButton
            posIndex = i - 1

            ' Calculate top position dynamically based on the number of rows
            topPos = offsetY + (posIndex \ maxRows) * (btnHeight + spaceBetween)

            ' Create a new button and set properties
            With targetFrame.Controls.Add("Forms.CommandButton.1", "Button" & column & i, True)
                .caption = caption
                .Width = btnWidth
                .Height = btnHeight
                .Left = offsetX + (posIndex Mod maxRows) * (btnWidth + spaceBetween)
                .Top = topPos
            End With

            dynamicButton.categoryName = caption
            Set dynamicButton.btn = targetFrame.Controls("Button" & column & i)
            Set dynamicButton.ParentForm = Me
            buttonCollection.Add dynamicButton
            existingButtonCaptions.Add caption, True
        End If
    Next i

    ' Adjust ScrollHeight for the frame based on the position of the last button
    If lastRow > 0 Then
        targetFrame.ScrollBars = fmScrollBarsVertical
        targetFrame.ScrollHeight = topPos + btnHeight + spaceBetween
    End If
End Sub

Public Sub ToggleCategoryAssignment(ByVal category As String, ByVal ButtonCaption As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PSHelperSheet")

    ' Step 1: Get the current image name
    Dim currentImageName As String
    currentImageName = GetCurrentImageName()
    
    If currentImageName = "" Then
        MsgBox "Current image name is empty."
        Exit Sub
    End If
    
    ' Step 2: Find the row for the current image
    Dim foundRow As Range
    Set foundRow = ws.Columns(1).Find(What:=currentImageName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundRow Is Nothing Then
        MsgBox "Image row not found in PSHelperSheet."
        Exit Sub
    End If

    ' Step 3: Find the column for the category
    Dim categoryColumn As Range
    Set categoryColumn = ws.Rows(1).Find(What:=category, LookIn:=xlValues, LookAt:=xlWhole)
    
    If categoryColumn Is Nothing Then
        MsgBox "Category column not found in PSHelperSheet for '" & category & "'."
        Exit Sub
    End If
    
    ' Step 4: Toggle the selected category
    With ws.Cells(foundRow.row, categoryColumn.column)
        If .value = 1 Then
            .value = ""
        Else
            .value = 1
            ' If assigning a new category, unassign "Unsorted"
            Dim unsortedColumn As Range
            Set unsortedColumn = ws.Rows(1).Find(What:="Unsorted", LookIn:=xlValues, LookAt:=xlWhole)
            If Not unsortedColumn Is Nothing Then
                ws.Cells(foundRow.row, unsortedColumn.column).value = ""
            End If
        End If
    End With

    ' Step 5: If "Unsorted" is clicked, unassign all other categories
    If category = "Unsorted" Then
        Dim colIndex As Long
        For colIndex = 2 To ws.Cells(1, ws.Columns.count).End(xlToLeft).column
            If ws.Cells(1, colIndex).value <> "Unsorted" Then
                ws.Cells(foundRow.row, colIndex).value = ""
            End If
        Next colIndex
        ' Ensure "Unsorted" is assigned
        ws.Cells(foundRow.row, categoryColumn.column).value = 1
    End If

    ' Step 6: Check if no categories are assigned, then move to Unsorted (root directory)
    Dim noCategories As Boolean
    noCategories = True
    For colIndex = 2 To ws.Cells(1, ws.Columns.count).End(xlToLeft).column
        If ws.Cells(foundRow.row, colIndex).value = 1 Then
            noCategories = False
            Exit For
        End If
    Next colIndex

    If noCategories Then
        Set categoryColumn = ws.Rows(1).Find(What:="Unsorted", LookIn:=xlValues, LookAt:=xlWhole)
        If Not categoryColumn Is Nothing Then
            ws.Cells(foundRow.row, categoryColumn.column).value = 1
        End If
    End If

    ' Ensure the file system matches the PSHelperSheet state
    EnsureFileSystemMatchesSheet currentImageName

    updateButtonStates
    UpdateImageCounts
End Sub

Private Sub EnsureFileSystemMatchesSheet(imageName As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PSHelperSheet")

    Dim foundRow As Range
    Set foundRow = ws.Columns(1).Find(What:=imageName, LookIn:=xlValues, LookAt:=xlWhole)
    If foundRow Is Nothing Then Exit Sub

    Dim colIndex As Long
    Dim sourcePath As String, destinationPath As String
    Dim imageInUnsorted As Boolean: imageInUnsorted = False

    ' --- only release handle if file physically moves ---
    Dim currentFullPath As String
    currentFullPath = rootPath & "\" & imageName
    If Dir(currentFullPath) <> "" Then
        ' Only release if we are about to delete or move this exact file
        If ws.Rows(1).Find("Unsorted", LookIn:=xlValues, LookAt:=xlWhole) Is Nothing Then
            ReleaseImageHandle
        End If
    End If

    ' === create/move into checked categories ===
    For colIndex = 2 To ws.Cells(1, ws.Columns.count).End(xlToLeft).column
        Dim category As String
        category = ws.Cells(1, colIndex).value
        If ws.Cells(foundRow.row, colIndex).value = 1 Then
            If category = "Unsorted" Then
                destinationPath = rootPath & "\" & imageName
                imageInUnsorted = True
            Else
                destinationPath = rootPath & "\" & category & "\" & imageName
                If Not fso.FolderExists(rootPath & "\" & category) Then
                    fso.CreateFolder rootPath & "\" & category
                End If
            End If
            If Not fso.FileExists(destinationPath) Then
                sourcePath = rootPath & "\" & imageName
                If fso.FileExists(sourcePath) Then
                    MoveFileSafe sourcePath, destinationPath
                Else
                    Dim subFolder As Object
                    For Each subFolder In fso.GetFolder(rootPath).SubFolders
                        sourcePath = subFolder.Path & "\" & imageName
                        If fso.FileExists(sourcePath) Then
                            MoveFileSafe sourcePath, destinationPath
                            Exit For
                        End If
                    Next subFolder
                End If
            End If
        End If
    Next colIndex

    ' === If image not in Unsorted, remove from root ===
    If Not imageInUnsorted Then
        sourcePath = rootPath & "\" & imageName
        If fso.FileExists(sourcePath) Then
            On Error Resume Next
            fso.DeleteFile sourcePath, True
            On Error GoTo 0
        End If
    End If

    ' === Remove from unchecked categories ===
    Dim folder As Object
    For Each folder In fso.GetFolder(rootPath).SubFolders
        Dim folderCategory As Range
        Set folderCategory = ws.Rows(1).Find(What:=folder.Name, LookIn:=xlValues, LookAt:=xlWhole)
        If Not folderCategory Is Nothing Then
            If ws.Cells(foundRow.row, folderCategory.column).value <> 1 Then
                Dim fileToDelete As String
                fileToDelete = folder.Path & "\" & imageName
                If fso.FileExists(fileToDelete) Then
                    On Error Resume Next
                    fso.DeleteFile fileToDelete, True
                    On Error GoTo 0
                End If
                If fso.GetFolder(folder.Path).Files.count = 0 Then
                    On Error Resume Next
                    fso.DeleteFolder folder.Path, True
                    On Error GoTo 0
                End If
            End If
        End If
    Next folder

    ' === Refresh display so the image stays visible ===
    UpdateImageDisplay
End Sub


Public Sub updateButtonStates()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PSHelperSheet")
    Dim currentImageName As String
    currentImageName = GetCurrentImageName()

    If currentImageName = "" Then Exit Sub
    
    Dim foundRow As Range
    Set foundRow = ws.Columns(1).Find(currentImageName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundRow Is Nothing Then
        Dim btn As Object
        For Each btn In buttonCollection
            Dim btnObj As CButton
            Set btnObj = btn
            Dim categoryCol As Range
            Set categoryCol = ws.Rows(1).Find(btnObj.categoryName, LookIn:=xlValues, LookAt:=xlWhole)
            If Not categoryCol Is Nothing Then
                If ws.Cells(foundRow.row, categoryCol.column).value = 1 Then
                    btnObj.btn.ForeColor = vbRed
                Else
                    btnObj.btn.ForeColor = vbBlack
                End If
            End If
        Next btn
    End If
End Sub

Public Function GetCurrentImageName() As String
    If imageFiles.count >= currentIndex Then
        Dim imagePath As String
        imagePath = imageFiles.item(currentIndex)
        GetCurrentImageName = Mid(imagePath, InStrRev(imagePath, "\") + 1)
    Else
        GetCurrentImageName = ""
    End If
    lblCurrentImageName.caption = GetCurrentImageName
End Function


Private Sub btnClearSheets_Click()
    Dim helperSheet As Worksheet
    Dim berichtLabelsSheet As Worksheet
    
    Set helperSheet = ThisWorkbook.Sheets("PSHelperSheet")
    Set berichtLabelsSheet = ThisWorkbook.Sheets("PSCategoryLabels")
    
    helperSheet.Cells.ClearContents
    berichtLabelsSheet.Columns("D").ClearContents
        
    ' Clear the ImageControl
    If Not Me.ImageControl Is Nothing Then
        Me.ImageControl.Picture = Nothing
    End If
    
    ' Clear imagePath (assuming imagePath is a String variable)
    imagePath = ""
    
    ' Reinitialize imageFiles Collection
    Set imageFiles = New Collection
    
    ' Reset directory path label
    lblDirectoryPath.caption = "No directory chosen"
    lblCurrentImageName = ""
    
End Sub


Private Sub btnShowCounts_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    ShowCategoryCounts
End Sub

Private Sub btnShowCounts_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RestoreCategoryCaptions
    InvisibleTextBox.SetFocus
End Sub

Private Sub ShowCategoryCounts()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PSHelperSheet")
    
    Dim categoryCounts As Scripting.Dictionary
    Set categoryCounts = GetCategoryCounts(ws)
    
    Dim btn As Object
    For Each btn In buttonCollection
        Dim btnObj As CButton
        Set btnObj = btn
        Dim categoryCount As Long
        categoryCount = categoryCounts(btnObj.categoryName)
        btnObj.btn.caption = "(" & categoryCount & ") " & btnObj.categoryName
        ' Make the button caption bold if count is non-zero
        If categoryCount > 0 Then
            btnObj.btn.Font.Bold = True
        Else
            btnObj.btn.Font.Bold = False
        End If
    Next btn
End Sub

Private Sub RestoreCategoryCaptions()
    Dim btn As Object
    For Each btn In buttonCollection
        Dim btnObj As CButton
        Set btnObj = btn
        btnObj.btn.caption = btnObj.categoryName
        btnObj.btn.Font.Bold = False ' Reset the font to normal
    Next btn
End Sub

Private Function GetCategoryCounts(ws As Worksheet) As Scripting.Dictionary
    Dim categoryCounts As Scripting.Dictionary
    Set categoryCounts = New Scripting.Dictionary
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim categoryName As String
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).column
    
    ' Initialize category counts
    For colIndex = 2 To lastCol
        categoryName = ws.Cells(1, colIndex).value
        If Not categoryCounts.Exists(categoryName) Then
            categoryCounts.Add categoryName, 0
        End If
    Next colIndex
    
    ' Count the images in each category
    For rowIndex = 2 To lastRow ' Assuming the first row is headers
        For colIndex = 2 To lastCol
            If ws.Cells(rowIndex, colIndex).value = 1 Then
                categoryName = ws.Cells(1, colIndex).value
                categoryCounts(categoryName) = categoryCounts(categoryName) + 1
            End If
        Next colIndex
    Next rowIndex
    
    Set GetCategoryCounts = categoryCounts
End Function



