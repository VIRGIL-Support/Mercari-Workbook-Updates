Attribute VB_Name = "modPhotos"
Option Explicit

Public Const FIRST_PHOTO_DATA_COLUMN As Long = 54

Public Function SafePhotoFolderName(ByVal folderName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SafePhotoFolderName = Trim$(CStr(folderName))
    For i = LBound(invalidChars) To UBound(invalidChars)
        SafePhotoFolderName = Replace(SafePhotoFolderName, CStr(invalidChars(i)), "-")
    Next i
    If SafePhotoFolderName = "" Then SafePhotoFolderName = "Unassigned Item"
End Function

Public Function GetPhotoFileExtension(ByVal filePath As String) As String
    If InStrRev(filePath, ".") = 0 Then Exit Function
    GetPhotoFileExtension = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
End Function

Public Function IsSupportedPhotoFile(ByVal filePath As String) As Boolean
    Select Case GetPhotoFileExtension(filePath)
        Case "jpg", "jpeg", "png", "gif", "bmp", "webp", "heic", "heif", "tif", "tiff"
            IsSupportedPhotoFile = True
    End Select
End Function

Public Function PhotoFileExists(ByVal filePath As String) As Boolean
    On Error GoTo ExitFunction
    If Trim$(filePath) = "" Then Exit Function
    PhotoFileExists = (Dir(filePath) <> "")
ExitFunction:
End Function

Private Sub ConvertPhotoToJpg(ByVal sourcePath As String, ByVal targetPath As String)
    Dim imageFile As Object
    Dim imageProcess As Object
    Set imageFile = CreateObject("WIA.ImageFile")
    Set imageProcess = CreateObject("WIA.ImageProcess")
    imageFile.LoadFile sourcePath
    imageProcess.Filters.Add imageProcess.FilterInfos("Convert").FilterID
    imageProcess.Filters(1).Properties("FormatID").Value = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    Set imageFile = imageProcess.Apply(imageFile)
    If PhotoFileExists(targetPath) Then Kill targetPath
    imageFile.SaveFile targetPath
End Sub

Private Function CopyPhotoToItemFolder(ByVal sourcePath As String, ByVal targetPath As String) As String
    On Error GoTo CopyFallback
    Select Case LCase$(GetPhotoFileExtension(sourcePath))
        Case "jpg", "jpeg"
            If PhotoFileExists(targetPath) Then Kill targetPath
            FileCopy sourcePath, targetPath
        Case Else
            ConvertPhotoToJpg sourcePath, targetPath
    End Select
    CopyPhotoToItemFolder = targetPath
    Exit Function
CopyFallback:
    Dim ext As String
    ext = UCase$(GetPhotoFileExtension(sourcePath))
    MsgBox "Mercari only accepts JPG images. The file '" & Mid$(sourcePath, InStrRev(sourcePath, "\") + 1) & "' could not be automatically converted to JPG format." & vbCrLf & vbCrLf & _
           "Please manually convert your " & ext & " format image to JPG prior to adding it to this workbook.", _
           vbCritical, _
           "Image Conversion Failed"
    CopyPhotoToItemFolder = ""
End Function

Public Function GetItemPhotoFolderPath(ByVal inventoryRow As Long) As String
    Dim ws As Worksheet
    Dim photosRoot As String
    Dim itemNumber As String
    Dim itemName As String
    Dim folderPath As String
    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    photosRoot = GetSettingValue("PHOTOS_FOLDER")
    If photosRoot = "" Then photosRoot = ThisWorkbook.Path & "\1 READY TO LIST"
    CreateFolderIfMissing photosRoot
    itemNumber = SafePhotoFolderName(ws.Cells(inventoryRow, COL_ITEM_NUMBER).Value)
    itemName = SafePhotoFolderName(ws.Cells(inventoryRow, COL_ITEM_NAME).Value)
    If Len(itemName) > 50 Then itemName = Left$(itemName, 50)
    folderPath = photosRoot & "\" & itemNumber & " - " & itemName
    If Len(folderPath) > 180 Then folderPath = photosRoot & "\" & itemNumber
    GetItemPhotoFolderPath = folderPath
End Function

Private Function GetCopiedPhotoFileName(ByVal inventoryRow As Long, ByVal sourcePath As String) As String
    Dim ws As Worksheet
    Dim itemNumber As String
    Dim itemName As String
    Dim sourceName As String
    Dim sourceBase As String
    Dim extensionPart As String
    Dim fileNamePrefix As String
    Dim maxSourceBaseLength As Long
    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    itemNumber = SafePhotoFolderName(ws.Cells(inventoryRow, COL_ITEM_NUMBER).Value)
    itemName = SafePhotoFolderName(ws.Cells(inventoryRow, COL_ITEM_NAME).Value)
    sourceName = Mid$(sourcePath, InStrRev(sourcePath, "\") + 1)
    If InStrRev(sourceName, ".") > 0 Then
        sourceBase = Left$(sourceName, InStrRev(sourceName, ".") - 1)
    Else
        sourceBase = sourceName
    End If
    extensionPart = "jpg"
    fileNamePrefix = itemNumber & "_-_" & itemName & "_"
    maxSourceBaseLength = 150 - Len(fileNamePrefix) - Len(extensionPart) - 1
    If maxSourceBaseLength < 20 Then maxSourceBaseLength = 20
    sourceBase = Left$(SafePhotoFolderName(sourceBase), maxSourceBaseLength)
    GetCopiedPhotoFileName = fileNamePrefix & sourceBase & "." & extensionPart
End Function

Public Sub ClearSavedPhotoPaths(ByVal dataRow As Long)
    Dim i As Long
    For i = 0 To DEFAULT_MAX_PHOTOS - 1
        SaveDataValue dataRow, FIRST_PHOTO_DATA_COLUMN + i, ""
    Next i
End Sub

Public Sub SavePhotoSourcePaths(ByVal dataRow As Long, ByVal selectedPhotoPaths As Collection)
    Dim i As Long
    ClearSavedPhotoPaths dataRow
    If selectedPhotoPaths Is Nothing Then Exit Sub
    For i = 1 To selectedPhotoPaths.Count
        If i > DEFAULT_MAX_PHOTOS Then Exit For
        SaveDataValue dataRow, FIRST_PHOTO_DATA_COLUMN + i - 1, CStr(selectedPhotoPaths(i))
    Next i
End Sub

Public Function LoadPhotoSourcePaths(ByVal dataRow As Long) As Collection
    Dim paths As New Collection
    Dim i As Long
    Dim savedPath As String
    For i = 0 To DEFAULT_MAX_PHOTOS - 1
        savedPath = GetDataValue(dataRow, FIRST_PHOTO_DATA_COLUMN + i)
        If Trim$(savedPath) <> "" Then paths.Add savedPath
    Next i
    Set LoadPhotoSourcePaths = paths
End Function

Public Sub MoveRemovedPhotoToRemovedFolder(ByVal photoPath As String, Optional ByVal showConfirmation As Boolean = True)
    Dim parentFolder As String
    Dim removedFolder As String
    Dim fileName As String
    Dim targetPath As String
    If Trim$(photoPath) = "" Then Exit Sub
    If PhotoFileExists(photoPath) = False Then Exit Sub
    parentFolder = Left$(photoPath, InStrRev(photoPath, "\") - 1)
    removedFolder = parentFolder & "\REMOVED PHOTOS"
    CreateFolderIfMissing removedFolder
    fileName = Mid$(photoPath, InStrRev(photoPath, "\") + 1)
    targetPath = removedFolder & "\" & fileName
    If PhotoFileExists(targetPath) Then
        targetPath = removedFolder & "\" & Format(Now, "yyyymmdd_hhnnss") & " - " & fileName
    End If
    Name photoPath As targetPath
    If showConfirmation Then
        MsgBox _
            "Removed photo moved to the REMOVED PHOTOS folder in:" & vbCrLf & vbCrLf & _
            parentFolder, _
            vbInformation
    End If
End Sub

Private Function IsPathInFolder(ByVal filePath As String, ByVal folderPath As String) As Boolean
    If Trim$(filePath) = "" Or Trim$(folderPath) = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    IsPathInFolder = (LCase$(Left$(filePath, Len(folderPath))) = LCase$(folderPath))
End Function

Public Sub MoveRemovedPhotosToRemovedFolder(ByVal inventoryRow As Long, ByVal removedPhotoPaths As Collection)
    Dim i As Long
    Dim movedCount As Long
    Dim firstFolder As String
    Dim itemFolderPath As String
    If removedPhotoPaths Is Nothing Then Exit Sub
    itemFolderPath = GetItemPhotoFolderPath(inventoryRow)
    For i = 1 To removedPhotoPaths.Count
        If PhotoFileExists(CStr(removedPhotoPaths(i))) And IsPathInFolder(CStr(removedPhotoPaths(i)), itemFolderPath) Then
            If firstFolder = "" Then
                firstFolder = Left$(CStr(removedPhotoPaths(i)), InStrRev(CStr(removedPhotoPaths(i)), "\") - 1)
            End If
            MoveRemovedPhotoToRemovedFolder CStr(removedPhotoPaths(i)), False
            movedCount = movedCount + 1
        End If
    Next i
    If movedCount > 0 Then
        MsgBox _
            movedCount & " removed photo(s) were moved to the REMOVED PHOTOS folder in:" & vbCrLf & vbCrLf & _
            firstFolder, _
            vbInformation
    End If
End Sub

Public Function CopySelectedPhotosToItemFolder(ByVal inventoryRow As Long, ByVal selectedPhotoPaths As Collection) As Collection
    Dim copiedPaths As New Collection
    Dim folderPath As String
    Dim i As Long
    Dim sourcePath As String
    Dim targetPath As String
    Dim skippedPaths As String
    Dim copiedCount As Long
    Set CopySelectedPhotosToItemFolder = copiedPaths
    If selectedPhotoPaths Is Nothing Then Exit Function
    If selectedPhotoPaths.Count = 0 Then Exit Function
    folderPath = GetItemPhotoFolderPath(inventoryRow)
    CreateFolderIfMissing folderPath
    For i = 1 To selectedPhotoPaths.Count
        If i > DEFAULT_MAX_PHOTOS Then Exit For
        sourcePath = CStr(selectedPhotoPaths(i))
        If PhotoFileExists(sourcePath) Then
            If LCase$(Left$(sourcePath, Len(folderPath) + 1)) = LCase$(folderPath & "\") Then
                copiedPaths.Add sourcePath
            Else
                targetPath = folderPath & "\" & GetCopiedPhotoFileName(inventoryRow, sourcePath)
                If PhotoFileExists(targetPath) Then
                    If MsgBox( _
                        "A photo with this file name already exists:" & vbCrLf & vbCrLf & _
                        Mid$(targetPath, InStrRev(targetPath, "\") + 1) & vbCrLf & vbCrLf & _
                        "Choose Yes to overwrite the existing photo, or No to keep the existing photo.", _
                        vbQuestion + vbYesNo, _
                        "Duplicate Photo Found") = vbYes Then
                        Dim tempPath As String
                        tempPath = CopyPhotoToItemFolder(sourcePath, targetPath)
                        If tempPath <> "" Then
                            copiedPaths.Add tempPath
                            copiedCount = copiedCount + 1
                        End If
                    Else
                        copiedPaths.Add targetPath
                    End If
                Else
                    Dim newTempPath As String
                    newTempPath = CopyPhotoToItemFolder(sourcePath, targetPath)
                    If newTempPath <> "" Then
                        copiedPaths.Add newTempPath
                        copiedCount = copiedCount + 1
                    End If
                End If
            End If
        Else
            skippedPaths = skippedPaths & sourcePath & vbCrLf
            copiedPaths.Add sourcePath
        End If
    Next i
    If copiedPaths.Count = 0 Then
        MsgBox "No photos were copied. The selected photo list is empty or the selected files could not be found.", vbExclamation
    ElseIf skippedPaths <> "" Then
        MsgBox _
            "Some selected photos could not be found and were not copied:" & vbCrLf & vbCrLf & _
            skippedPaths & vbCrLf & _
            "Please re-add these photos from their current folder, then click Save and Close again.", _
            vbExclamation
    ElseIf copiedCount > 0 Then
        MsgBox copiedCount & " photo(s) copied to:" & vbCrLf & vbCrLf & folderPath, vbInformation
    End If
End Function

Public Sub OpenItemPhotoFolder(ByVal inventoryRow As Long)
    Dim folderPath As String
    folderPath = GetItemPhotoFolderPath(inventoryRow)
    CreateFolderIfMissing folderPath
    ThisWorkbook.FollowHyperlink folderPath
End Sub

Public Sub OpenAllPhotoFolders()
    Dim photosRoot As String
    photosRoot = GetSettingValue("PHOTOS_FOLDER")
    If photosRoot = "" Then photosRoot = ThisWorkbook.Path & "\1 READY TO LIST"
    CreateFolderIfMissing photosRoot
    ThisWorkbook.FollowHyperlink photosRoot
End Sub

Public Sub CreatePhotosTabControls(ByVal frm As frmItemEditor)
    Dim pg As MSForms.Page
    Dim pageItem As MSForms.Page
    Dim btnSelect As MSForms.CommandButton
    Dim btnRemoveAll As MSForms.CommandButton
    Dim lblInfo As MSForms.Label
    For Each pageItem In frm.mpItemTabs.Pages
        If UCase$(Trim$(pageItem.caption)) = "PHOTOS" Then
            Set pg = pageItem
            Exit For
        End If
    Next pageItem
    If pg Is Nothing Then Exit Sub
    Set lblInfo = pg.Controls.Add("Forms.Label.1", "lblPhotosInfo")
    With lblInfo
        .caption = "Select up to 12 photos. Photos are copied and renamed when Save and Close is clicked."
        .Left = 18
        .Top = 18
        .width = 760
        .height = 22
        .Font.Size = 10
    End With
    Set btnSelect = pg.Controls.Add("Forms.CommandButton.1", "btnSelectPhotos")
    With btnSelect
        .caption = "Select Photos"
        .Left = 18
        .Top = 48
        .width = 120
        .height = 26
    End With
    frm.RegisterPhotoButtonEvent btnSelect, "SELECT"
    Set btnRemoveAll = pg.Controls.Add("Forms.CommandButton.1", "btnRemoveAllPhotos")
    With btnRemoveAll
        .caption = "Remove All"
        .Left = 150
        .Top = 48
        .width = 120
        .height = 26
    End With
    frm.RegisterPhotoButtonEvent btnRemoveAll, "REMOVE_ALL"
    frm.RefreshPhotoTab
End Sub
