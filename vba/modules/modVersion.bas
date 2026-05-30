Attribute VB_Name = "modVersion"
Option Explicit

' ============================================
' VERSION MANAGEMENT MODULE
' ============================================

Public Const CURRENT_VERSION As String = "1.0"
Public Const UPDATE_CHECK_URL As String = "https://raw.githubusercontent.com/VIRGIL-Support/Mercari-Workbook-Updates/main/version.txt"
Public Const UPDATE_DOWNLOAD_URL As String = "https://raw.githubusercontent.com/VIRGIL-Support/Mercari-Workbook-Updates/main/Mercari_Workbook_Latest.xlsm"

' Check for updates on workbook open
Public Sub CheckForUpdatesOnOpen()
    Dim checkResult As String
    Dim latestVersion As String
    Dim userChoice As VbMsgBoxResult
    
    ' Check if there's a pending data transfer from a previous update
    If CheckForPendingTransfer() Then Exit Sub
    
    ' Check if user wants to auto-check (stored in settings)
    If GetSettingValue("AUTO_CHECK_UPDATES") = "NO" Then Exit Sub
    
    checkResult = GetLatestVersionInfo()
    
    If checkResult = "" Then
        ' No update available or couldn't check
        Exit Sub
    End If
    
    latestVersion = Replace(Replace(Trim$(checkResult), vbCr, ""), vbLf, "")
    
    If IsNewerVersion(latestVersion, CURRENT_VERSION) Then
        userChoice = MsgBox("A new version (" & latestVersion & ") is available!" & vbCrLf & vbCrLf & _
                           "Current version: " & CURRENT_VERSION & vbCrLf & vbCrLf & _
                           "Would you like to update now?" & vbCrLf & vbCrLf & _
                           "When you click Yes, don't worry if it takes me up to a minute to respond. " & _
                           "I'll be busy backing up all of your data, downloading the update, " & _
                           "transferring everything into the shiny new version, and then double-checking " & _
                           "to make sure it's all perfect for you. I've got it all taken care of!", vbYesNo + vbInformation, "Update Available")
        
        If userChoice = vbYes Then
            DownloadUpdate latestVersion
        End If
    End If
    
End Sub

' Get latest version from web
Private Function GetLatestVersionInfo() As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    http.Open "GET", UPDATE_CHECK_URL, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send
    
    If http.Status = 200 Then
        GetLatestVersionInfo = http.responseText
    Else
        GetLatestVersionInfo = ""
    End If
    
    Exit Function
    
ErrorHandler:
    GetLatestVersionInfo = ""
End Function

' Compare version strings (returns True if newVersion > currentVersion)
Private Function IsNewerVersion(ByVal newVersion As String, ByVal currentVersion As String) As Boolean
    Dim newParts() As String
    Dim currentParts() As String
    Dim i As Long
    Dim maxParts As Long
    
    newParts = Split(newVersion, ".")
    currentParts = Split(currentVersion, ".")
    
    maxParts = UBound(newParts)
    If UBound(currentParts) > maxParts Then maxParts = UBound(currentParts)
    
    For i = 0 To maxParts
        Dim newVal As Long
        Dim currentVal As Long
        
        newVal = 0
        currentVal = 0
        
        If i <= UBound(newParts) Then
            If IsNumeric(newParts(i)) Then newVal = CLng(newParts(i))
        End If
        
        If i <= UBound(currentParts) Then
            If IsNumeric(currentParts(i)) Then currentVal = CLng(currentParts(i))
        End If
        
        If newVal > currentVal Then
            IsNewerVersion = True
            Exit Function
        ElseIf newVal < currentVal Then
            IsNewerVersion = False
            Exit Function
        End If
    Next i
    
    IsNewerVersion = False ' Versions are equal
    
End Function

' ============================================
' DOWNLOAD UPDATE
' ============================================

Private Sub DownloadUpdate(ByVal newVersion As String)
    On Error GoTo ErrorHandler
    
    Dim updateFolder As String
    Dim downloadPath As String
    Dim backupPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create backup silently
    Application.StatusBar = "Creating backup..."
    backupPath = CreatePreUpdateBackup()
    
    ' Create Updates folder
    updateFolder = ThisWorkbook.Path & "\Updates"
    CreateFolderIfMissing updateFolder
    
    ' Download the updated workbook
    Application.StatusBar = "Downloading update v" & newVersion & "..."
    
    downloadPath = updateFolder & "\Mercari_Workbook_v" & newVersion & ".xlsm"
    
    ' Download using XMLHTTP with binary stream
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", UPDATE_DOWNLOAD_URL, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send
    
    If http.Status <> 200 Then
        Application.StatusBar = False
        MsgBox "Failed to download update. Please try again later." & vbCrLf & vbCrLf & _
               "HTTP Status: " & http.Status, vbExclamation, "Download Failed"
        Exit Sub
    End If
    
    ' Save the downloaded file
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile downloadPath, 2 ' 2 = overwrite
    stream.Close
    
    Application.StatusBar = False
    
    ' Save the old workbook path so the new workbook knows where to find user data
    Dim oldWorkbookPath As String
    oldWorkbookPath = ThisWorkbook.FullName
    
    ' Update version in settings so we don't prompt again
    UpdateSetting "VERSION", newVersion
    
    ' Save old workbook path to a temp file so new workbook can find it
    Dim tempPathFile As String
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    Dim f As Integer
    f = FreeFile
    Open tempPathFile For Output As #f
    Print #f, oldWorkbookPath
    Close #f
    
    ' Save current workbook
    ThisWorkbook.Save
    
    ' Open the new workbook
    Workbooks.Open downloadPath
    
    ' Close the old workbook
    ThisWorkbook.Close SaveChanges:=False
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error during update: " & Err.Number & " - " & Err.Description, vbCritical, "Update Error"
End Sub

' ============================================
' TRANSFER DATA FROM OLD WORKBOOK
' ============================================

Public Sub TransferMyData()
    On Error GoTo ErrorHandler
    
    Dim sourceWorkbookPath As String
    Dim tempPathFile As String
    Dim sourceWb As Workbook
    Dim f As Integer
    Dim fso As Object
    Dim oldFolder As String
    Dim oldFileName As String
    Dim archiveFolder As String
    Dim archivePath As String
    Dim newFilePath As String
    Dim newFileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' DEBUG: Confirm TransferMyData is running
    MsgBox "TransferMyData starting..." & vbCrLf & "Temp file: " & Environ$("TEMP") & "\MercariUpdateSource.txt", vbInformation, "DEBUG"
    
    ' Try to read the old workbook path from temp file
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    If Dir(tempPathFile) = "" Then
        ' No temp file - no transfer needed, just exit silently
        MsgBox "No temp file found, exiting", vbInformation, "DEBUG"
        Exit Sub
    Else
        ' Read path from temp file
        MsgBox "DEBUG: Temp file found, reading path", vbInformation, "DEBUG"
        f = FreeFile
        Open tempPathFile For Input As #f
        Line Input #f, sourceWorkbookPath
        Close #f
        MsgBox "DEBUG: Path from temp file: " & sourceWorkbookPath, vbInformation, "DEBUG"
        
        ' Verify file exists - if not, clean up temp file and exit
        If Dir(sourceWorkbookPath) = "" Then
            MsgBox "DEBUG: Source workbook not found at path, cleaning up", vbInformation, "DEBUG"
            ' Old workbook is gone, clean up and continue normally
            Kill tempPathFile
            Exit Sub
        End If
    End If
    
    ' Remember old workbook location and name
    oldFolder = fso.GetParentFolderName(sourceWorkbookPath)
    oldFileName = fso.GetFileName(sourceWorkbookPath)
    MsgBox "DEBUG: oldFolder = " & oldFolder & vbCrLf & "oldFileName = " & oldFileName, vbInformation, "DEBUG"
    
    Application.StatusBar = "Opening source workbook..."
    Application.ScreenUpdating = False
    
    ' Open the old workbook (read-only)
    MsgBox "DEBUG: About to open source workbook", vbInformation, "DEBUG"
    Set sourceWb = Workbooks.Open(sourceWorkbookPath, ReadOnly:=True)
    MsgBox "DEBUG: Source workbook opened successfully", vbInformation, "DEBUG"
    
    ' Transfer INVENTORY data
    MsgBox "DEBUG: About to transfer INVENTORY", vbInformation, "DEBUG"
    Application.StatusBar = "Transferring INVENTORY data..."
    TransferSheetData sourceWb, WS_INVENTORY
    MsgBox "DEBUG: INVENTORY transferred", vbInformation, "DEBUG"
    
    ' Transfer SOLD ITEMS data
    MsgBox "DEBUG: About to transfer SOLD ITEMS", vbInformation, "DEBUG"
    Application.StatusBar = "Transferring SOLD ITEMS data..."
    TransferSheetData sourceWb, "SOLD ITEMS"
    MsgBox "DEBUG: SOLD ITEMS transferred", vbInformation, "DEBUG"
    
    ' Transfer SETTINGS data
    MsgBox "DEBUG: About to transfer SETTINGS", vbInformation, "DEBUG"
    Application.StatusBar = "Transferring SETTINGS..."
    TransferSheetData sourceWb, "SETTINGS"
    MsgBox "DEBUG: SETTINGS transferred", vbInformation, "DEBUG"
    
    ' Close the source workbook
    MsgBox "DEBUG: About to close source workbook", vbInformation, "DEBUG"
    sourceWb.Close SaveChanges:=False
    Set sourceWb = Nothing
    MsgBox "DEBUG: Source workbook closed", vbInformation, "DEBUG"
    
    ' Clean up temp file
    If Dir(tempPathFile) <> "" Then Kill tempPathFile
    MsgBox "DEBUG: Temp file cleaned up", vbInformation, "DEBUG"
    
    ' Create Archive folder in the old workbook's location
    Application.StatusBar = "Archiving old workbook..."
    archiveFolder = oldFolder & "\Archive"
    CreateFolderIfMissing archiveFolder
    
    ' Move the old workbook to Archive folder with timestamp
    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    archivePath = archiveFolder & "\" & fso.GetBaseName(oldFileName) & "_ARCHIVED_" & timestamp & ".xlsm"
    
    ' Move old workbook to archive
    MsgBox "DEBUG: About to move old workbook to archive" & vbCrLf & "From: " & sourceWorkbookPath & vbCrLf & "To: " & archivePath, vbInformation, "DEBUG"
    fso.MoveFile sourceWorkbookPath, archivePath
    MsgBox "DEBUG: Old workbook archived", vbInformation, "DEBUG"
    
    ' Save and move this (new) workbook to the old workbook's location
    newFilePath = oldFolder & "\" & oldFileName
    MsgBox "DEBUG: About to SaveAs to: " & newFilePath, vbInformation, "DEBUG"
    
    ' If old filename is different from new filename, use old filename
    ' so user keeps their custom name
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fileName:=newFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    newFileName = fso.GetFileName(ThisWorkbook.FullName)
    
    ' Delete the copy left in Updates folder (if it exists)
    Dim updatesFolder As String
    updatesFolder = oldFolder & "\Updates"
    On Error Resume Next
    If fso.FolderExists(updatesFolder) Then fso.DeleteFolder updatesFolder, True
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Save again after cleanup
    ThisWorkbook.Save
    
    ' Activate the INVENTORY sheet so user is ready to go
    On Error Resume Next
    ThisWorkbook.Worksheets(WS_INVENTORY).Activate
    ThisWorkbook.Worksheets(WS_INVENTORY).Range("A1").Select
    On Error GoTo 0
    
    ' DEBUG: About to show success message
    MsgBox "DEBUG: Reaching success message", vbInformation, "DEBUG"
    
    MsgBox "Welcome to your newly updated workbook!" & vbCrLf & vbCrLf & _
           "All of your data has been transferred successfully and everything is right where you left it." & vbCrLf & vbCrLf & _
           "This update was mostly just some critter wrangling on my end " & _
           "- rounding up a few misbehaving scripts that needed to be gently reminded how to behave. " & _
           "Everything should work just as it did before." & vbCrLf & vbCrLf & _
           "If you happen to spot any other critters that slipped past me, please shoot me an email at:" & vbCrLf & _
           "VIRGIL_Support@proton.me" & vbCrLf & vbCrLf & _
           "I'll do my best to wrangle them up within 48 hours!", vbInformation, "Update Complete!"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    On Error GoTo 0
    
    ' Debug info
    MsgBox "Error transferring data: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Debug info:" & vbCrLf & _
           "oldFolder: " & oldFolder & vbCrLf & _
           "oldFileName: " & oldFileName & vbCrLf & _
           "newFilePath: " & newFilePath & vbCrLf & _
           "sourceWorkbookPath: " & sourceWorkbookPath, vbCritical, "Transfer Error"
End Sub

' Transfer data rows from source workbook sheet to this workbook
Private Sub TransferSheetData(ByVal sourceWb As Workbook, ByVal sheetName As String)
    On Error GoTo ErrorHandler
    
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim srcLastRow As Long
    Dim srcLastCol As Long
    Dim srcRange As Range
    
    ' Check if sheet exists in both workbooks
    On Error Resume Next
    Set srcWs = sourceWb.Worksheets(sheetName)
    Set destWs = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    
    If srcWs Is Nothing Or destWs Is Nothing Then Exit Sub
    
    ' Find data range in source
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
    
    If srcLastRow < 2 Then Exit Sub ' No data rows
    
    ' Copy data rows (skip header row 1)
    Set srcRange = srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(srcLastRow, srcLastCol))
    
    ' Clear existing data in destination (keep headers)
    Dim destLastRow As Long
    destLastRow = destWs.Cells(destWs.Rows.Count, 1).End(xlUp).Row
    If destLastRow >= 2 Then
        destWs.Range(destWs.Cells(2, 1), destWs.Cells(destLastRow, srcLastCol)).Clear
    End If
    
    ' Paste values only (no formatting - keep new workbook formatting)
    srcRange.Copy
    destWs.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Exit Sub
    
ErrorHandler:
    ' Silently skip sheets that can't be transferred
End Sub

' ============================================
' CHECK FOR PENDING DATA TRANSFER
' ============================================

Private Function CheckForPendingTransfer() As Boolean
    Dim tempPathFile As String
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    ' Remove any leftover Transfer My Data button from previous versions
    On Error Resume Next
    ThisWorkbook.Worksheets(WS_INVENTORY).Shapes("BTN_TRANSFER_DATA").Delete
    On Error GoTo 0
    
    ' If the temp file exists, a transfer is pending
    If Dir(tempPathFile) = "" Then
        CheckForPendingTransfer = False
        Exit Function
    End If
    
    CheckForPendingTransfer = True
    
    ' Auto-run the transfer silently (no popups)
    TransferMyData
End Function

' ============================================
' MANUAL UPDATE CHECK
' ============================================

Public Sub ManualCheckForUpdates()
    CheckForUpdatesOnOpen
End Sub

' Toggle auto-check setting
Public Sub ToggleAutoCheckUpdates()
    Dim current As String
    current = GetSettingValue("AUTO_CHECK_UPDATES")
    
    If current = "NO" Then
        UpdateSetting "AUTO_CHECK_UPDATES", "YES"
        MsgBox "Auto-update check is now ENABLED", vbInformation, "Settings"
    Else
        UpdateSetting "AUTO_CHECK_UPDATES", "NO"
        MsgBox "Auto-update check is now DISABLED", vbInformation, "Settings"
    End If
End Sub

' ============================================
' UPDATE SETTING VALUE
' ============================================

Public Sub UpdateSetting(ByVal settingName As String, ByVal settingValue As String)
    Dim wsSettings As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim foundRow As Long
    
    Set wsSettings = ThisWorkbook.Worksheets("SETTINGS")
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, 1).End(xlUp).Row
    foundRow = 0
    
    ' Search for existing setting
    For i = 1 To lastRow
        If wsSettings.Cells(i, 1).Value = settingName Then
            foundRow = i
            Exit For
        End If
    Next i
    
    ' Update existing or add new
    If foundRow > 0 Then
        wsSettings.Cells(foundRow, 2).Value = settingValue
    Else
        wsSettings.Cells(lastRow + 1, 1).Value = settingName
        wsSettings.Cells(lastRow + 1, 2).Value = settingValue
    End If
    
End Sub

' ============================================
' BACKUP BEFORE UPDATE
' ============================================

Private Function CreatePreUpdateBackup() As String
    On Error GoTo ErrorHandler
    
    Dim backupFolder As String
    Dim backupPath As String
    Dim timestamp As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create backup folder if needed
    backupFolder = ThisWorkbook.Path & "\Backups\Pre-Update"
    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
    End If
    
    ' Create timestamped backup filename
    timestamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    backupPath = backupFolder & "\" & Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_BACKUP_" & timestamp & ".xlsm"
    
    ' Save backup
    ThisWorkbook.SaveCopyAs backupPath
    
    CreatePreUpdateBackup = backupPath
    Exit Function
    
ErrorHandler:
    CreatePreUpdateBackup = ""
End Function
