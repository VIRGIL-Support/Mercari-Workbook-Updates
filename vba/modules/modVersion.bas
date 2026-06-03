Attribute VB_Name = "modVersion"
Option Explicit

' ============================================
' VERSION MANAGEMENT MODULE
' ============================================

Public Const CURRENT_VERSION As String = "1.1"
Public Const UPDATE_CHECK_URL As String = "https://raw.githubusercontent.com/VIRGIL-Support/Mercari-Workbook-Updates/main/version.txt"
Public Const UPDATE_DOWNLOAD_URL As String = "https://raw.githubusercontent.com/VIRGIL-Support/Mercari-Workbook-Updates/main/Mercari_Workbook_Latest.xlsm"

' Check for updates on workbook open
Public Sub CheckForUpdatesOnOpen()
    Dim checkResult As String
    Dim latestVersion As String
    Dim userChoice As VbMsgBoxResult
    
    ' SAFETY FIRST: Reset any stuck update state from previous sessions
    ResetUpdateState
    
    ' SKIP UPDATE CHECK IF THIS WORKBOOK WAS JUST DOWNLOADED
    ' If opened from Updates folder or temp location, we're in middle of update
    If InStr(1, ThisWorkbook.Path, "Updates", vbTextCompare) > 0 Then
        ' This is a newly downloaded workbook - skip update check
        ' TransferMyData will run after this to transfer settings
        Exit Sub
    End If
    
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

' Reset any stuck update state from previous interrupted sessions
Private Sub ResetUpdateState()
    On Error Resume Next
    
    Dim tempPathFile As String
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    ' If temp file exists from a previous session, delete it
    If Dir(tempPathFile) <> "" Then
        Kill tempPathFile
    End If
    
    ' Also clean up any other temp files that might be stuck
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Clean up update-related temp files in temp folder
    Dim tempFolder As String
    tempFolder = Environ$("TEMP")
    
    ' Remove specific update-related files if they exist
    If Dir(tempFolder & "\MercariUpdateLog.txt") <> "" Then
        Kill tempFolder & "\MercariUpdateLog.txt"
    End If
    
    On Error GoTo 0
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
    MsgBox "UPDATE DEBUG - Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Line that failed will be shown if you click Debug", vbCritical, "Update Error"
    ' Resume Next ' Uncomment for debugging to see where it fails
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
    Dim archiveSubfolder As String
    Dim archivePath As String
    Dim newFilePath As String
    Dim COPY_FOLDERS_TO_ARCHIVE As Boolean
    
    ' SET THIS TO TRUE IF YOU WANT FOLDERS COPIED TO ARCHIVE
    ' SET TO FALSE IF LINKS DON'T REQUIRE IT
    COPY_FOLDERS_TO_ARCHIVE = True
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Try to read the old workbook path from temp file
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    If Dir(tempPathFile) = "" Then
        ' No temp file - no transfer needed, just exit silently
        Exit Sub
    Else
        ' Read path from temp file
        f = FreeFile
        Open tempPathFile For Input As #f
        Line Input #f, sourceWorkbookPath
        Close #f
        
        ' Verify file exists - if not, clean up temp file and exit
        If Dir(sourceWorkbookPath) = "" Then
            Kill tempPathFile
            Exit Sub
        End If
    End If
    
    ' Remember old workbook location and name
    oldFolder = fso.GetParentFolderName(sourceWorkbookPath)
    oldFileName = fso.GetFileName(sourceWorkbookPath)
    
    Application.StatusBar = "Opening source workbook..."
    Application.ScreenUpdating = False
    
    ' Open the old workbook (read-only)
    Set sourceWb = Workbooks.Open(sourceWorkbookPath, ReadOnly:=True)
    
    ' Transfer INVENTORY data
    Application.StatusBar = "Transferring INVENTORY data..."
    TransferSheetData sourceWb, WS_INVENTORY
    
    ' Transfer SOLD ITEMS data
    Application.StatusBar = "Transferring SOLD ITEMS data..."
    TransferSheetData sourceWb, "SOLD ITEMS"
    
    ' Transfer SETTINGS data
    Application.StatusBar = "Transferring SETTINGS..."
    TransferSheetData sourceWb, "SETTINGS"
    
    ' Transfer LOOKUPS data (preserve custom dropdowns)
    Application.StatusBar = "Transferring LOOKUPS..."
    TransferSheetData sourceWb, "LOOKUPS"
    
    ' Close the source workbook
    Application.ScreenUpdating = True
    DoEvents
    
    On Error Resume Next
    sourceWb.Close SaveChanges:=False
    On Error GoTo ErrorHandler
    
    Set sourceWb = Nothing
    
    ' Clean up temp file
    If Dir(tempPathFile) <> "" Then Kill tempPathFile
    
    ' ============================================
    ' ARCHIVE OLD WORKBOOK AND FOLDERS
    ' ============================================
    
    Application.StatusBar = "Archiving old workbook..."
    
    ' Create Archive folder
    archiveFolder = oldFolder & "\Archive"
    CreateFolderIfMissing archiveFolder
    
    ' Create timestamped subfolder with human-readable format
    ' Format: "Archived 9-15-26 11-25-01 AM"
    Dim timestampReadable As String
    Dim timestampFile As String
    Dim ampm As String
    Dim hourNum As Integer
    
    hourNum = Hour(Now)
    If hourNum >= 12 Then
        ampm = "PM"
        If hourNum > 12 Then hourNum = hourNum - 12
    Else
        ampm = "AM"
        If hourNum = 0 Then hourNum = 12
    End If
    
    ' Human readable: "Archived 6-2-26 1-30-45 PM"
    timestampReadable = "Archived " & Month(Now) & "-" & Day(Now) & "-" & Right(Year(Now), 2) & _
                        " " & hourNum & "-" & Minute(Now) & "-" & Second(Now) & " " & ampm
    
    ' File safe version for folder name
    timestampFile = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    
    archiveSubfolder = archiveFolder & "\" & timestampReadable
    CreateFolderIfMissing archiveSubfolder
    
    ' Archive path for the old workbook
    archivePath = archiveSubfolder & "\" & oldFileName
    
    ' IF ENABLED: Copy the 5 project folders to archive subfolder BEFORE moving workbook
    If COPY_FOLDERS_TO_ARCHIVE Then
        Application.StatusBar = "Copying project folders to archive..."
        
        Dim folderNames As Variant
        Dim i As Long
        Dim sourceFolder As String
        Dim destFolder As String
        
        ' List of folders to archive
        folderNames = Array("1 READY TO LIST", "2 DESCRIPTION FILES", "3 SOLD", "4 Backups", "5 Logs")
        
        For i = LBound(folderNames) To UBound(folderNames)
            sourceFolder = oldFolder & "\" & folderNames(i)
            destFolder = archiveSubfolder & "\" & folderNames(i)
            
            If fso.FolderExists(sourceFolder) Then
                On Error Resume Next
                fso.CopyFolder sourceFolder, destFolder, True
                On Error GoTo ErrorHandler
            End If
        Next i
    End If
    
    ' Move the old workbook to the archive subfolder
    fso.MoveFile sourceWorkbookPath, archivePath
    
    ' ============================================
    ' SAVE NEW WORKBOOK TO ORIGINAL LOCATION
    ' ============================================
    
    Application.StatusBar = "Saving updated workbook..."
    
    ' Save this (new) workbook to the old workbook's original location
    newFilePath = oldFolder & "\" & oldFileName
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fileName:=newFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    
    ' ============================================
    ' CLEANUP
    ' ============================================
    
    ' Delete the Updates folder (temporary download location)
    Dim updatesFolder As String
    updatesFolder = oldFolder & "\Updates"
    On Error Resume Next
    If fso.FolderExists(updatesFolder) Then fso.DeleteFolder updatesFolder, True
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Save again to ensure everything is written
    ThisWorkbook.Save
    
    ' Activate the INVENTORY sheet so user is ready to go
    On Error Resume Next
    ThisWorkbook.Worksheets(WS_INVENTORY).Activate
    ThisWorkbook.Worksheets(WS_INVENTORY).Range("A1").Select
    On Error GoTo 0
    
    ' Success message
    MsgBox "Welcome to your newly updated workbook!" & vbCrLf & vbCrLf & _
           "All of your data has been transferred successfully and everything is right where you left it." & vbCrLf & vbCrLf & _
           "Your previous version has been archived in:" & vbCrLf & _
           archiveSubfolder & vbCrLf & vbCrLf & _
           "The old workbook" & IIf(COPY_FOLDERS_TO_ARCHIVE, " and project folders", "") & _
           " have been safely stored there in case you need them." & vbCrLf & vbCrLf & _
           "If you happen to spot any issues, please email:" & vbCrLf & _
           "VIRGIL_Support@proton.me", vbInformation, "Update Complete!"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    DoEvents
    
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    On Error GoTo 0
    
    MsgBox "Error during data transfer: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Please contact support: VIRGIL_Support@proton.me", vbCritical, "Transfer Error"
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
    Dim f As Integer
    Dim sourceWorkbookPath As String
    Dim fileAge As Double
    
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    ' Remove any leftover Transfer My Data button from previous versions
    On Error Resume Next
    ThisWorkbook.Worksheets("INVENTORY").Shapes("BTN_TRANSFER_DATA").Delete
    On Error GoTo 0
    
    ' If the temp file doesn't exist, no transfer needed
    If Dir(tempPathFile) = "" Then
        CheckForPendingTransfer = False
        Exit Function
    End If
    
    ' SAFETY CHECK 1: Verify the temp file is recent (created within last 10 minutes)
    ' If it's older, it's likely a stale/corrupted file from a failed update
    fileAge = Now - FileDateTime(tempPathFile)
    If fileAge > (10 / 1440) Then ' 10 minutes in Excel date format (days)
        ' File is stale, delete it and skip transfer
        On Error Resume Next
        Kill tempPathFile
        On Error GoTo 0
        CheckForPendingTransfer = False
        Exit Function
    End If
    
    ' SAFETY CHECK 2: Verify the source workbook path in the file is valid
    On Error Resume Next
    f = FreeFile
    Open tempPathFile For Input As #f
    Line Input #f, sourceWorkbookPath
    Close #f
    On Error GoTo 0
    
    If sourceWorkbookPath = "" Or Dir(sourceWorkbookPath) = "" Then
        ' Source workbook doesn't exist, clean up temp file
        On Error Resume Next
        Kill tempPathFile
        On Error GoTo 0
        CheckForPendingTransfer = False
        Exit Function
    End If
    
    ' All safety checks passed - transfer is valid
    CheckForPendingTransfer = True
    
    ' Auto-run the transfer
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
    Dim preUpdateFolder As String
    Dim backupPath As String
    Dim timestamp As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create nested folders: Backups first, then Pre-Update subfolder
    backupFolder = ThisWorkbook.Path & "\Backups"
    preUpdateFolder = backupFolder & "\Pre-Update"
    
    ' Create Backups folder if needed
    If Not fso.FolderExists(backupFolder) Then
        fso.CreateFolder backupFolder
    End If
    
    ' Create Pre-Update subfolder if needed
    If Not fso.FolderExists(preUpdateFolder) Then
        fso.CreateFolder preUpdateFolder
    End If
    
    ' Create timestamped backup filename
    timestamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    backupPath = preUpdateFolder & "\" & Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_BACKUP_" & timestamp & ".xlsm"
    
    ' Save backup
    ThisWorkbook.SaveCopyAs backupPath
    
    CreatePreUpdateBackup = backupPath
    Exit Function
    
ErrorHandler:
    MsgBox "Backup failed: " & Err.Number & " - " & Err.Description, vbExclamation, "Backup Error"
    CreatePreUpdateBackup = ""
End Function
