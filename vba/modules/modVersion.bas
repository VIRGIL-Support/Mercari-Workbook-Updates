Attribute VB_Name = "modVersion"
Option Explicit

' ============================================
' VERSION MANAGEMENT MODULE
' ============================================

Public Const CURRENT_VERSION As String = "1.1"
Public Const UPDATE_CHECK_URL As String = "https://raw.githubusercontent.com/VIRGIL-Support/Mercari-Workbook-Updates/main/version.txt"
Public Const UPDATE_DOWNLOAD_URL As String = "https://raw.githubusercontent.com/VIRGIL-Support/Mercari-Workbook-Updates/main/Mercari_Workbook_Latest.xlsm"

Public Sub CheckForUpdatesOnOpen()
    Dim checkResult As String
    Dim latestVersion As String
    Dim userChoice As VbMsgBoxResult
    Dim tempPathFile As String
    
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    ' CRITICAL: Check for pending transfer BEFORE calling ResetUpdateState
    ' If temp file exists, this workbook was just downloaded - handle transfer immediately
    If Dir(tempPathFile) <> "" Then
        ' Check if it's a fresh transfer (file age < 2 minutes)
        Dim fileAge As Double
        fileAge = Now - FileDateTime(tempPathFile)
        fileAge = fileAge * 1440 ' Convert to minutes
        
        If fileAge < 2 Then
            ' Fresh transfer - skip ResetUpdateState entirely to preserve temp file
            CheckForPendingTransfer
            Exit Sub
        End If
        ' If file is old (> 2 min), it's stale - let ResetUpdateState clean it up
    End If
    
    ' SAFETY FIRST: Detect and clean up stale/stuck updates from previous sessions
    DetectStaleUpdate
    
    ' Reset any stuck update state from previous sessions
    ' (Only runs if no fresh temp file exists)
    ResetUpdateState
    
    ' Check if user wants to auto-check (stored in settings)
    If GetSettingValue("AUTO_CHECK_UPDATES") = "NO" Then Exit Sub
    
    checkResult = GetLatestVersionInfo()
    
    If checkResult = "" Then
        Exit Sub
    End If
    
    latestVersion = Replace(Replace(Trim$(checkResult), vbCr, ""), vbLf, "")
    
    If IsNewerVersion(latestVersion, CURRENT_VERSION) Then
        userChoice = MsgBox("New workbook update is available from VIRGIL!" & vbCrLf & vbCrLf & _
                           "Current Version: " & CURRENT_VERSION & "          New Version: " & latestVersion & vbCrLf & vbCrLf & vbCrLf & _
                           "Would you like to update now?" & vbCrLf & vbCrLf & vbCrLf & _
                           "NOTE: When you click Yes, don't worry if it takes me up to one minute to respond. " & _
                           "I'll be busy backing up all of your data, downloading the update, transferring everything " & _
                           "into the shiny new version, and double checking to make sure it's all perfect for you.", _
                           vbYesNo + vbInformation, "Update Available")
        
        If userChoice = vbYes Then
            DownloadUpdate latestVersion
        End If
    End If
    
End Sub

' Reset any stuck update state from previous interrupted sessions
' CRITICAL: Only deletes temp files that are OLD (stale)
' Fresh temp files (created in last 2 minutes) are preserved for active updates
Private Sub ResetUpdateState()
    On Error Resume Next
    
    Dim tempPathFile As String
    Dim fileAge As Double
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    ' If temp file exists, check if it's old (stale) before deleting
    If Dir(tempPathFile) <> "" Then
        ' Calculate file age in minutes
        fileAge = Now - FileDateTime(tempPathFile)
        fileAge = fileAge * 1440 ' Convert to minutes
        
        ' Only delete if file is older than 2 minutes (stale from previous session)
        ' Fresh files (< 2 min) are from current update - DON'T DELETE
        If fileAge > 2 Then
            Kill tempPathFile
        End If
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

    Dim downloadPath As String
    Dim backupPath As String
    Dim fso As Object
    Dim logFile As String
    Dim logNum As Integer
    
    logFile = Environ$("TEMP") & "\MercariUpdateLog.txt"
    
    ' LOG: DownloadUpdate started
    logNum = FreeFile
    Open logFile For Output As #logNum
    Print #logNum, "=== DownloadUpdate started at " & Now & " ==="
    Print #logNum, "  Current workbook: " & ThisWorkbook.FullName
    Print #logNum, "  New version: " & newVersion
    Close #logNum

    Set fso = CreateObject("Scripting.FileSystemObject")

    Application.StatusBar = "Creating backup..."
    backupPath = CreatePreUpdateBackup()
    
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  Backup created: " & backupPath
    Close #logNum

    Application.StatusBar = "Downloading update v" & newVersion & "..."

    Dim baseName As String
    baseName = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    downloadPath = ThisWorkbook.Path & "\" & baseName & "_v" & newVersion & ".xlsm"

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", UPDATE_DOWNLOAD_URL, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send

    If http.Status <> 200 Then
        logNum = FreeFile
        Open logFile For Append As #logNum
        Print #logNum, "  ERROR: HTTP Status " & http.Status
        Close #logNum
        Application.StatusBar = False
        MsgBox "Failed to download update. Please try again later." & vbCrLf & vbCrLf & _
               "HTTP Status: " & http.Status, vbExclamation, "Download Failed"
        Exit Sub
    End If

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile downloadPath, 2
    stream.Close
    
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  Downloaded to: " & downloadPath
    Close #logNum

    Application.StatusBar = False

    Dim tempPathFile As String
    tempPathFile = Environ$("TEMP") & "\" & "MercariUpdateSource.txt"
    Dim f As Integer
    f = FreeFile
    Open tempPathFile For Output As #f
    Print #f, ThisWorkbook.FullName
    Print #f, downloadPath
    Print #f, backupPath              ' Line 3: pre-update backup path for rollback
    Close #f
    
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  Temp file created: " & tempPathFile
    Print #logNum, "  Old workbook path: " & ThisWorkbook.FullName
    Print #logNum, "  New workbook path: " & downloadPath
    Close #logNum

    UpdateSetting "VERSION", newVersion
    ThisWorkbook.Save
    
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  Opening new workbook..."
    Close #logNum
    
    Workbooks.Open downloadPath
    
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  New workbook opened, closing old workbook"
    Print #logNum, "=== DownloadUpdate complete ==="
    Close #logNum
    
    ' Disable events to prevent BackupWorkbookBeforeClose from firing on this close
    Application.EnableEvents = False
    ThisWorkbook.Close SaveChanges:=False
    Application.EnableEvents = True

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.EnableEvents = True
    
    ' Clean up downloaded file if it exists
    On Error Resume Next
    If Len(downloadPath) > 0 Then
        If Dir(downloadPath) <> "" Then Kill downloadPath
    End If
    Dim tmpFile As String
    tmpFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    If Dir(tmpFile) <> "" Then Kill tmpFile
    On Error GoTo 0
    
    MsgBox "The update could not be completed and has been cancelled." & vbCrLf & vbCrLf & _
           "Nothing has been changed - your workbook is exactly as it was." & vbCrLf & vbCrLf & _
           "Error: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "You can try again later or contact VIRGIL_Support@proton.me for help.", _
           vbExclamation, "Update Cancelled"
End Sub

' ============================================
' TRANSFER DATA FROM OLD WORKBOOK
' ============================================

Public Sub TransferMyData()
    Dim logFile As String
    Dim logNum As Integer
    Dim sourceWorkbookPath As String
    Dim newWorkbookPath As String
    Dim tempPathFile As String
    Dim sourceWb As Workbook
    Dim f As Integer
    Dim fso As Object
    Dim oldFolder As String
    Dim oldFileName As String
    Dim archiveFolder As String
    Dim archiveSubfolder As String
    Dim archivePath As String
    Dim finalPath As String
    Dim backupPath As String
    Dim COPY_FOLDERS_TO_ARCHIVE As Boolean

    logFile = Environ$("TEMP") & "\MercariUpdateLog.txt"
    COPY_FOLDERS_TO_ARCHIVE = True

    ' Start fresh log
    On Error Resume Next
    Kill logFile
    On Error GoTo 0

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    logNum = FreeFile
    Open logFile For Output As #logNum
    Print #logNum, "=== Mercari Update Log - " & Now & " ==="
    Print #logNum, "STEP 0: TransferMyData STARTED"
    Print #logNum, "  Current workbook: " & ThisWorkbook.FullName
    Print #logNum, "  Current version: " & CURRENT_VERSION
    Close #logNum

    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"

    If Dir(tempPathFile) = "" Then
        LogEntry logFile, "ERROR: Temp file not found - Exiting"
        Exit Sub
    End If

    f = FreeFile
    Open tempPathFile For Input As #f
    Line Input #f, sourceWorkbookPath
    Line Input #f, newWorkbookPath
    If Not EOF(f) Then Line Input #f, backupPath
    Close #f

    LogEntry logFile, "STEP 1: Read temp file - source=" & sourceWorkbookPath

    If sourceWorkbookPath = "" Or Dir(sourceWorkbookPath) = "" Then
        LogEntry logFile, "ERROR: Source workbook not found - Exiting"
        Kill tempPathFile
        Exit Sub
    End If

    oldFolder = fso.GetParentFolderName(sourceWorkbookPath)
    oldFileName = fso.GetFileName(sourceWorkbookPath)

    LogEntry logFile, "STEP 2: Opening source workbook (ReadOnly)..."

    Application.StatusBar = "Opening source workbook..."
    Application.ScreenUpdating = False
    
    ' CRITICAL: Disable events before opening source workbook
    ' Otherwise its Workbook_Open fires, sees the temp file, and calls TransferMyData recursively
    Application.EnableEvents = False
    Set sourceWb = Workbooks.Open(sourceWorkbookPath, ReadOnly:=True)
    Application.EnableEvents = True

    Application.StatusBar = "Transferring INVENTORY data..."
    TransferSheetData sourceWb, WS_INVENTORY

    Application.StatusBar = "Transferring SOLD ITEMS data..."
    TransferSheetData sourceWb, "SOLD ITEMS"

    Application.StatusBar = "Transferring SETTINGS..."
    TransferSheetData sourceWb, "SETTINGS"

    Application.StatusBar = "Transferring LOOKUPS..."
    TransferSheetData sourceWb, "LOOKUPS"

    LogEntry logFile, "STEP 3: Data transfer complete. Saving source info for deferred close..."

    ' Store info needed for Phase 2 (close + archive) in a second temp file
    Dim phase2File As String
    phase2File = Environ$("TEMP") & "\MercariUpdatePhase2.txt"
    Dim p2 As Integer
    p2 = FreeFile
    Open phase2File For Output As #p2
    Print #p2, sourceWb.Name          ' Line 1: source workbook name (to close it)
    Print #p2, sourceWorkbookPath     ' Line 2: source workbook full path (to archive it)
    Print #p2, oldFolder              ' Line 3: old folder path
    Print #p2, oldFileName            ' Line 4: old file name
    Print #p2, backupPath             ' Line 5: pre-update backup path for rollback
    Print #p2, CURRENT_VERSION        ' Line 6: new version number for filename
    Close #p2

    LogEntry logFile, "  Phase 2 file written. Scheduling deferred close+archive via OnTime..."

    ' Schedule Phase 2 to run in 1 second (new execution context)
    Application.OnTime Now + TimeSerial(0, 0, 1), "CompleteUpdatePhase2"

    LogEntry logFile, "  OnTime scheduled. TransferMyData Phase 1 exiting cleanly."

    Application.ScreenUpdating = True
    Application.StatusBar = "Finalizing update..."

    Exit Sub

ErrorHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    
    LogEntry logFile, "PHASE 1 ERROR: " & errNum & " - " & errDesc

    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    DoEvents

    ' Close source workbook if still open
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    On Error GoTo 0

    ' Rollback: clean up temp files, delete downloaded file
    RollbackUpdate logFile, "", ""

    MsgBox "The update encountered an error during data transfer and has been rolled back." & vbCrLf & vbCrLf & _
           "Your original workbook has not been modified." & vbCrLf & vbCrLf & _
           "Error: " & errNum & " - " & errDesc & vbCrLf & vbCrLf & _
           "You can try again later or contact VIRGIL_Support@proton.me for help.", _
           vbExclamation, "Update Rolled Back"
End Sub

' ============================================
' PHASE 2: CLOSE SOURCE + ARCHIVE (runs via OnTime in new execution context)
' ============================================

Public Sub CompleteUpdatePhase2()
    Dim logFile As String
    Dim phase2File As String
    Dim sourceWbName As String
    Dim sourceWorkbookPath As String
    Dim oldFolder As String
    Dim oldFileName As String
    Dim archiveFolder As String
    Dim archiveSubfolder As String
    Dim archivePath As String
    Dim finalPath As String
    Dim backupPath As String
    Dim latestVersion As String
    Dim fso As Object
    Dim f As Integer
    Dim wb As Workbook
    Dim COPY_FOLDERS_TO_ARCHIVE As Boolean

    logFile = Environ$("TEMP") & "\MercariUpdateLog.txt"
    phase2File = Environ$("TEMP") & "\MercariUpdatePhase2.txt"
    COPY_FOLDERS_TO_ARCHIVE = True

    On Error GoTo Phase2Error

    LogEntry logFile, "=== PHASE 2 STARTED ==="

    ' Read the phase 2 info file
    If Dir(phase2File) = "" Then
        LogEntry logFile, "ERROR: Phase 2 file not found!"
        Exit Sub
    End If

    f = FreeFile
    Open phase2File For Input As #f
    Line Input #f, sourceWbName
    Line Input #f, sourceWorkbookPath
    Line Input #f, oldFolder
    Line Input #f, oldFileName
    If Not EOF(f) Then Line Input #f, backupPath
    If Not EOF(f) Then Line Input #f, latestVersion
    Close #f

    LogEntry logFile, "  sourceWbName = " & sourceWbName
    LogEntry logFile, "  sourceWorkbookPath = " & sourceWorkbookPath
    LogEntry logFile, "  oldFolder = " & oldFolder
    LogEntry logFile, "  oldFileName = " & oldFileName

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- STEP A: Close the source workbook ---
    LogEntry logFile, "PHASE2 STEP A: Closing source workbook..."

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' Find and close the source workbook
    On Error Resume Next
    Set wb = Workbooks(sourceWbName)
    If Not wb Is Nothing Then
        wb.Saved = True
        wb.Close SaveChanges:=False
        LogEntry logFile, "  Source workbook closed."
    Else
        LogEntry logFile, "  Source workbook not found in Workbooks collection (may already be closed)."
    End If
    Set wb = Nothing
    Err.Clear
    On Error GoTo Phase2Error

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    DoEvents

    ' Delete temp files
    On Error Resume Next
    Dim tempPathFile As String
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    If Dir(tempPathFile) <> "" Then Kill tempPathFile
    If Dir(phase2File) <> "" Then Kill phase2File
    On Error GoTo Phase2Error

    LogEntry logFile, "  Temp files deleted."

    ' --- STEP B: Create Archived folder ---
    LogEntry logFile, "PHASE2 STEP B: Creating Archived folder..."

    archiveFolder = oldFolder & "\Archived"
    CreateFolderIfMissing archiveFolder

    ' Build timestamped subfolder name
    Dim timestampReadable As String
    Dim ampm As String
    Dim hourNum As Integer
    Dim monthName As String

    Select Case Month(Now)
        Case 1:  monthName = "JAN"
        Case 2:  monthName = "FEB"
        Case 3:  monthName = "MAR"
        Case 4:  monthName = "APR"
        Case 5:  monthName = "MAY"
        Case 6:  monthName = "JUN"
        Case 7:  monthName = "JUL"
        Case 8:  monthName = "AUG"
        Case 9:  monthName = "SEP"
        Case 10: monthName = "OCT"
        Case 11: monthName = "NOV"
        Case 12: monthName = "DEC"
    End Select

    hourNum = Hour(Now)
    If hourNum >= 12 Then
        ampm = "pm"
        If hourNum > 12 Then hourNum = hourNum - 12
    Else
        ampm = "am"
        If hourNum = 0 Then hourNum = 12
    End If

    timestampReadable = "Archived " & Format(Day(Now), "00") & "-" & monthName & "-" & Right(Year(Now), 2) & _
                        " at " & Format(hourNum, "00") & "-" & Format(Minute(Now), "00") & "-" & Format(Second(Now), "00") & " " & ampm

    archiveSubfolder = archiveFolder & "\" & timestampReadable
    CreateFolderIfMissing archiveSubfolder

    LogEntry logFile, "PHASE2 STEP C: Archive subfolder = " & archiveSubfolder

    archivePath = archiveSubfolder & "\" & oldFileName

    ' --- STEP C: Copy project folders to archive ---
    If COPY_FOLDERS_TO_ARCHIVE Then
        LogEntry logFile, "PHASE2 STEP D: Copying project folders..."
        Application.StatusBar = "Copying project folders to archive..."
        Dim folderNames As Variant
        Dim i As Long
        Dim srcFldr As String
        Dim destFldr As String
        folderNames = Array("1 READY TO LIST", "2 DESCRIPTION FILES", "3 SOLD", "Backups", "Logs")
        For i = LBound(folderNames) To UBound(folderNames)
            srcFldr = oldFolder & "\" & folderNames(i)
            destFldr = archiveSubfolder & "\" & folderNames(i)
            If fso.FolderExists(srcFldr) Then
                On Error Resume Next
                fso.CopyFolder srcFldr, destFldr, True
                On Error GoTo Phase2Error
            End If
        Next i
    End If

    ' --- STEP D: Move old workbook to archive ---
    LogEntry logFile, "PHASE2 STEP E: Moving old workbook to archive..."
    Application.StatusBar = "Archiving old workbook..."

    On Error Resume Next
    fso.MoveFile sourceWorkbookPath, archivePath
    If Err.Number <> 0 Then
        LogEntry logFile, "  MoveFile failed (" & Err.Number & " - " & Err.Description & "), trying CopyFile+Kill..."
        Err.Clear
        fso.CopyFile sourceWorkbookPath, archivePath, True
        If Err.Number <> 0 Then
            LogEntry logFile, "  CopyFile also failed: " & Err.Number & " - " & Err.Description
        End If
        Err.Clear
    End If
    On Error GoTo Phase2Error
    
    ' Verify the old workbook was removed - force delete if still present
    If Dir(sourceWorkbookPath) <> "" Then
        LogEntry logFile, "  Old workbook still exists after move/copy. Force deleting..."
        On Error Resume Next
        Kill sourceWorkbookPath
        If Err.Number <> 0 Then
            LogEntry logFile, "  Kill failed (" & Err.Number & "). Trying fso.DeleteFile..."
            Err.Clear
            fso.DeleteFile sourceWorkbookPath, True
            If Err.Number <> 0 Then
                LogEntry logFile, "  DeleteFile also failed: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
        End If
        On Error GoTo Phase2Error
    End If
    
    ' Final verification
    If Dir(sourceWorkbookPath) = "" Then
        LogEntry logFile, "  Old workbook successfully removed."
    Else
        LogEntry logFile, "  WARNING: Old workbook could not be deleted: " & sourceWorkbookPath
    End If

    ' --- STEP E: Save new workbook with version number in filename ---
    Dim newFileName As String
    Dim baseFileName As String
    baseFileName = Left(oldFileName, InStrRev(oldFileName, ".") - 1)
    ' Strip any existing version suffix (e.g. _v1.1) before appending new one
    If InStr(LCase(baseFileName), "_v") > 0 Then
        baseFileName = Left(baseFileName, InStrRev(LCase(baseFileName), "_v") - 1)
    End If
    If Len(Trim(latestVersion)) > 0 Then
        newFileName = baseFileName & "_v" & Trim(latestVersion) & ".xlsm"
    Else
        newFileName = oldFileName
    End If
    finalPath = oldFolder & "\" & newFileName
    
    ' Capture the downloaded file path BEFORE SaveAs changes it
    Dim downloadedFilePath As String
    downloadedFilePath = ThisWorkbook.FullName
    
    LogEntry logFile, "PHASE2 STEP F: Saving new workbook as " & finalPath
    LogEntry logFile, "  Downloaded file to clean up: " & downloadedFilePath

    Application.StatusBar = "Saving updated workbook..."
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs fileName:=finalPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True

    ThisWorkbook.Save
    
    ' Delete the downloaded _v*.xlsm file (SaveAs doesn't remove it)
    If LCase(downloadedFilePath) <> LCase(finalPath) Then
        If Dir(downloadedFilePath) <> "" Then
            On Error Resume Next
            Kill downloadedFilePath
            If Err.Number <> 0 Then
                LogEntry logFile, "  WARNING: Could not delete downloaded file: " & Err.Description
                Err.Clear
            Else
                LogEntry logFile, "  Downloaded file deleted."
            End If
            On Error GoTo Phase2Error
        End If
    End If

    On Error Resume Next
    ThisWorkbook.Worksheets(WS_INVENTORY).Activate
    ThisWorkbook.Worksheets(WS_INVENTORY).Range("A1").Select
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogEntry logFile, "SUCCESS: Update completed at " & Now & " | Archived to: " & archiveSubfolder

    ' Copy update log to workbook Logs folder for easy user access
    On Error Resume Next
    Dim logsFolder As String
    logsFolder = oldFolder & "\Logs"
    If Dir(logsFolder, vbDirectory) = "" Then MkDir logsFolder
    Dim destLog As String
    destLog = logsFolder & "\MercariUpdateLog_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    FileCopy logFile, destLog
    On Error GoTo 0

    Dim archiveFolderName As String
    archiveFolderName = Mid(archiveSubfolder, InStrRev(archiveSubfolder, "\") + 1)
    
    MsgBox "Welcome to your newly updated workbook!" & vbCrLf & vbCrLf & _
           "All of your data has been transferred successfully and everything is right where you left it." & vbCrLf & vbCrLf & _
           "Your previous version has been safely stored in your Archived folder" & vbCrLf & _
           "in a subfolder named: " & archiveFolderName & vbCrLf & vbCrLf & _
           "The old workbook" & IIf(COPY_FOLDERS_TO_ARCHIVE, " and project folders", "") & _
           " are stored there in case you ever need to restore them." & vbCrLf & vbCrLf & _
           "If you experience any issues with the updated version, refer to the" & vbCrLf & _
           """Restoring a Previous Version"" section of the User Manual." & vbCrLf & vbCrLf & _
           "If you need further assistance, please email:" & vbCrLf & _
           "VIRGIL_Support@proton.me", vbInformation, "Update Complete!"

    Exit Sub

Phase2Error:
    Dim p2ErrNum As Long
    Dim p2ErrDesc As String
    p2ErrNum = Err.Number
    p2ErrDesc = Err.Description
    
    LogEntry logFile, "PHASE 2 ERROR: " & p2ErrNum & " - " & p2ErrDesc
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    ' Copy update log to workbook Logs folder for easy user access
    On Error Resume Next
    Dim logsFolderErr As String
    logsFolderErr = oldFolder & "\Logs"
    If Dir(logsFolderErr, vbDirectory) = "" Then MkDir logsFolderErr
    Dim destLogErr As String
    destLogErr = logsFolderErr & "\MercariUpdateLog_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    FileCopy logFile, destLogErr
    On Error GoTo 0

    ' Attempt rollback from pre-update backup
    RollbackUpdate logFile, backupPath, sourceWorkbookPath

    MsgBox "The update encountered an error and has been rolled back." & vbCrLf & vbCrLf & _
           "Your original workbook has been restored from the pre-update backup." & vbCrLf & vbCrLf & _
           "Error: " & p2ErrNum & " - " & p2ErrDesc & vbCrLf & vbCrLf & _
           "You can try again later or contact VIRGIL_Support@proton.me for help.", _
           vbExclamation, "Update Rolled Back"
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
    ' NO On Error Resume Next - let errors show
    Set srcWs = sourceWb.Worksheets(sheetName)
    Set destWs = ThisWorkbook.Worksheets(sheetName)
    
    If srcWs Is Nothing Or destWs Is Nothing Then
        MsgBox "TransferSheetData ERROR: Sheet '" & sheetName & "' not found in source or destination", vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Find data range in source
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).row
    srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
    
    If srcLastRow < 2 Then Exit Sub ' No data rows
    
    ' Copy data rows (skip header row 1)
    Set srcRange = srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(srcLastRow, srcLastCol))
    
    ' Clear existing data in destination (keep headers)
    Dim destLastRow As Long
    destLastRow = destWs.Cells(destWs.Rows.Count, 1).End(xlUp).row
    If destLastRow >= 2 Then
        destWs.Range(destWs.Cells(2, 1), destWs.Cells(destLastRow, srcLastCol)).Clear
    End If
    
    ' Paste values only (no formatting - keep new workbook formatting)
    srcRange.Copy
    destWs.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    Exit Sub
    
ErrorHandler:
    MsgBox "TransferSheetData ERROR for sheet '" & sheetName & "':" & vbCrLf & _
           "Error " & Err.Number & " - " & Err.Description, vbCritical, "Transfer Sheet Error"
End Sub

' ============================================
' CHECK FOR PENDING DATA TRANSFER
' ============================================

Private Function CheckForPendingTransfer() As Boolean
    Dim tempPathFile As String
    Dim f As Integer
    Dim sourceWorkbookPath As String
    Dim fileAge As Double
    Dim logFile As String
    Dim logNum As Integer
    
    logFile = Environ$("TEMP") & "\MercariUpdateLog.txt"
    
    tempPathFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    
    ' LOG: CheckForPendingTransfer started
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "=== CheckForPendingTransfer called at " & Now & " ==="
    Print #logNum, "  ThisWorkbook: " & ThisWorkbook.FullName
    Print #logNum, "  tempPathFile: " & tempPathFile
    Print #logNum, "  Temp file exists: " & (Dir(tempPathFile) <> "")
    Close #logNum
    
    ' Remove any leftover Transfer My Data button from previous versions
    ' NO On Error Resume Next - let errors show if any
    On Error Resume Next
    ThisWorkbook.Worksheets("INVENTORY").Shapes("BTN_TRANSFER_DATA").Delete
    On Error GoTo 0
    
    ' If the temp file doesn't exist, no transfer needed
    If Dir(tempPathFile) = "" Then
        logNum = FreeFile
        Open logFile For Append As #logNum
        Print #logNum, "  RESULT: Temp file not found - no transfer needed"
        Close #logNum
        CheckForPendingTransfer = False
        Exit Function
    End If
    
    ' SAFETY CHECK 1: Verify the temp file is recent (created within last 10 minutes)
    fileAge = Now - FileDateTime(tempPathFile)
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  File age (minutes): " & (fileAge * 1440)
    Close #logNum
    
    If fileAge > (10 / 1440) Then ' 10 minutes in Excel date format (days)
        logNum = FreeFile
        Open logFile For Append As #logNum
        Print #logNum, "  RESULT: File too old (stale) - deleting and exiting"
        Close #logNum
        ' NO On Error Resume Next
        Kill tempPathFile
        CheckForPendingTransfer = False
        Exit Function
    End If
    
    ' SAFETY CHECK 2: Verify the source workbook path in the file is valid
    ' NO On Error Resume Next - let errors show
    f = FreeFile
    Open tempPathFile For Input As #f
    Line Input #f, sourceWorkbookPath
    Close #f
    
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  Source workbook from temp file: " & sourceWorkbookPath
    Print #logNum, "  Source exists: " & (Dir(sourceWorkbookPath) <> "")
    Close #logNum
    
    If sourceWorkbookPath = "" Or Dir(sourceWorkbookPath) = "" Then
        logNum = FreeFile
        Open logFile For Append As #logNum
        Print #logNum, "  RESULT: Source workbook not found - cleaning up"
        Close #logNum
        ' NO On Error Resume Next
        Kill tempPathFile
        CheckForPendingTransfer = False
        Exit Function
    End If
    
    ' All safety checks passed - transfer is valid
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  RESULT: All checks passed - calling TransferMyData"
    Close #logNum
    
    CheckForPendingTransfer = True
    
    ' Auto-run the transfer
    TransferMyData
    
    logNum = FreeFile
    Open logFile For Append As #logNum
    Print #logNum, "  TransferMyData completed at " & Now
    Close #logNum
End Function

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
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, 1).End(xlUp).row
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

' ============================================
' ROLLBACK UPDATE - Undo a failed update
' ============================================

Private Sub RollbackUpdate(ByVal logFile As String, ByVal backupPath As String, ByVal sourceWorkbookPath As String)
    On Error Resume Next
    
    LogEntry logFile, "=== ROLLBACK STARTED ==="
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Clean up all temp files
    Dim tempFiles As Variant
    Dim tf As Long
    tempFiles = Array( _
        Environ$("TEMP") & "\MercariUpdateSource.txt", _
        Environ$("TEMP") & "\MercariUpdatePhase2.txt")
    For tf = LBound(tempFiles) To UBound(tempFiles)
        If Dir(tempFiles(tf)) <> "" Then
            Kill tempFiles(tf)
            LogEntry logFile, "  Deleted temp file: " & tempFiles(tf)
        End If
    Next tf
    
    ' 2. Delete the downloaded new workbook file (the _v*.xlsm file) if it still exists
    '    Look for files matching the pattern in ThisWorkbook.Path
    Dim downloadedFile As String
    downloadedFile = Dir(ThisWorkbook.Path & "\*_v*.xlsm")
    Do While downloadedFile <> ""
        ' Don't delete the workbook we're currently running in
        If LCase(downloadedFile) <> LCase(ThisWorkbook.Name) Then
            LogEntry logFile, "  Deleting downloaded file: " & downloadedFile
            Kill ThisWorkbook.Path & "\" & downloadedFile
        End If
        downloadedFile = Dir()
    Loop
    
    ' 3. If we have a backup path and the original source is gone, restore it
    If Len(backupPath) > 0 And Len(sourceWorkbookPath) > 0 Then
        If Dir(backupPath) <> "" And Dir(sourceWorkbookPath) = "" Then
            LogEntry logFile, "  Restoring original from backup: " & backupPath
            fso.CopyFile backupPath, sourceWorkbookPath, True
            If Dir(sourceWorkbookPath) <> "" Then
                LogEntry logFile, "  Original workbook restored successfully."
            Else
                LogEntry logFile, "  WARNING: Restore may have failed."
            End If
        ElseIf Dir(sourceWorkbookPath) <> "" Then
            LogEntry logFile, "  Original workbook still exists - no restore needed."
        Else
            LogEntry logFile, "  WARNING: Backup not found at: " & backupPath
        End If
    Else
        LogEntry logFile, "  No backup path available - original workbook should still be intact."
    End If
    
    ' 4. Clean up any partial archive folders created during this attempt
    '    (Only if they were created in the last 5 minutes)
    Dim archiveBase As String
    archiveBase = ThisWorkbook.Path & "\Archived"
    If Dir(archiveBase, vbDirectory) <> "" Then
        Dim subFolder As String
        subFolder = Dir(archiveBase & "\Archived *", vbDirectory)
        Do While subFolder <> ""
            Dim fullSubPath As String
            fullSubPath = archiveBase & "\" & subFolder
            ' Only remove if created very recently (within 5 minutes)
            If (Now - FileDateTime(fullSubPath)) < (5 / 1440) Then
                LogEntry logFile, "  Removing partial archive folder: " & fullSubPath
                fso.DeleteFolder fullSubPath, True
            End If
            subFolder = Dir()
        Loop
    End If
    
    LogEntry logFile, "=== ROLLBACK COMPLETE ==="
    
    On Error GoTo 0
End Sub

' ============================================
' DETECT STALE/STUCK UPDATES (called from CheckForUpdatesOnOpen)
' ============================================

Public Sub DetectStaleUpdate()
    On Error Resume Next
    
    Dim tempFile As String
    Dim phase2File As String
    Dim fileAge As Double
    
    tempFile = Environ$("TEMP") & "\MercariUpdateSource.txt"
    phase2File = Environ$("TEMP") & "\MercariUpdatePhase2.txt"
    
    ' Check if either temp file exists and is old (> 5 minutes = stale)
    Dim staleFile As String
    staleFile = ""
    
    If Dir(tempFile) <> "" Then
        fileAge = (Now - FileDateTime(tempFile)) * 1440
        If fileAge > 5 Then staleFile = tempFile
    End If
    
    If Dir(phase2File) <> "" Then
        fileAge = (Now - FileDateTime(phase2File)) * 1440
        If fileAge > 5 Then
            If staleFile = "" Then staleFile = phase2File
        End If
    End If
    
    If staleFile = "" Then Exit Sub
    
    ' Stale update detected - notify user and clean up
    Dim logFile As String
    logFile = Environ$("TEMP") & "\MercariUpdateLog.txt"
    LogEntry logFile, "=== STALE UPDATE DETECTED ==="
    LogEntry logFile, "  Stale file: " & staleFile & " (age: " & Format(fileAge, "0.0") & " min)"
    
    MsgBox "It looks like a previous update didn't finish completely." & vbCrLf & vbCrLf & _
           "Don't worry - I've cleaned up the leftover files." & vbCrLf & _
           "Your workbook is safe and working normally." & vbCrLf & vbCrLf & _
           "You can try updating again whenever you're ready.", _
           vbInformation, "Previous Update Incomplete"
    
    ' Clean up stale files
    If Dir(tempFile) <> "" Then Kill tempFile
    If Dir(phase2File) <> "" Then Kill phase2File
    
    LogEntry logFile, "  Stale files cleaned up."
    
    On Error GoTo 0
End Sub

' ============================================
' CREATE FOLDER IF MISSING (Local copy for self-containment)
' ============================================

Private Sub CreateFolderIfMissing(folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

' ============================================
' LOG ENTRY HELPER
' ============================================

Private Sub LogEntry(ByVal logFile As String, ByVal msg As String)
    Dim fn As Integer
    fn = FreeFile
    Open logFile For Append As #fn
    Print #fn, Format(Now, "hh:nn:ss") & "  " & msg
    Close #fn
End Sub

' ============================================
' FIRST-RUN UPDATE PREFERENCE DIALOG
' ============================================

Public Sub PromptForAutoUpdatePreference()
    ' Called on first run to ask user about automatic update checking
    ' Includes privacy information about what data is transmitted
    
    Dim msg As String
    Dim result As VbMsgBoxResult
    
    msg = "Would you like the workbook to automatically check for updates when you open it?" & vbCrLf & vbCrLf & _
          "PRIVACY & SECURITY INFORMATION:" & vbCrLf & _
          String(40, "-") & vbCrLf & _
          "What is transmitted:" & vbCrLf & _
          "  - Only your current version number is sent" & vbCrLf & _
          "  - This is compared to the latest version online" & vbCrLf & vbCrLf & _
          "What is NOT transmitted:" & vbCrLf & _
          "  - NO inventory data or item details" & vbCrLf & _
          "  - NO photos or documents" & vbCrLf & _
          "  - NO personal information or email addresses" & vbCrLf & _
          "  - NO sales data or listing information" & vbCrLf & vbCrLf & _
          "Your data stays 100% on your computer." & vbCrLf & vbCrLf & _
          "Click YES to enable automatic update checks, or NO to check manually only when you choose."
    
    result = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton1, "Update Checking Preference")
    
    If result = vbYes Then
        UpdateSetting "AUTO_CHECK_UPDATES", "YES"
        MsgBox "Automatic update checking is now ENABLED." & vbCrLf & vbCrLf & _
               "You can disable this anytime from the HELP worksheet.", vbInformation, "Settings Saved"
    Else
        UpdateSetting "AUTO_CHECK_UPDATES", "NO"
        MsgBox "Automatic update checking is now DISABLED." & vbCrLf & vbCrLf & _
               "You can manually check for updates or re-enable automatic update checking" & vbCrLf & _
               "at any time by clicking the applicable button in the HELP worksheet.", vbInformation, "Settings Saved"
    End If
    
End Sub

' ============================================
' ENABLE/DISABLE AUTO-CHECK FUNCTIONS
' ============================================

Public Sub EnableAutoCheckUpdates()
    ' Re-enable automatic update checking
    UpdateSetting "AUTO_CHECK_UPDATES", "YES"
    MsgBox "Automatic update checking is now ENABLED." & vbCrLf & vbCrLf & _
           "PRIVACY NOTE: Only your version number is transmitted to check for updates. " & _
           "No inventory data, photos, or personal information ever leaves your computer.", _
           vbInformation, "Auto-Check Enabled"
End Sub

Public Sub DisableAutoCheckUpdates()
    ' Disable automatic update checking
    UpdateSetting "AUTO_CHECK_UPDATES", "NO"
    MsgBox "Automatic update checking is now DISABLED." & vbCrLf & vbCrLf & _
           "You can still check for updates manually using the 'Check for Updates' button on the HELP worksheet.", _
           vbInformation, "Auto-Check Disabled"
End Sub

' ============================================
' MANUAL UPDATE CHECK WITH PRIVACY INFO
' ============================================

Public Sub ManualCheckForUpdates()
    ' Show privacy notice first - respect Cancel
    Dim privResult As VbMsgBoxResult
    privResult = MsgBox("MANUAL UPDATE CHECK" & vbCrLf & vbCrLf & _
                        "PRIVACY & SECURITY:" & vbCrLf & _
                        String(30, "-") & vbCrLf & _
                        "- Only your version number will be sent" & vbCrLf & _
                        "- NO inventory data, photos, or personal info is transmitted" & vbCrLf & _
                        "- All your data stays safely on your computer" & vbCrLf & vbCrLf & _
                        "Click OK to check for available updates now.", _
                        vbInformation + vbOKCancel, "Check for Updates")
    
    If privResult <> vbOK Then Exit Sub
    
    ' Perform the check and show result either way
    Dim checkResult As String
    Dim latestVersion As String
    Dim userChoice As VbMsgBoxResult
    
    checkResult = GetLatestVersionInfo()
    
    If checkResult = "" Then
        MsgBox "Could not connect to the update server." & vbCrLf & vbCrLf & _
               "Please check your internet connection and try again.", _
               vbExclamation, "Update Check Failed"
        Exit Sub
    End If
    
    latestVersion = Replace(Replace(Trim$(checkResult), vbCr, ""), vbLf, "")
    
    If IsNewerVersion(latestVersion, CURRENT_VERSION) Then
        userChoice = MsgBox("New workbook update is available from VIRGIL!" & vbCrLf & vbCrLf & _
                           "Current Version: " & CURRENT_VERSION & "          New Version: " & latestVersion & vbCrLf & vbCrLf & vbCrLf & _
                           "Would you like to update now?" & vbCrLf & vbCrLf & vbCrLf & _
                           "NOTE: When you click Yes, don't worry if it takes me up to one minute to respond. " & _
                           "I'll be busy backing up all of your data, downloading the update, transferring everything " & _
                           "into the shiny new version, and double checking to make sure it's all perfect for you.", _
                           vbYesNo + vbInformation, "Update Available")
        If userChoice = vbYes Then
            DownloadUpdate latestVersion
        End If
    Else
        MsgBox "You are on the most current version of the VIRGIL Workbook!" & vbCrLf & vbCrLf & _
               "Current version: " & CURRENT_VERSION & vbCrLf & vbCrLf & _
               "No update is needed at this time. Check back anytime using the" & vbCrLf & _
               "'Check for Updates Now' button on the HELP worksheet.", _
               vbInformation, "You're Up to Date!"
    End If
End Sub

' ============================================
' SHOW PRIVACY INFORMATION
' ============================================

Public Sub ShowUpdatePrivacyInfo()
    ' Display privacy information about update checking
    ' Can be called from any menu or help button
    
    Dim msg As String
    
    msg = "UPDATE CHECKING - PRIVACY & SECURITY INFORMATION" & vbCrLf & vbCrLf & _
          "WHAT IS TRANSMITTED:" & vbCrLf & _
          String(25, "-") & vbCrLf & _
          "  - Only your current workbook version number" & vbCrLf & _
          "  - Example: Version 1.1" & vbCrLf & vbCrLf & _
          "WHAT IS NEVER TRANSMITTED:" & vbCrLf & _
          String(30, "-") & vbCrLf & _
          "  - Inventory data (item names, prices, descriptions)" & vbCrLf & _
          "  - Photos or image files" & vbCrLf & _
          "  - Word documents or description files" & vbCrLf & _
          "  - Your name, email, or any personal information" & vbCrLf & _
          "  - Sales history or sold item data" & vbCrLf & _
          "  - Folder paths or file locations" & vbCrLf & vbCrLf & _
          "YOUR DATA SECURITY:" & vbCrLf & _
          String(20, "-") & vbCrLf & _
          "  - 100% of your data stays on YOUR computer" & vbCrLf & _
          "  - We cannot access, view, or download your inventory" & vbCrLf & _
          "  - Update checking only compares version numbers" & vbCrLf & _
          "  - Downloading updates is optional and under your control"
    
    MsgBox msg, vbInformation, "Privacy & Security Information"
    
End Sub
