Attribute VB_Name = "modVersion"
Option Explicit

' ============================================
' VERSION MANAGEMENT MODULE
' ============================================

Public Const CURRENT_VERSION As String = "1.0"
Public Const UPDATE_CHECK_URL As String = "https://raw.githubusercontent.com/YOURUSERNAME/YOURREPO/main/version.txt"
Public Const UPDATE_DOWNLOAD_URL As String = "https://raw.githubusercontent.com/YOURUSERNAME/YOURREPO/main/modules/"

' Check for updates on workbook open
Public Sub CheckForUpdatesOnOpen()
    Dim checkResult As String
    Dim latestVersion As String
    Dim userChoice As VbMsgBoxResult
    
    ' Check if user wants to auto-check (stored in settings)
    If GetSettingValue("AUTO_CHECK_UPDATES") = "NO" Then Exit Sub
    
    checkResult = GetLatestVersionInfo()
    
    If checkResult = "" Then
        ' No update available or couldn't check
        Exit Sub
    End If
    
    latestVersion = Trim$(checkResult)
    
    If IsNewerVersion(latestVersion, CURRENT_VERSION) Then
        userChoice = MsgBox("A new version (" & latestVersion & ") is available!" & vbCrLf & vbCrLf & _
                           "Current version: " & CURRENT_VERSION & vbCrLf & vbCrLf & _
                           "Would you like to update now?" & vbCrLf & vbCrLf & _
                           "Your data will be preserved.", vbYesNo + vbInformation, "Update Available")
        
        If userChoice = vbYes Then
            PerformUpdate latestVersion
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
' UPDATE PERFORMED - DOWNLOAD AND INSTALL
' ============================================

Private Sub PerformUpdate(ByVal newVersion As String)
    On Error GoTo ErrorHandler
    
    Dim modulesToUpdate As Variant
    Dim i As Long
    Dim downloadSuccess As Boolean
    Dim tempFolder As String
    Dim moduleCode As String
    
    ' List of modules to update (add your module names here)
    modulesToUpdate = Array( _
        "modAIExport.bas", _
        "modInventory.bas", _
        "modPhotos.bas", _
        "modUtilities.bas", _
        "modVersion.bas" _
    )
    
    ' Create temp folder for downloads
    tempFolder = Environ$("TEMP") & "\MercariWorkbookUpdate\"
    CreateFolderIfMissing tempFolder
    
    ' Show progress
    Application.StatusBar = "Downloading update..."
    
    ' Download each module
    For i = LBound(modulesToUpdate) To UBound(modulesToUpdate)
        moduleCode = DownloadModule(CStr(modulesToUpdate(i)))
        
        If moduleCode = "" Then
            MsgBox "Failed to download " & modulesToUpdate(i) & vbCrLf & _
                   "Update cancelled. Please try again later.", vbExclamation, "Update Failed"
            Application.StatusBar = False
            Exit Sub
        End If
        
        ' Save to temp file
        SaveTextToFile tempFolder & modulesToUpdate(i), moduleCode
        
        Application.StatusBar = "Downloaded " & (i + 1) & " of " & (UBound(modulesToUpdate) + 1) & " files..."
    Next i
    
    Application.StatusBar = False
    
    ' Confirm before installing
    If MsgBox("All files downloaded successfully!" & vbCrLf & vbCrLf & _
              "The workbook will now close and reopen with the update." & vbCrLf & _
              "Your data will be preserved." & vbCrLf & vbCrLf & _
              "Save any unsaved work now. Continue?", vbYesNo + vbQuestion, "Ready to Install") = vbNo Then
        Exit Sub
    End If
    
    ' Perform installation
    InstallUpdate tempFolder, newVersion
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error during update: " & Err.Number & " - " & Err.Description, vbCritical, "Update Error"
    
End Sub

' Download a single module from URL
Private Function DownloadModule(ByVal moduleName As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    
    url = UPDATE_DOWNLOAD_URL & moduleName
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send
    
    If http.Status = 200 Then
        DownloadModule = http.responseText
    Else
        DownloadModule = ""
    End If
    
    Exit Function
    
ErrorHandler:
    DownloadModule = ""
End Function

' Save text to file
Private Sub SaveTextToFile(ByVal filePath As String, ByVal content As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content;
    Close #fileNum
End Sub

' ============================================
' INSTALL UPDATE - REPLACE MODULES
' ============================================

Private Sub InstallUpdate(ByVal tempFolder As String, ByVal newVersion As String)
    On Error GoTo ErrorHandler
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim fso As Object
    Dim file As Object
    Dim folder As Object
    Dim moduleName As String
    Dim filePath As String
    
    Set vbProj = ThisWorkbook.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(tempFolder)
    
    ' Remove old modules (except protected ones)
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = 1 Then ' 1 = Standard Module
            moduleName = vbComp.Name
            ' Don't remove essential modules until we have replacement
            If moduleName <> "modVersion" And moduleName <> "ThisWorkbook" Then
                vbProj.VBComponents.Remove vbComp
            End If
        End If
    Next vbComp
    
    ' Import new modules
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "bas" Then
            filePath = file.Path
            vbProj.VBComponents.Import filePath
        End If
    Next file
    
    ' Update version number in Settings
    UpdateSetting "VERSION", newVersion
    
    ' Save and reopen
    ThisWorkbook.Save
    
    MsgBox "Update complete! Workbook will now reopen.", vbInformation, "Success"
    
    ' Reopen workbook
    Application.OnTime Now + TimeValue("00:00:02"), "ReopenWorkbook"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error installing update: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Your original workbook is still intact.", vbCritical, "Installation Error"
End Sub

' Reopen workbook after update
Public Sub ReopenWorkbook()
    Dim wbPath As String
    wbPath = ThisWorkbook.FullName
    
    ' Close this instance
    ThisWorkbook.Close SaveChanges:=False
    
    ' Reopen fresh
    Workbooks.Open wbPath
End Sub

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
