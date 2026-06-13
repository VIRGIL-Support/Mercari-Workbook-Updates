Attribute VB_Name = "modHelp"
' =====================================================
' HELP SYSTEM - Auto-create help files and handle links
' =====================================================

' Help folder paths
Public Const HELP_FOLDER As String = "Help"
Public Const QUICKSTART_FOLDER As String = "VirgilQuickStartMercari"
Public Const USERGUIDE_FOLDER As String = "VirgilUserGuideMercari"

Public Const FILENAME_QUICKSTART As String = "VirgilQuickStartMercari.pdf"
Public Const FILENAME_USERGUIDE_INDEX As String = "index.html"

' =====================================================
' HELP FOLDER & FILE CREATION
' =====================================================

' Create help folder structure and default files
Public Sub CreateHelpStructure()
    Dim wbPath As String
    Dim helpPath As String
    Dim quickstartPath As String
    Dim userguidePath As String
    
    wbPath = ThisWorkbook.Path
    helpPath = wbPath & "\" & HELP_FOLDER
    quickstartPath = helpPath & "\" & QUICKSTART_FOLDER
    userguidePath = helpPath & "\" & USERGUIDE_FOLDER
    
    ' Create folders if missing
    CreateFolderIfMissing helpPath
    CreateFolderIfMissing quickstartPath
    CreateFolderIfMissing userguidePath
    
    ' Create placeholder files if missing
    CreateDefaultQuickStartIfMissing quickstartPath & "\" & FILENAME_QUICKSTART
    CreateDefaultUserGuideIfMissing userguidePath & "\" & FILENAME_USERGUIDE_INDEX
End Sub

Private Sub CreateDefaultQuickStartIfMissing(ByVal filePath As String)
    If Dir(filePath) <> "" Then Exit Sub
    
    ' Create a placeholder text file explaining where to put the PDF
    Dim f As Integer
    f = FreeFile
    Open filePath & ".placeholder.txt" For Output As #f
    Print #f, "QUICK START GUIDE PLACEHOLDER"
    Print #f, "=============================="
    Print #f, ""
    Print #f, "Please place the VirgilQuickStartMercari.pdf file in this folder:"
    Print #f, filePath
    Print #f, ""
    Print #f, "Once the PDF is in place, delete this placeholder file."
    Close #f
End Sub

Private Sub CreateDefaultUserGuideIfMissing(ByVal filePath As String)
    If Dir(filePath) <> "" Then Exit Sub
    
    ' Create a placeholder index.html
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, "<!DOCTYPE html>"
    Print #f, "<html><head><title>Virgil User Guide - Mercari Workbook</title></head>"
    Print #f, "<body style='font-family:Arial,sans-serif; max-width:800px; margin:50px auto; padding:20px;'>"
    Print #f, "<h1>Virgil User Guide - Mercari Workbook</h1>"
    Print #f, "<p><strong>Version:</strong> " & APP_VERSION & "</p>"
    Print #f, "<p><strong>Support:</strong> <a href='mailto:" & SUPPORT_EMAIL & "'>" & SUPPORT_EMAIL & "</a></p>"
    Print #f, "<hr>"
    Print #f, "<p>This is a placeholder for the User Guide. Please replace this file with the actual user guide content.</p>"
    Print #f, "<p>Location: " & filePath & "</p>"
    Print #f, "</body></html>"
    Close #f
End Sub

' =====================================================
' HELP LINK HANDLERS (Assign these to buttons/shapes)
' =====================================================

' Open Quick Start PDF
Public Sub OpenQuickStartGuide()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & HELP_FOLDER & "\" & QUICKSTART_FOLDER & "\" & FILENAME_QUICKSTART
    
    If Dir(filePath) = "" Then
        MsgBox "Quick Start Guide not found at:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
               "Please ensure the PDF is in the correct location.", vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink filePath
    If Err.Number <> 0 Then
        Shell "explorer.exe """ & filePath & """", vbNormalFocus
    End If
    On Error GoTo 0
End Sub

' Open User Guide HTML
Public Sub OpenUserGuide()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & HELP_FOLDER & "\" & USERGUIDE_FOLDER & "\" & FILENAME_USERGUIDE_INDEX
    
    If Dir(filePath) = "" Then
        MsgBox "User Guide not found at:" & vbCrLf & filePath, vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink filePath
    If Err.Number <> 0 Then
        Shell "explorer.exe """ & filePath & """", vbNormalFocus
    End If
    On Error GoTo 0
End Sub

' Placeholder for Video Tutorial
Public Sub OpenVideoTutorial()
    MsgBox "Video Tutorial coming soon!" & vbCrLf & vbCrLf & _
           "This feature will be available in a future update.", vbInformation, "Coming Soon"
    ' When ready, replace with:
    ' ActiveWorkbook.FollowHyperlink "https://your-video-url-here"
End Sub

' Open EULA
Public Sub OpenEULAFromHelp()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\EULA_Full.txt"
    
    If Dir(filePath) = "" Then
        MsgBox "EULA file not found.", vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "notepad.exe """ & filePath & """", vbNormalFocus
    If Err.Number <> 0 Then
        ActiveWorkbook.FollowHyperlink filePath
    End If
    On Error GoTo 0
End Sub

' Open Troubleshooting Guide
Public Sub OpenTroubleshootingFromHelp()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\TROUBLESHOOTING_GUIDE.txt"
    
    If Dir(filePath) = "" Then
        MsgBox "Troubleshooting Guide not found.", vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "notepad.exe """ & filePath & """", vbNormalFocus
    If Err.Number <> 0 Then
        ActiveWorkbook.FollowHyperlink filePath
    End If
    On Error GoTo 0
End Sub

' Open Version History
Public Sub OpenVersionHistoryFromHelp()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\VERSION_HISTORY.txt"
    
    If Dir(filePath) = "" Then
        MsgBox "Version History not found.", vbExclamation, "File Not Found"
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "notepad.exe """ & filePath & """", vbNormalFocus
    If Err.Number <> 0 Then
        ActiveWorkbook.FollowHyperlink filePath
    End If
    On Error GoTo 0
End Sub

' Email Virgil Support
Public Sub EmailVIRGILSupport()
    Dim mailto As String
    Dim subject As String
    
    subject = "Mercari Workbook Support - v" & APP_VERSION
    mailto = "mailto:" & SUPPORT_EMAIL & "?subject=" & subject
    
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink mailto
    If Err.Number <> 0 Then
        MsgBox "Please email:" & vbCrLf & vbCrLf & SUPPORT_EMAIL & vbCrLf & vbCrLf & _
               "Subject: " & subject, vbInformation, "Email Support"
    End If
    On Error GoTo 0
End Sub

' =====================================================
' DYNAMIC COPYRIGHT YEAR
' =====================================================

' Update copyright year in specified cell
Public Sub UpdateCopyrightYear(ByVal wsName As String, ByVal cellAddress As String)
    Dim ws As Worksheet
    Dim currentYear As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    currentYear = Format(Now, "yyyy")
    ws.Range(cellAddress).Value = "© " & currentYear & " VIRGIL Support. All rights reserved."
End Sub

' Auto-update copyright on startup
Public Sub RefreshCopyright()
    UpdateCopyrightYear "HELP", "B37"
End Sub

' =====================================================
' DEVELOPER BUTTONS: Hide/Unhide All Worksheets
' =====================================================

Public Sub CreateUnhideButton()
    ' Creates buttons at A100 and A101 on HELP worksheet for testing
    
    Dim ws As Worksheet
    Dim btn As Shape
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("HELP")
    If ws Is Nothing Then Exit Sub
    
    ' Delete existing buttons if present
    ws.Shapes("BTN_UNHIDE_ALL").Delete
    ws.Shapes("BTN_HIDE_ALL").Delete
    
    ' Button dimensions
    btnLeft = ws.Range("A100").Left + 5
    btnWidth = ws.Range("A100:B100").Width - 10
    
    ' Create UNHIDE button at A100
    btnTop = ws.Range("A100").Top + 3
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, 20)
    
    With btn
        .Name = "BTN_UNHIDE_ALL"
        .TextFrame.Characters.Text = "UNHIDE ALL WORKSHEETS"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.ForeColor.RGB = RGB(92, 127, 168)  ' Blue
        .Line.ForeColor.RGB = RGB(61, 95, 130)
        .OnAction = "UnhideAllWorksheets"
    End With
    
    ' Create HIDE button at A101
    btnTop = ws.Range("A101").Top + 3
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, 20)
    
    With btn
        .Name = "BTN_HIDE_ALL"
        .TextFrame.Characters.Text = "HIDE ALL WORKSHEETS"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .Fill.ForeColor.RGB = RGB(61, 95, 130)  ' Darker blue
        .Line.ForeColor.RGB = RGB(41, 70, 100)
        .OnAction = "HideAllWorksheetsForTesting"
    End With
    
    On Error GoTo 0
    
End Sub

' =====================================================
' HIDE ALL WORKSHEETS (for testing)
' =====================================================

Public Sub HideAllWorksheetsForTesting()
    ' For development use - rehides all system worksheets
    
    On Error Resume Next
    
    ThisWorkbook.Worksheets("DATA").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("LOOKUPS").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("SETTINGS").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("TABLES").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("COPYRIGHT_INFO").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("SOLD_DATA").Visible = xlSheetVeryHidden
    
    On Error GoTo 0
    
    MsgBox "All system worksheets are now hidden.", vbInformation
    
End Sub

' ============================================
' FULL RESET FOR FRESH TESTING
' ============================================

Public Sub Reset_1_Full_Reset()
    ' Combines ResetFirstRunSettings + ResetWorkbookForFreshTesting
    ' Clears ALL first-run flags AND all inventory/data for a completely clean slate
    
    Dim confirmResult As VbMsgBoxResult
    
    confirmResult = MsgBox("FULL RESET - This will clear EVERYTHING:" & vbCrLf & vbCrLf & _
                          "FIRST-RUN SETTINGS:" & vbCrLf & _
                          "  - EULA acceptance status" & vbCrLf & _
                          "  - Update preference choice" & vbCrLf & _
                          "  - First-run completion flag" & vbCrLf & vbCrLf & _
                          "WORKBOOK DATA:" & vbCrLf & _
                          "  - All inventory rows and item data" & vbCrLf & _
                          "  - All sold items records" & vbCrLf & _
                          "  - All item editor field data" & vbCrLf & vbCrLf & _
                          "The workbook will behave as if opened for the very first time." & vbCrLf & vbCrLf & _
                          "THIS CANNOT BE UNDONE. Continue?", _
                          vbYesNo + vbCritical, "1 Full Reset - Confirm")
    
    If confirmResult <> vbYes Then
        MsgBox "Reset cancelled. No changes were made.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    
    ' ---- PART 1: Reset first-run settings ----
    UpdateSetting "FIRST_RUN_COMPLETE", ""
    UpdateSetting "AUTO_CHECK_UPDATES", ""
    UpdateSetting "EULA_ACCEPTED", ""
    UpdateSetting "EULA_ACCEPTED_DATE", ""
    
    ' ---- PART 2: Clear inventory worksheet ----
    Dim wsInv As Worksheet
    Set wsInv = ThisWorkbook.Worksheets("INVENTORY")
    If Not wsInv Is Nothing Then
        Dim lastInvRow As Long
        lastInvRow = wsInv.Cells(wsInv.Rows.Count, 1).End(xlUp).Row
        If lastInvRow > 1 Then
            wsInv.Rows("2:" & lastInvRow).Delete
        End If
    End If
    
    ' ---- PART 3: Clear DATA sheet ----
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Worksheets("DATA")
    If Not wsData Is Nothing Then
        Dim lastDataRow As Long
        lastDataRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
        If lastDataRow > 1 Then
            wsData.Rows("2:" & lastDataRow).ClearContents
        End If
    End If
    
    ' ---- PART 4: Clear SOLD_ITEMS worksheet ----
    Dim wsSold As Worksheet
    Set wsSold = ThisWorkbook.Worksheets("SOLD_ITEMS")
    If Not wsSold Is Nothing Then
        Dim lastSoldRow As Long
        lastSoldRow = wsSold.Cells(wsSold.Rows.Count, 1).End(xlUp).Row
        If lastSoldRow > 1 Then
            wsSold.Rows("2:" & lastSoldRow).Delete
        End If
    End If
    
    ' ---- PART 5: Clear SOLD_DATA sheet ----
    Dim wsSoldData As Worksheet
    Set wsSoldData = ThisWorkbook.Worksheets("SOLD_DATA")
    If Not wsSoldData Is Nothing Then
        Dim lastSoldDataRow As Long
        lastSoldDataRow = wsSoldData.Cells(wsSoldData.Rows.Count, 1).End(xlUp).Row
        If lastSoldDataRow > 1 Then
            wsSoldData.Rows("2:" & lastSoldDataRow).ClearContents
        End If
    End If
    
    On Error GoTo 0
    Application.ScreenUpdating = True
    
    MsgBox "Full Reset Complete!" & vbCrLf & vbCrLf & _
           "All inventory data and first-run settings have been cleared." & vbCrLf & vbCrLf & _
           "Close and reopen the workbook to experience the complete first-time setup flow:" & vbCrLf & _
           "  - EULA display" & vbCrLf & _
           "  - Update preference prompt" & vbCrLf & _
           "  - Clean inventory ready for new items", vbInformation, "1 Full Reset Complete"

End Sub

' ============================================
' RESET FIRST-RUN SETTINGS (FOR TESTING)
' ============================================

Public Sub ResetFirstRunSettings()
    ' Resets first-run flags so you can test the initial user experience
    ' This is for development/testing purposes only
    
    Dim confirmResult As VbMsgBoxResult
    
    confirmResult = MsgBox("This will reset ALL first-run settings including:" & vbCrLf & vbCrLf & _
                          "• EULA acceptance status" & vbCrLf & _
                          "• Update preference choice" & vbCrLf & _
                          "• First-run completion flag" & vbCrLf & vbCrLf & _
                          "The workbook will behave as if opened for the very first time." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNo + vbExclamation, "Reset First-Run Settings")
    
    If confirmResult <> vbYes Then Exit Sub
    
    On Error Resume Next
    
    ' Clear first-run related settings
    UpdateSetting "FIRST_RUN_COMPLETE", ""
    UpdateSetting "AUTO_CHECK_UPDATES", ""
    UpdateSetting "EULA_ACCEPTED", ""
    UpdateSetting "EULA_ACCEPTED_DATE", ""
    
    On Error GoTo 0
    
    MsgBox "First-run settings have been RESET." & vbCrLf & vbCrLf & _
           "Close and reopen the workbook to experience the first-time setup flow:" & vbCrLf & vbCrLf & _
           "• Macro security info" & vbCrLf & _
           "• EULA display" & vbCrLf & _
           "• Update preference prompt (with privacy info)" & vbCrLf & _
           "• Initial setup completion", vbInformation, "Reset Complete"
    
End Sub

' =====================================================
' UPDATE CHECKING FUNCTIONS
' =====================================================

' Check for Updates button handler
Public Sub CheckForUpdatesButtonClick()
    On Error Resume Next
    ManualCheckForUpdates
    On Error GoTo 0
End Sub

' Toggle Auto-Check Updates button handler
Public Sub ToggleAutoCheckButtonClick()
    Dim current As String
    Dim newState As String
    Dim msg As String
    
    current = GetSettingValue("AUTO_CHECK_UPDATES")
    
    If current = "NO" Or Trim$(current) = "" Then
        newState = "YES"
        UpdateSetting "AUTO_CHECK_UPDATES", "YES"
        msg = "Automatic update checking is now ENABLED." & vbCrLf & vbCrLf & _
              "VIRGIL will check for updates each time the workbook opens." & vbCrLf & vbCrLf & _
              "PRIVACY NOTE: Only your version number is transmitted. " & _
              "No inventory data, photos, or personal information ever leaves your computer."
    Else
        newState = "NO"
        UpdateSetting "AUTO_CHECK_UPDATES", "NO"
        msg = "Automatic update checking is now DISABLED." & vbCrLf & vbCrLf & _
              "You can still check for updates manually using the 'Check for Updates' button."
    End If
    
    MsgBox msg, vbInformation, "Auto-Check Updates: " & IIf(newState = "YES", "ENABLED", "DISABLED")
End Sub
