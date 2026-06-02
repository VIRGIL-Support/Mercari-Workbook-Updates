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
' RECOMMENDED TOOLS & AFFILIATE LINKS
' =====================================================

' Open the Recommended Tools document (markdown or HTML version)
Public Sub OpenRecommendedTools()
    Dim helpPath As String
    Dim toolsFile As String
    Dim htmlFile As String
    
    helpPath = ThisWorkbook.Path & "\HELP DOCS"
    
    ' Check for HTML version first (better user experience)
    htmlFile = helpPath & "\Recommended Tools & Affiliate Opportunities.html"
    toolsFile = helpPath & "\Recommended Tools & Affiliate Opportunities.md"
    
    On Error Resume Next
    
    ' Try HTML version first
    If Dir(htmlFile) <> "" Then
        ThisWorkbook.FollowHyperlink htmlFile
        Exit Sub
    End If
    
    ' Try markdown version
    If Dir(toolsFile) <> "" Then
        ' Open with default application (usually Notepad or markdown viewer)
        ThisWorkbook.FollowHyperlink toolsFile
        Exit Sub
    End If
    
    ' Try alternative filename without spaces
    toolsFile = helpPath & "\Recommended Tools and Affiliate Opportunities.md"
    If Dir(toolsFile) <> "" Then
        ThisWorkbook.FollowHyperlink toolsFile
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' If neither found, show message
    MsgBox "Recommended Tools document not found." & vbCrLf & vbCrLf & _
           "Expected location:" & vbCrLf & _
           helpPath & "\Recommended Tools & Affiliate Opportunities.md" & vbCrLf & vbCrLf & _
           "Please ensure the file exists in the HELP DOCS folder.", _
           vbExclamation, "File Not Found"
End Sub

' Open the Quick Start Guide
Public Sub OpenQuickStartGuide()
    Dim helpPath As String
    Dim quickFile As String
    
    helpPath = ThisWorkbook.Path & "\HELP DOCS"
    quickFile = helpPath & "\Quick Start.md"
    
    On Error Resume Next
    
    If Dir(quickFile) <> "" Then
        ThisWorkbook.FollowHyperlink quickFile
    Else
        MsgBox "Quick Start Guide not found in HELP DOCS folder.", vbExclamation
    End If
    
    On Error GoTo 0
End Sub

' Open the Full User Guide
Public Sub OpenUserGuide()
    Dim helpPath As String
    Dim guideFile As String
    
    helpPath = ThisWorkbook.Path & "\HELP DOCS"
    guideFile = helpPath & "\User Guide.md"
    
    On Error Resume Next
    
    If Dir(guideFile) <> "" Then
        ThisWorkbook.FollowHyperlink guideFile
    Else
        MsgBox "User Guide not found in HELP DOCS folder.", vbExclamation
    End If
    
    On Error GoTo 0
End Sub

' Open the FAQ
Public Sub OpenFAQ()
    Dim helpPath As String
    Dim faqFile As String
    
    helpPath = ThisWorkbook.Path & "\HELP DOCS"
    faqFile = helpPath & "\FAQ.md"
    
    On Error Resume Next
    
    If Dir(faqFile) <> "" Then
        ThisWorkbook.FollowHyperlink faqFile
    Else
        MsgBox "FAQ not found in HELP DOCS folder.", vbExclamation
    End If
    
    On Error GoTo 0
End Sub

' Open EULA document
Public Sub OpenEULA()
    Dim helpPath As String
    Dim eulaFile As String
    
    helpPath = ThisWorkbook.Path & "\HELP DOCS"
    eulaFile = helpPath & "\EULA.txt"
    
    On Error Resume Next
    
    If Dir(eulaFile) <> "" Then
        ThisWorkbook.FollowHyperlink eulaFile
    Else
        MsgBox "EULA not found in HELP DOCS folder.", vbExclamation
    End If
    
    On Error GoTo 0
End Sub
