Attribute VB_Name = "modSupport"
' =====================================================
' USER SUPPORT FRAMEWORK
' =====================================================

' Support contact constants
Public Const SUPPORT_EMAIL As String = "VIRGIL_Support@proton.me"
Public Const SUPPORT_RESPONSE_TIME As String = "48 hours (typical, not guaranteed)"

' File locations
Public Const FILENAME_VERSION_HISTORY As String = "VERSION_HISTORY.txt"
Public Const FILENAME_TROUBLESHOOTING As String = "TROUBLESHOOTING_GUIDE.txt"
Public Const FILENAME_EULA As String = "EULA_Full.txt"

' =====================================================
' SUPPORT FUNCTIONS
' =====================================================

' Open troubleshooting guide
Public Sub OpenTroubleshootingGuide()
    Dim guidePath As String
    guidePath = ThisWorkbook.Path & "\" & FILENAME_TROUBLESHOOTING
    
    If Dir(guidePath) = "" Then
        MsgBox "Troubleshooting guide not found at:" & vbCrLf & guidePath & vbCrLf & vbCrLf & _
               "Please contact support at:" & vbCrLf & SUPPORT_EMAIL, vbExclamation, "Guide Not Found"
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "notepad.exe """ & guidePath & """", vbNormalFocus
    If Err.Number <> 0 Then
        ' Fallback
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run "cmd /c start "" """ & guidePath & """", 0, False
        Set wsh = Nothing
    End If
    On Error GoTo 0
End Sub

' Open version history
Public Sub OpenVersionHistory()
    Dim versionPath As String
    versionPath = ThisWorkbook.Path & "\" & FILENAME_VERSION_HISTORY
    
    If Dir(versionPath) = "" Then
        MsgBox "Version history not found." & vbCrLf & vbCrLf & _
               "Current version: " & APP_VERSION, vbInformation, "Version Info"
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "notepad.exe """ & versionPath & """", vbNormalFocus
    If Err.Number <> 0 Then
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run "cmd /c start "" """ & versionPath & """", 0, False
        Set wsh = Nothing
    End If
    On Error GoTo 0
End Sub

' Show support information dialog
Public Sub ShowSupportInfo()
    Dim msg As String
    
    msg = "SUPPORT INFORMATION" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & _
          "Support Email: " & SUPPORT_EMAIL & vbCrLf & vbCrLf & _
          "Response Time: " & SUPPORT_RESPONSE_TIME & vbCrLf & vbCrLf & _
          "Current Version: " & APP_VERSION & vbCrLf & vbCrLf & _
          "Resources Available:" & vbCrLf & _
          "- Troubleshooting Guide (TROUBLESHOOTING_GUIDE.txt)" & vbCrLf & _
          "- Version History (VERSION_HISTORY.txt)" & vbCrLf & _
          "- Error Logs (\\Logs\\ErrorLog.txt)" & vbCrLf & _
          "- EULA (EULA_Full.txt)" & vbCrLf & vbCrLf & _
          "Click 'Report Issue' on the Welcome page to submit a bug report."
    
    MsgBox msg, vbInformation, "Support Information"
End Sub

' Email support with pre-filled subject
Public Sub EmailSupport()
    Dim mailto As String
    Dim subject As String
    Dim body As String
    
    subject = "Mercari Workbook Support Request - v" & APP_VERSION
    body = "Hi VIRGIL Support,%0A%0A" & _
           "I need assistance with the Mercari Workbook.%0A%0A" & _
           "Issue Description:%0A" & _
           "[Please describe what happened]%0A%0A" & _
           "Steps to Reproduce:%0A" & _
           "1. %0A" & _
           "2. %0A" & _
           "3. %0A%0A" & _
           "Error Message (if any):%0A" & _
           "[Paste any error text here]%0A%0A" & _
           "Workbook Version: " & APP_VERSION & "%0A" & _
           "Excel Version: " & Application.Version & "%0A%0A" & _
           "I've attached the ErrorLog.txt file from the Logs folder.%0A%0A" & _
           "Thanks!"
    
    mailto = "mailto:" & SUPPORT_EMAIL & "?subject=" & subject & "&body=" & body
    
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink mailto
    If Err.Number <> 0 Then
        ' Fallback to showing email address
        MsgBox "Please email:" & vbCrLf & vbCrLf & SUPPORT_EMAIL & vbCrLf & vbCrLf & _
               "Subject: " & subject, vbInformation, "Email Support"
    End If
    On Error GoTo 0
End Sub

' Check if all support documents exist
Public Function VerifySupportDocuments() As Boolean
    Dim allFound As Boolean
    Dim missing As String
    
    allFound = True
    missing = ""
    
    If Dir(ThisWorkbook.Path & "\" & FILENAME_TROUBLESHOOTING) = "" Then
        allFound = False
        missing = missing & "- " & FILENAME_TROUBLESHOOTING & vbCrLf
    End If
    
    If Dir(ThisWorkbook.Path & "\" & FILENAME_VERSION_HISTORY) = "" Then
        allFound = False
        missing = missing & "- " & FILENAME_VERSION_HISTORY & vbCrLf
    End If
    
    If Dir(ThisWorkbook.Path & "\" & FILENAME_EULA) = "" Then
        allFound = False
        missing = missing & "- " & FILENAME_EULA & vbCrLf
    End If
    
    If Not allFound Then
        MsgBox "Warning: Some support documents are missing:" & vbCrLf & vbCrLf & _
               missing & vbCrLf & _
               "Please contact support if you need these documents.", vbExclamation, "Missing Documents"
    End If
    
    VerifySupportDocuments = allFound
End Function

' Create default support documents if missing
Public Sub CreateDefaultSupportDocuments()
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    
    ' Create EULA if missing
    If Dir(wbPath & "\" & FILENAME_EULA) = "" Then
        CreateDefaultEULA wbPath & "\" & FILENAME_EULA
    End If
    
    ' Create Version History if missing
    If Dir(wbPath & "\" & FILENAME_VERSION_HISTORY) = "" Then
        CreateDefaultVersionHistory wbPath & "\" & FILENAME_VERSION_HISTORY
    End If
    
    ' Create Troubleshooting Guide if missing
    If Dir(wbPath & "\" & FILENAME_TROUBLESHOOTING) = "" Then
        CreateDefaultTroubleshootingGuide wbPath & "\" & FILENAME_TROUBLESHOOTING
    End If
End Sub

Private Sub CreateDefaultEULA(ByVal filePath As String)
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, "END USER LICENSE AGREEMENT (EULA)"
    Print #f, ""
    Print #f, "IMPORTANT: PLEASE READ THIS AGREEMENT CAREFULLY BEFORE USING THIS SOFTWARE."
    Print #f, ""
    Print #f, "1. LICENSE GRANT"
    Print #f, "This software is licensed, not sold. The author grants you a single-user, non-exclusive, non-transferable license to use this Excel workbook and its associated VBA code (the ""Software"") on one computer for personal or business use."
    Print #f, ""
    Print #f, "2. SINGLE USER LICENSE"
    Print #f, "This license is valid for one (1) user only. Each additional user must purchase a separate license. The license may not be shared, sublicensed, or transferred to another user without prior written consent."
    Print #f, ""
    Print #f, "3. NO REDISTRIBUTION"
    Print #f, "You may not distribute, copy, publish, display, sublicense, or resell this Software or any modified version thereof to any third party. This includes sharing the workbook file, exporting VBA modules, or posting code online."
    Print #f, ""
    Print #f, "4. NO REVERSE ENGINEERING"
    Print #f, "You may not decompile, disassemble, reverse engineer, or attempt to derive the source code of this Software except to the extent that such activity is expressly permitted by applicable law notwithstanding this limitation."
    Print #f, ""
    Print #f, "5. INTELLECTUAL PROPERTY"
    Print #f, "All title, ownership rights, and intellectual property rights in and to the Software remain with the author. The Software is protected by copyright laws and international treaties."
    Print #f, ""
    Print #f, "6. DISCLAIMER OF LIABILITY"
    Print #f, "THE SOFTWARE IS PROVIDED ""AS IS"" WITHOUT WARRANTY OF ANY KIND. THE AUTHOR DISCLAIMS ALL WARRANTIES, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DAMAGES, INCLUDING WITHOUT LIMITATION, LOST PROFITS, LOST SAVINGS, OR ANY OTHER INCIDENTAL, CONSEQUENTIAL, SPECIAL, OR PUNITIVE DAMAGES ARISING OUT OF THE USE OR INABILITY TO USE THIS SOFTWARE, EVEN IF THE AUTHOR HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES."
    Print #f, ""
    Print #f, "7. DATA BACKUP"
    Print #f, "You are solely responsible for maintaining backup copies of any data entered into this Software. The author is not responsible for data loss or corruption."
    Print #f, ""
    Print #f, "8. SUPPORT"
    Print #f, "Support is provided via email at VIRGIL_Support@proton.me. Response time is typically within 48 hours but is not guaranteed."
    Print #f, ""
    Print #f, "9. UPDATES"
    Print #f, "This license includes free updates for the purchased version. Major version upgrades may require an additional license fee."
    Print #f, ""
    Print #f, "10. TERMINATION"
    Print #f, "This license is effective until terminated. Your rights under this license will terminate automatically without notice if you fail to comply with any term. Upon termination, you must cease all use of the Software and destroy all copies."
    Print #f, ""
    Print #f, "11. GOVERNING LAW"
    Print #f, "This agreement shall be governed by and construed in accordance with the laws of the United States."
    Print #f, ""
    Print #f, "12. ENTIRE AGREEMENT"
    Print #f, "This agreement constitutes the entire agreement between you and the author regarding the Software and supersedes all prior agreements."
    Print #f, ""
    Print #f, "BY USING THIS SOFTWARE, YOU ACKNOWLEDGE THAT YOU HAVE READ THIS AGREEMENT, UNDERSTAND IT, AND AGREE TO BE BOUND BY ITS TERMS AND CONDITIONS."
    Close #f
End Sub

Private Sub CreateDefaultVersionHistory(ByVal filePath As String)
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, "MERCARI WORKBOOK - VERSION HISTORY"
    Print #f, "===================================="
    Print #f, "Current Version: " & APP_VERSION
    Print #f, "Release Date: " & Format(Now, "yyyy-mm-dd")
    Print #f, ""
    Print #f, "VERSION NUMBERING SCHEME"
    Print #f, "========================"
    Print #f, "Format: MAJOR.MINOR.PATCH (e.g., 1.2.3)"
    Print #f, "- MAJOR: Significant new features or major changes"
    Print #f, "- MINOR: New functionality, backward compatible"
    Print #f, "- PATCH: Bug fixes and small improvements"
    Print #f, ""
    Print #f, "RELEASE HISTORY"
    Print #f, "==============="
    Print #f, ""
    Print #f, "Version " & APP_VERSION & " (" & Format(Now, "yyyy-mm-dd") & ")"
    Print #f, "--------------------------"
    Print #f, "Initial Release"
    Print #f, "- Complete inventory management system"
    Print #f, "- Sold items tracking"
    Print #f, "- Photo management with drag-and-drop"
    Print #f, "- AI-powered description generation"
    Print #f, "- Settings and configuration panel"
    Print #f, "- Backup and archive functionality"
    Print #f, "- EULA acceptance on first run"
    Print #f, "- Error logging and reporting system"
    Print #f, ""
    Print #f, "UPCOMING FEATURES (Planned)"
    Print #f, "==========================="
    Print #f, "- Auto-update mechanism"
    Print #f, "- Enhanced reporting dashboard"
    Print #f, "- Bulk import/export functionality"
    Print #f, ""
    Print #f, "KNOWN ISSUES"
    Print #f, "============"
    Print #f, "- None reported for v" & APP_VERSION
    Print #f, ""
    Print #f, "Please report any issues to: " & SUPPORT_EMAIL
    Close #f
End Sub

Private Sub CreateDefaultTroubleshootingGuide(ByVal filePath As String)
    Dim f As Integer
    f = FreeFile
    Open filePath For Output As #f
    Print #f, "MERCARI WORKBOOK - TROUBLESHOOTING GUIDE"
    Print #f, "========================================"
    Print #f, "Support Email: " & SUPPORT_EMAIL
    Print #f, "Response Time: Typically within 48 hours (not guaranteed)"
    Print #f, ""
    Print #f, "================================"
    Print #f, "COMMON ISSUES AND SOLUTIONS"
    Print #f, "================================"
    Print #f, ""
    Print #f, "1. WORKBOOK WON'T OPEN"
    Print #f, "----------------------"
    Print #f, "Symptom: Excel shows an error when opening the workbook"
    Print #f, ""
    Print #f, "Solutions:"
    Print #f, "a) Check if macros are enabled:"
    Print #f, "   - File > Options > Trust Center > Trust Center Settings"
    Print #f, "   - Macro Settings > Enable all macros (or Disable all macros with notification)"
    Print #f, "   - Restart Excel and click Enable Content when prompted"
    Print #f, ""
    Print #f, "b) Try opening in Safe Mode:"
    Print #f, "   - Close all Excel windows"
    Print #f, "   - Press Windows Key + R, type: excel /safe"
    Print #f, "   - Try opening the workbook"
    Print #f, ""
    Print #f, "c) Check for Excel updates:"
    Print #f, "   - File > Account > Update Options > Update Now"
    Print #f, ""
    Print #f, "2. BUTTONS NOT WORKING"
    Print #f, "----------------------"
    Print #f, "Symptom: Clicking buttons does nothing"
    Print #f, ""
    Print #f, "Solutions:"
    Print #f, "a) Ensure macros are enabled (see Issue #1)"
    Print #f, ""
    Print #f, "b) Check if workbook is in Protected View:"
    Print #f, "   - Look for yellow bar at top saying PROTECTED VIEW"
    Print #f, "   - Click Enable Editing"
    Print #f, ""
    Print #f, "c) Restart Excel and try again"
    Print #f, ""
    Print #f, "d) Check ErrorLog.txt in the Logs folder for error details"
    Print #f, ""
    Print #f, "================================"
    Print #f, "REPORTING A BUG"
    Print #f, "================================"
    Print #f, ""
    Print #f, "If you encounter an issue not listed above:"
    Print #f, ""
    Print #f, "1. Use the Report an Issue button (if available)"
    Print #f, "2. Or email: " & SUPPORT_EMAIL
    Print #f, ""
    Print #f, "Include the following information:"
    Print #f, "- What you were trying to do"
    Print #f, "- What actually happened"
    Print #f, "- Any error messages you saw"
    Print #f, "- The ErrorLog.txt file from the Logs folder"
    Print #f, ""
    Print #f, "================================"
    Print #f, "BACKUP RECOMMENDATIONS"
    Print #f, "================================"
    Print #f, ""
    Print #f, "To prevent data loss:"
    Print #f, "1. Automatic backups are created daily in the Backups folder"
    Print #f, "2. For extra safety, periodically save a copy to cloud storage"
    Print #f, "3. Before major operations, manually create a backup copy"
    Close #f
End Sub

' Log version update to history
Public Sub LogVersionUpdate(ByVal oldVersion As String, ByVal newVersion As String)
    Dim historyPath As String
    Dim fileNumber As Integer
    
    On Error Resume Next
    historyPath = ThisWorkbook.Path & "\" & FILENAME_VERSION_HISTORY
    fileNumber = FreeFile
    
    ' Append to version history
    Open historyPath For Append As #fileNumber
    Print #fileNumber, ""
    Print #fileNumber, "Version " & newVersion & " (" & Format(Now, "yyyy-mm-dd") & ")"
    Print #fileNumber, String(50, "-")
    Print #fileNumber, "Updated from version " & oldVersion
    Print #fileNumber, ""
    Close #fileNumber
    
    On Error GoTo 0
End Sub

' Get full version string with app name
Public Function GetFullVersionString() As String
    GetFullVersionString = APP_NAME & " v" & APP_VERSION
End Function

' Display about dialog
Public Sub ShowAboutDialog()
    Dim msg As String
    
    msg = APP_NAME & vbCrLf & String(50, "=") & vbCrLf & vbCrLf & _
          "Version: " & APP_VERSION & vbCrLf & vbCrLf & _
          "A comprehensive inventory management system" & vbCrLf & _
          "for Mercari sellers." & vbCrLf & vbCrLf & _
          "Support: " & SUPPORT_EMAIL & vbCrLf & vbCrLf & _
          "By clicking OK, you acknowledge that you have" & vbCrLf & _
          "read and accept the End User License Agreement."
    
    MsgBox msg, vbInformation, "About " & APP_NAME
End Sub
