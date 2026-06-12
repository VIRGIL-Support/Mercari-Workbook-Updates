Attribute VB_Name = "modLogging"
Option Explicit

' =====================================================
' CENTRAL ERROR HANDLER & ENHANCED LOGGING
' =====================================================

Public Sub HandleError( _
    ByVal procedureName As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String)

    Dim logPath As String

    WriteErrorLog procedureName, errorNumber, errorDescription
    logPath = GetErrorLogPath()

    MsgBox _
        "Looks like I made a goof with the script." & vbCrLf & vbCrLf & _
        "Here are some tips:" & vbCrLf & _
        "1 - If you can, save and close the workbook." & vbCrLf & _
        "2 - Reopen the workbook and try again." & vbCrLf & _
        "3 - If the same error occurs again, please send the troubleshooting log for review." & vbCrLf & vbCrLf & _
        "A troubleshooting log was saved here:" & vbCrLf & _
        logPath & vbCrLf & vbCrLf & _
        "Technical details:" & vbCrLf & _
        "Procedure: " & procedureName & vbCrLf & _
        "Error " & errorNumber & ": " & errorDescription, vbCritical

End Sub

' Enhanced error logging with full context for user-reported issues
Public Sub WriteDetailedErrorLog( _
    ByVal procedureName As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String, _
    Optional ByVal userAction As String = "")

    Dim fileNumber As Integer
    Dim logPath As String
    Dim contextInfo As String

    On Error Resume Next
    logPath = GetErrorLogPath()
    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
    
    Print #fileNumber, "=================================================="
    Print #fileNumber, "Timestamp: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNumber, "Procedure: " & procedureName
    Print #fileNumber, "Error Number: " & CStr(errorNumber)
    Print #fileNumber, "Error Description: " & errorDescription
    
    ' Additional context for debugging
    Print #fileNumber, "--- Application State ---"
    Print #fileNumber, "Workbook: " & ThisWorkbook.FullName
    Print #fileNumber, "Excel Version: " & Application.VERSION
    Print #fileNumber, "ScreenUpdating: " & Application.ScreenUpdating
    Print #fileNumber, "Calculation: " & Application.Calculation
    
    ' Sheet and cell context
    On Error Resume Next
    Print #fileNumber, "ActiveSheet: " & ActiveSheet.Name
    Print #fileNumber, "ActiveCell: " & ActiveCell.Address
    Print #fileNumber, "SelectedRange: " & Selection.Address
    On Error GoTo 0
    
    ' User context
    Print #fileNumber, "--- User Context ---"
    Print #fileNumber, "User: " & Environ$("USERNAME")
    Print #fileNumber, "Computer: " & Environ$("COMPUTERNAME")
    Print #fileNumber, "OS: " & Environ$("OS")
    
    ' User action if provided
    If Trim$(userAction) <> "" Then
        Print #fileNumber, "--- User Action ---"
        Print #fileNumber, "Action: " & userAction
    End If
    
    Print #fileNumber, ""
    Close #fileNumber

End Sub

' Log user actions for troubleshooting workflows
Public Sub LogUserAction(ByVal actionDescription As String)
    Dim fileNumber As Integer
    Dim logPath As String

    On Error Resume Next
    logPath = GetLogFolderPath() & "\UserActions.txt"
    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
    Print #fileNumber, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & Environ$("USERNAME") & " | " & actionDescription
    Close #fileNumber
End Sub

' User-friendly manual issue reporting
Public Sub ReportIssue()
    Dim userDescription As String
    Dim logPath As String
    Dim fileNumber As Integer
    Dim userResponse As VbMsgBoxResult
    
    ' Ask user what went wrong
    userDescription = InputBox("Please describe what happened or what went wrong:" & vbCrLf & vbCrLf & _
                               "(Example: 'I clicked the SELL button but nothing happened')" & vbCrLf & vbCrLf & _
                               "Your feedback helps improve the workbook!", _
                               "Report an Issue", "")
    
    ' Cancelled or empty
    If Trim$(userDescription) = "" Then
        Exit Sub
    End If
    
    ' Log the user report with full context
    On Error Resume Next
    logPath = GetLogFolderPath() & "\UserReports.txt"
    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
    
    Print #fileNumber, "=================================================="
    Print #fileNumber, "USER REPORT - " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNumber, "User: " & Environ$("USERNAME")
    Print #fileNumber, "Computer: " & Environ$("COMPUTERNAME")
    Print #fileNumber, "--- User Description ---"
    Print #fileNumber, userDescription
    Print #fileNumber, "--- Application State ---"
    Print #fileNumber, "Workbook: " & ThisWorkbook.FullName
    Print #fileNumber, "Excel Version: " & Application.VERSION
    On Error Resume Next
    Print #fileNumber, "ActiveSheet: " & ActiveSheet.Name
    Print #fileNumber, "ActiveCell: " & ActiveCell.Address
    On Error GoTo 0
    Print #fileNumber, ""
    Close #fileNumber
    
    ' Confirm to user
    userResponse = MsgBox("Thank you! Your issue has been logged." & vbCrLf & vbCrLf & _
                          "Log saved to:" & vbCrLf & logPath & vbCrLf & vbCrLf & _
                          "Would you like to open the log folder to send the file?", _
                          vbYesNo + vbInformation, "Issue Reported")
    
    If userResponse = vbYes Then
        Shell "explorer.exe """ & GetLogFolderPath() & """", vbNormalFocus
    End If
End Sub

' View logs easily (for Welcome page or menu)
Public Sub ViewErrorLogs()
    Dim logFolder As String
    logFolder = GetLogFolderPath()
    
    If Dir(logFolder, vbDirectory) = "" Then
        MsgBox "No logs folder found yet. Errors will be logged to:" & vbCrLf & logFolder, vbInformation
        Exit Sub
    End If
    
    Shell "explorer.exe """ & logFolder & """", vbNormalFocus
End Sub

Public Function GetLogFolderPath() As String
    GetLogFolderPath = GetSettingValue("LOGS_FOLDER")
    If Trim$(GetLogFolderPath) = "" Then GetLogFolderPath = ThisWorkbook.Path & "\Logs"
    CreateFolderIfMissing GetLogFolderPath
End Function

Public Function GetErrorLogPath() As String
    GetErrorLogPath = GetLogFolderPath() & "\" & ERROR_LOG_FILE
End Function

Public Sub WriteErrorLog( _
    ByVal procedureName As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String)

    Dim fileNumber As Integer
    Dim logPath As String

    On Error Resume Next
    logPath = GetErrorLogPath()
    fileNumber = FreeFile
    Open logPath For Append As #fileNumber
    Print #fileNumber, "=================================================="
    Print #fileNumber, "Timestamp: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNumber, "Procedure: " & procedureName
    Print #fileNumber, "Error Number: " & CStr(errorNumber)
    Print #fileNumber, "Error Description: " & errorDescription
    Print #fileNumber, "Workbook: " & ThisWorkbook.FullName
    Print #fileNumber, "User: " & Environ$("USERNAME")
    Print #fileNumber, "Computer: " & Environ$("COMPUTERNAME")
    Print #fileNumber, ""
    Close #fileNumber

End Sub

