Attribute VB_Name = "modLogging"
Option Explicit

' =====================================================
' CENTRAL ERROR HANDLER
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

