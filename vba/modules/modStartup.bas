Attribute VB_Name = "modStartup"
Option Explicit

' =====================================================
' INITIALIZE WORKBOOK
' =====================================================

Public Sub InitializeWorkbook()

    On Error GoTo ErrorHandler

    Dim rootFolder As String
    Dim firstRunValue As String
    Dim wsSettings As Worksheet
    Dim wasSettingsProtected As Boolean
    
    ' CHECK IF THIS IS AN UPDATE TRANSFER IN PROGRESS
    ' If temp file exists, new workbook was just opened for data transfer
    ' Skip ALL first-run logic (EULA, folders, etc.) - transfer will handle everything
    If Dir(Environ$("TEMP") & "\MercariUpdateSource.txt") <> "" Then
        ' This is an update scenario - exit immediately
        ' TransferMyData will run after this and handle the setup
        Exit Sub
    End If

    rootFolder = ThisWorkbook.Path
    Set wsSettings = ThisWorkbook.Worksheets(WS_SETTINGS)
    wasSettingsProtected = wsSettings.ProtectContents
    If wasSettingsProtected Then wsSettings.Unprotect Password:=""

    firstRunValue = GetSettingValue("FIRST_RUN_COMPLETE")
    EnsureWorkbookFolders rootFolder
    
    ' Create default support documents if missing (EULA, Version History, Troubleshooting Guide)
    On Error Resume Next
    Call CreateDefaultSupportDocuments
    On Error GoTo 0
    
    ' Create help folder structure (Quick Start, User Guide)
    On Error Resume Next
    Call CreateHelpStructure
    On Error GoTo 0
    
    ' Update dynamic copyright year
    On Error Resume Next
    Call RefreshCopyright
    On Error GoTo 0
    
    ApplyColumnFormatting

    ' =====================================================
    ' FIRST RUN SETUP
    ' =====================================================

    If UCase(firstRunValue) <> "YES" Then

        ' Display EULA and get user acceptance
        If Not DisplayEULA() Then
            ' User declined - close workbook without saving
            MsgBox "You must accept the license agreement to use this software." & vbCrLf & vbCrLf & _
                   "The workbook will now close.", vbExclamation, "License Required"
            ThisWorkbook.Close SaveChanges:=False
            Exit Sub
        End If

        ' Set hidden watermark for copyright tracking
        Call SetHiddenWatermark

        ' MARK FIRST RUN COMPLETE
        Call SetSettingValue("FIRST_RUN_COMPLETE", "Yes")

        MsgBox "Thank you for accepting the license agreement!" & vbCrLf & vbCrLf & _
               "Initial setup completed successfully.", vbInformation

    Else

        ' Still set watermark on subsequent opens (in case it was cleared)
        Call SetHiddenWatermark

        ' Only show startup message if there's no pending update transfer
        If Dir(Environ$("TEMP") & "\MercariUpdateSource.txt") = "" Then
            MsgBox "Workbook startup completed.", vbInformation
        End If

    End If

    If wasSettingsProtected Then wsSettings.Protect Password:="", UserInterfaceOnly:=True
    Exit Sub

ErrorHandler:

    On Error Resume Next
    If wasSettingsProtected Then wsSettings.Protect Password:="", UserInterfaceOnly:=True
    On Error GoTo 0
    Call HandleError("InitializeWorkbook", Err.Number, Err.Description)

End Sub

Public Sub EnsureWorkbookFolders(ByVal rootFolder As String)

    If Trim$(rootFolder) = "" Then Exit Sub

    Call SetSettingValue("ROOT_FOLDER", rootFolder)

    Call CreateFolderIfMissing(rootFolder & "\1 READY TO LIST")
    Call CreateFolderIfMissing(rootFolder & "\2 DESCRIPTION FILES")
    Call CreateFolderIfMissing(rootFolder & "\3 SOLD")
    Call CreateFolderIfMissing(rootFolder & "\Logs")
    Call CreateFolderIfMissing(rootFolder & "\Backups")

    Call SetSettingValue("PHOTOS_FOLDER", rootFolder & "\1 READY TO LIST")
    Call SetSettingValue("DOCX_FOLDER", rootFolder & "\2 DESCRIPTION FILES")
    Call SetSettingValue("SOLD_FOLDER", rootFolder & "\3 SOLD")
    Call SetSettingValue("LOGS_FOLDER", rootFolder & "\Logs")
    Call SetSettingValue("BACKUPS_FOLDER", rootFolder & "\Backups")

End Sub
