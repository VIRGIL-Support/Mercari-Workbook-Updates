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

    rootFolder = ThisWorkbook.Path
    Set wsSettings = ThisWorkbook.Worksheets(WS_SETTINGS)
    wasSettingsProtected = wsSettings.ProtectContents
    If wasSettingsProtected Then wsSettings.Unprotect Password:=""

    firstRunValue = GetSettingValue("FIRST_RUN_COMPLETE")
    EnsureWorkbookFolders rootFolder
    ApplyColumnFormatting

    ' =====================================================
    ' FIRST RUN SETUP
    ' =====================================================

    If UCase(firstRunValue) <> "YES" Then

        ' MARK FIRST RUN COMPLETE
        Call SetSettingValue("FIRST_RUN_COMPLETE", "Yes")

        MsgBox "Initial setup completed successfully.", vbInformation

    Else

        MsgBox "Workbook startup completed.", vbInformation

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
