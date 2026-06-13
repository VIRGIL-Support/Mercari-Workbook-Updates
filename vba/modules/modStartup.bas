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
    
    ' Create Mercari signup button on Welcome sheet
    On Error Resume Next
    Call CreateMercariSignupButton
    On Error GoTo 0
    
    ' Display current version number on Welcome worksheet cell I5
    On Error Resume Next
    With ThisWorkbook.Worksheets(WS_WELCOME).Range("I5")
        .Value = "Version " & CURRENT_VERSION
        .Font.Name = "Atkinson Hyperlegible"
        .Font.Size = 12
        .Font.Color = RGB(238, 240, 253)
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    On Error GoTo 0
    
    ApplyColumnFormatting

    ' =====================================================
    ' FIRST RUN SETUP
    ' =====================================================

    If UCase(firstRunValue) <> "YES" Then

        ' Display EULA - opens Notepad behind popup, thank-you shown on acceptance
        If Not DisplayEULA() Then
            ' User declined - close workbook without saving
            MsgBox "You must accept the End User License Agreement (EULA) to use this software." & vbCrLf & vbCrLf & _
                   "The workbook will now close.", vbExclamation, "License Required"
            ThisWorkbook.Close SaveChanges:=False
            Exit Sub
        End If

        ' Set hidden watermark for copyright tracking
        Call SetHiddenWatermark

        ' Ask about automatic update checking (appears after EULA thank-you is dismissed)
        On Error Resume Next
        Call PromptForAutoUpdatePreference
        On Error GoTo 0

        ' MARK FIRST RUN COMPLETE
        Call SetSettingValue("FIRST_RUN_COMPLETE", "Yes")

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

' =====================================================
' WELCOME SHEET: Create Mercari Signup Button
' =====================================================

Public Sub CreateMercariSignupButton()
    ' Creates a proper button shape on Welcome worksheet row 36
    
    Dim ws As Worksheet
    Dim btn As Shape
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim fontColor As Long
    Dim borderColor As Long
    Dim bgColor As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_WELCOME)
    If ws Is Nothing Then Exit Sub
    
    ' Delete existing button if present
    ws.Shapes("BTN_MERCARI_SIGNUP").Delete
    
    ' Color specifications
    fontColor = RGB(94, 109, 242)      ' Font color
    borderColor = RGB(197, 203, 250)   ' Border color
    bgColor = RGB(94, 109, 242)        ' Background behind button
    
    ' Clear old text/hyperlink from the cells first (A36:K36)
    ws.Range("A36:K36").ClearContents
    ws.Range("A36:K36").Hyperlinks.Delete
    
    ' Set background color behind the button
    ws.Range("A36:K36").Interior.Color = bgColor
    
    ' Button dimensions
    btnHeight = 32
    btnWidth = 480
    
    ' Center horizontally between columns A and K
    Dim areaLeft As Double, areaWidth As Double
    areaLeft = ws.Range("A36").Left
    areaWidth = ws.Range("A36:K36").Width
    btnLeft = areaLeft + (areaWidth - btnWidth) / 2
    
    ' Center vertically between rows 34 and 39
    Dim sectionTop As Double, sectionHeight As Double
    sectionTop = ws.Rows(34).Top
    sectionHeight = ws.Range("34:39").Height
    btnTop = sectionTop + (sectionHeight - btnHeight) / 2
    
    ' Create rounded rectangle button
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    
    With btn
        .Name = "BTN_MERCARI_SIGNUP"
        .TextFrame.Characters.Text = "Click Here to Join Mercari and Claim Your $20 Bonus!"
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.MarginLeft = 5
        .TextFrame.MarginRight = 5
        .TextFrame.MarginTop = 3
        .TextFrame.MarginBottom = 3
        
        ' Atkinson Hyperlegible font at 16pt
        .TextFrame.Characters.Font.Name = "Atkinson Hyperlegible"
        .TextFrame.Characters.Font.Size = 16
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = fontColor
        
        ' White background with border color
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 2.5
        
        ' Shadow effect
        .Shadow.Visible = msoTrue
        .Shadow.Style = msoShadowStyleOuterShadow
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 2
        .Shadow.Blur = 4
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.5
        
        ' Assign click action
        .OnAction = "OpenMercariSignupLink"
    End With
    
    On Error GoTo 0
    
End Sub

' =====================================================
' Open Mercari Signup Link
' =====================================================

Public Sub OpenMercariSignupLink()
    ' Opens the Mercari signup link in default browser
    
    Dim signupURL As String
    signupURL = "https://www.mercari.com/invitations/?iv_code=NVCYKW"
    
    On Error Resume Next
    ActiveWorkbook.FollowHyperlink signupURL
    If Err.Number <> 0 Then
        ' Fallback if FollowHyperlink fails
        Shell "explorer.exe """ & signupURL & """", vbNormalFocus
    End If
    On Error GoTo 0
    
End Sub
