Attribute VB_Name = "modEULA"
' ============================================
' END USER LICENSE AGREEMENT (EULA) MODULE
' ============================================

Public Function DisplayEULA() As Boolean
    ' Returns True if user accepts, False if declined
    
    Dim eulaPath As String
    eulaPath = ThisWorkbook.Path & "\EULA_Full.txt"
    
    ' Step 1: Offer to open the EULA text file
    Dim viewResult As VbMsgBoxResult
    viewResult = MsgBox("VIRGIL Mercari Workbook - End User License Agreement (EULA)" & vbCrLf & vbCrLf & _
                        "Click YES to open and read the full EULA." & vbCrLf & vbCrLf & _
                        "Click NO to skip reading and proceed to acceptance.", _
                        vbYesNo + vbQuestion, "View End User License Agreement (EULA)?")
    
    If viewResult = vbYes Then
        On Error Resume Next
        If Dir(eulaPath) <> "" Then
            Shell "notepad.exe """ & eulaPath & """", vbNormalFocus
        Else
            MsgBox "EULA file not found. Please contact support at VIRGIL_Support@proton.me", _
                   vbExclamation, "File Not Found"
        End If
        On Error GoTo 0
        
        ' Small pause so Notepad opens before the next popup
        Application.Wait Now + TimeValue("00:00:02")
    End If
    
    ' Step 2: Ask for acceptance
    Dim acceptResult As VbMsgBoxResult
    acceptResult = MsgBox("Do you accept all terms of the End User License Agreement (EULA)?" & vbCrLf & vbCrLf & _
                          "Key terms:" & vbCrLf & _
                          "  - Single user license (no sharing)" & vbCrLf & _
                          "  - No redistribution or reverse engineering" & vbCrLf & _
                          "  - Software provided AS IS with no warranty" & vbCrLf & _
                          "  - Support via VIRGIL_Support@proton.me" & vbCrLf & vbCrLf & _
                          "You can review the End User License Agreement (EULA) at any time" & vbCrLf & _
                          "from the HELP worksheet." & vbCrLf & vbCrLf & _
                          "Click YES to accept and continue." & vbCrLf & _
                          "Click NO to decline and close the workbook.", _
                          vbYesNo + vbQuestion + vbDefaultButton1, "End User License Agreement (EULA)")
    
    If acceptResult = vbYes Then
        MsgBox "Thank you for accepting the End User License Agreement (EULA)!" & vbCrLf & vbCrLf & _
               "You can review it at any time from the HELP worksheet.", _
               vbInformation, "EULA Accepted"
        DisplayEULA = True
    Else
        DisplayEULA = False
    End If
    
End Function

' Open EULA file using ShellExecute (handles any file type)
Private Sub OpenEULAFile(ByVal filePath As String)
    On Error Resume Next
    
    ' Try to use ShellExecute via WScript.Shell
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c start "" """ & filePath & """", 0, False
    Set wsh = Nothing
    
    If Err.Number <> 0 Then
        ' Last resort - try Shell
        Shell "cmd /c start "" """ & filePath & """", vbNormalFocus
    End If
    
    On Error GoTo 0
End Sub

' Public function to view EULA anytime (for Welcome sheet button)
Public Sub ViewEULA()
    Dim eulaPath As String
    eulaPath = ThisWorkbook.Path & "\EULA_Full.txt"
    
    If Dir(eulaPath) = "" Then
        MsgBox "EULA file not found at:" & vbCrLf & eulaPath, vbExclamation, "EULA Not Found"
        Exit Sub
    End If
    
    On Error Resume Next
    Shell "notepad.exe """ & eulaPath & """", vbNormalFocus
    If Err.Number <> 0 Then
        Call OpenEULAFile(eulaPath)
    End If
    On Error GoTo 0
End Sub
