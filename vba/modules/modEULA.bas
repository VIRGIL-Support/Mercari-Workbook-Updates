Attribute VB_Name = "modEULA"
' ============================================
' END USER LICENSE AGREEMENT (EULA) MODULE
' ============================================

Public Function DisplayEULA() As Boolean
    ' Returns True if user accepts, False if declined
    
    Dim eulaPath As String
    Dim userResponse As VbMsgBoxResult
    
    ' Build path to full EULA text file
    eulaPath = ThisWorkbook.Path & "\EULA_Full.txt"
    
    ' Open EULA file automatically (like clicking a link)
    On Error Resume Next
    If Dir(eulaPath) <> "" Then
        Shell "notepad.exe """ & eulaPath & """", vbNormalFocus
    End If
    On Error GoTo 0
    
    ' Show popup asking for acceptance (EULA is already open in Notepad)
    userResponse = MsgBox("The End User License Agreement (EULA) has been opened in Notepad." & vbCrLf & vbCrLf & _
                          "After reading the EULA, please confirm:" & vbCrLf & vbCrLf & _
                          "Do you accept all terms of the End User License Agreement?" & vbCrLf & vbCrLf & _
                          "Key terms:" & vbCrLf & _
                          "- Single user license (no sharing)" & vbCrLf & _
                          "- No redistribution or reverse engineering" & vbCrLf & _
                          "- Software provided AS IS with no warranty" & vbCrLf & _
                          "- Support via VIRGIL_Support@proton.me" & vbCrLf & vbCrLf & _
                          "Click YES to accept and continue." & vbCrLf & _
                          "Click NO to decline and close.", _
                          vbYesNo + vbQuestion + vbDefaultButton1, "Accept EULA")
    
    If userResponse = vbYes Then
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
