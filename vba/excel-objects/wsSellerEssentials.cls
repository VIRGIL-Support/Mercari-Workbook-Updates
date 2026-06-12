VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsSellerEssentials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    ' Handle nav button clicks
    If Target.Row = 3 And Target.Column >= 2 And Target.Column <= 8 Then
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Call modSellerEssentials.Essentials_NavClick(Target.Column)
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    Static buttonsCreated As Boolean
    
    ' First time only: Create buttons and default to Photography
    If Not buttonsCreated Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        ' Create the navigation buttons
        Call modSellerEssentials.Essentials_CreateNavButtons
        
        ' Default to Photography section (section 1) on first visit only
        Call modSellerEssentials.Essentials_ShowSection(1)
        
        buttonsCreated = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    ' On subsequent activations, do nothing - preserve user's current view
    On Error GoTo 0
End Sub
