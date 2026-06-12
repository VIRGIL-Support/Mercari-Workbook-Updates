Attribute VB_Name = "modSellerEssentials"
Option Explicit

'=================================================================
' Seller Essentials — Navigation & Button Module
' Provides equal-width navigation buttons with muted/active colors
'=================================================================

Private m_ActiveSection As Integer   '  0 = show all

'-- Row constants (based on actual worksheet layout) ----------------
Private Const NAV_ROW     As Long = 3
Private Const DYN_HDR_ROW As Long = 4

' Section 1: Photography (Teal)
Private Const S1_BLOCK_START As Long = 5
Private Const S1_BLOCK_END   As Long = 38
Private Const S1_ACCENT      As String = "00A896"   ' Full/Active
Private Const S1_MUTED       As String = "9ED5CF"   ' Muted/Inactive
Private Const S1_LABEL       As String = "Photography"

' Section 2: Boxes (Orange)
Private Const S2_BLOCK_START As Long = 39
Private Const S2_BLOCK_END   As Long = 59
Private Const S2_ACCENT      As String = "E8622A"
Private Const S2_MUTED       As String = "F2B9A0"
Private Const S2_LABEL       As String = "Boxes"

' Section 3: Shipping (Amber)
Private Const S3_BLOCK_START As Long = 60
Private Const S3_BLOCK_END   As Long = 97
Private Const S3_ACCENT      As String = "F5A623"
Private Const S3_MUTED       As String = "FAD99A"
Private Const S3_LABEL       As String = "Shipping"

' Section 4: Storage (Emerald)
Private Const S4_BLOCK_START As Long = 98
Private Const S4_BLOCK_END   As Long = 142
Private Const S4_ACCENT      As String = "27AE60"
Private Const S4_MUTED       As String = "99D4B3"
Private Const S4_LABEL       As String = "Storage"

' Section 5: Cleaning (Rose)
Private Const S5_BLOCK_START As Long = 143
Private Const S5_BLOCK_END   As Long = 191
Private Const S5_ACCENT      As String = "C94B8A"
Private Const S5_MUTED       As String = "E8A8CC"
Private Const S5_LABEL       As String = "Cleaning"

' Section 6: Measuring (Steel Blue)
Private Const S6_BLOCK_START As Long = 192
Private Const S6_BLOCK_END   As Long = 243
Private Const S6_ACCENT      As String = "3D7BC5"
Private Const S6_MUTED       As String = "A8C4E2"
Private Const S6_LABEL       As String = "Measuring"

' Section 7: Office (Violet)
Private Const S7_BLOCK_START As Long = 244
Private Const S7_BLOCK_END   As Long = 297
Private Const S7_ACCENT      As String = "7F5AC4"
Private Const S7_MUTED       As String = "C4B3E8"
Private Const S7_LABEL       As String = "Office"

'-- Button configuration -------------------------------------------
Private Const BTN_WIDTH    As Double = 138   ' Fixed width for all buttons (25% wider)
Private Const BTN_HEIGHT   As Double = 24    ' Height of buttons
Private Const BTN_SPACING  As Double = 24    ' Padding between buttons (doubled)
Private Const BTN_START_COL As String = "B" ' Starting column (B = column 2)

'-- Nav click handler (called from sheet SelectionChange) ----------
Public Sub Essentials_NavClick(colNum As Long)
    Dim sNum As Integer
    Select Case colNum
        Case 2: sNum = 1
        Case 3: sNum = 2
        Case 4: sNum = 3
        Case 5: sNum = 4
        Case 6: sNum = 5
        Case 7: sNum = 6
        Case 8: sNum = 7
        Case Else: Exit Sub
    End Select
    If sNum = m_ActiveSection Then
        Call Essentials_ShowAll
    Else
        Call Essentials_ShowSection(sNum)
    End If
End Sub

'-- Show one section (hides all others) ----------------------------
Public Sub Essentials_ShowSection(sNum As Integer)
    On Error GoTo ShowSectionError
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("SELLER ESSENTIALS")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Unhide everything first
    ws.Rows.Hidden = False
    
    ' Hide all section blocks except the selected one
    Dim i As Integer
    For i = 1 To 7
        If i <> sNum Then
            ws.Rows(Essentials_BlockStart(i) & ":" & Essentials_BlockEnd(i)).Hidden = True
        End If
    Next i
    
    m_ActiveSection = sNum
    
    ' Update button colors
    Call Essentials_UpdateNavButtons(ws, sNum)
    
    ' Update dynamic header color
    Call Essentials_UpdateDynHeader(ws, sNum)
    
    ' Scroll to top of selected section
    Application.Goto ws.Cells(Essentials_BlockStart(sNum), 2), True
    
ShowSectionCleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ShowSectionError:
    MsgBox "Error showing section: " & Err.Description & " (Error " & Err.Number & ")", vbExclamation, "Error"
    Resume ShowSectionCleanup
End Sub

'-- Show all sections ----------------------------------------------
Public Sub Essentials_ShowAll()
    On Error GoTo ShowAllError
    
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("SELLER ESSENTIALS")
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ws.Rows.Hidden = False
    m_ActiveSection = 0
    
    ' Reset all buttons to muted
    Call Essentials_UpdateNavButtons(ws, 0)
    
    ' Reset header to default
    Call Essentials_UpdateDynHeader(ws, 0)
    
    ws.Cells(1, 1).Select
    
ShowAllCleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ShowAllError:
    MsgBox "Error showing all: " & Err.Description, vbExclamation, "Error"
    Resume ShowAllCleanup
End Sub

'-- Update nav button fills (active = full color, others = muted) --
Private Sub Essentials_UpdateNavButtons(ws As Worksheet, activeS As Integer)
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim tag As String
        tag = shp.AlternativeText
        If IsNumeric(tag) Then
            Dim sn As Integer: sn = CInt(tag)
            If sn >= 1 And sn <= 7 Then
                If sn = activeS Then
                    ' Active: full accent color with white text
                    shp.Fill.ForeColor.RGB = HexToRGB(Essentials_Accent(sn))
                    shp.Line.ForeColor.RGB = RGB(255, 255, 255)
                    shp.Line.Weight = 2
                    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                Else
                    ' Inactive: muted color with dark gray text
                    shp.Fill.ForeColor.RGB = HexToRGB(Essentials_Muted(sn))
                    shp.Line.ForeColor.RGB = RGB(180, 180, 180)
                    shp.Line.Weight = 1
                    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
                End If
            End If
        End If
    Next shp
End Sub

'-- Update dynamic header row based on which section is active -----
Public Sub Essentials_UpdateDynHeader(ws As Worksheet, sNum As Integer)
    Dim bg As Long, fg As Long
    
    If sNum = 0 Then
        ' Default: dark background when showing all
        bg = HexToRGB("1A1A1A")
        fg = RGB(255, 255, 255)
    Else
        ' Active section color
        bg = HexToRGB(Essentials_Accent(sNum))
        fg = RGB(255, 255, 255)
    End If
    
    ' Apply color to Row 4, columns A through W only (not beyond)
    ws.Range("A" & DYN_HDR_ROW & ":W" & DYN_HDR_ROW).Interior.Color = bg
    ' Clear color from columns X onward
    ws.Range("X" & DYN_HDR_ROW & ":AZ" & DYN_HDR_ROW).Interior.ColorIndex = xlNone
    
    ' Set font color for header text
    ws.Cells(DYN_HDR_ROW, 2).Font.Color = fg
End Sub

'-- Create equal-width rounded-rectangle buttons ---------------------
Public Sub Essentials_CreateNavButtons()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "SELLER ESSENTIALS"
    
    ' Check if sheet exists
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation, "Error"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Clear original cell contents and colors in the nav button area (B3:H3)
    With ws.Range("B3:H3")
        .ClearContents
        .Interior.Color = RGB(238, 240, 253)  ' Pale tint #EEF0FD
        .Borders.LineStyle = xlNone  ' Remove cell borders too
    End With
    
    ' Remove existing nav shapes (those with numeric AlternativeText) - limit iterations
    Dim shp As Shape
    Dim deletedCount As Integer
    deletedCount = 0
    For Each shp In ws.Shapes
        If IsNumeric(shp.AlternativeText) Then
            shp.Delete
            deletedCount = deletedCount + 1
            If deletedCount > 20 Then Exit For ' Safety limit
        End If
    Next shp
    
    ' Calculate position - vertically centered in Row 3
    Dim navTop As Double
    Dim rowHeight As Double
    rowHeight = ws.Rows(NAV_ROW).Height
    navTop = ws.Rows(NAV_ROW).Top + (rowHeight - BTN_HEIGHT) / 2
    
    ' Calculate centered starting position for button bar
    Dim startLeft As Double
    Dim totalBtnBarWidth As Double
    Dim rowWidth As Double
    totalBtnBarWidth = (7 * BTN_WIDTH) + (6 * BTN_SPACING)  ' 7 buttons + 6 spaces
    rowWidth = ws.Columns("B").Width + ws.Columns("C").Width + ws.Columns("D").Width + _
               ws.Columns("E").Width + ws.Columns("F").Width + ws.Columns("G").Width + _
               ws.Columns("H").Width
    startLeft = ws.Columns("B").Left + (rowWidth - totalBtnBarWidth) / 2
    
    ' Create 7 equal-width buttons
    Dim labels(1 To 7) As String
    labels(1) = S1_LABEL
    labels(2) = S2_LABEL
    labels(3) = S3_LABEL
    labels(4) = S4_LABEL
    labels(5) = S5_LABEL
    labels(6) = S6_LABEL
    labels(7) = S7_LABEL
    
    Dim i As Integer
    For i = 1 To 7
        Dim btnLeft As Double
        btnLeft = startLeft + (i - 1) * (BTN_WIDTH + BTN_SPACING)
        
        Dim s As Shape
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            btnLeft, navTop, BTN_WIDTH, BTN_HEIGHT)
        
        With s
            ' Store section number in AlternativeText for identification
            .AlternativeText = CStr(i)
            
            ' Fill with muted color (inactive by default)
            .Fill.ForeColor.RGB = HexToRGB(Essentials_Muted(i))
            .Fill.Visible = msoTrue
            .Fill.Solid
            
            ' Subtle border
            .Line.ForeColor.RGB = RGB(180, 180, 180)
            .Line.Weight = 1
            
            ' Shadow effect
            .Shadow.Visible = msoTrue
            .Shadow.Style = msoShadowStyleOuterShadow
            .Shadow.OffsetX = 1
            .Shadow.OffsetY = 2
            .Shadow.Blur = 4
            .Shadow.ForeColor.RGB = RGB(0, 0, 0)
            .Shadow.Transparency = 0.7
            
            ' Text formatting
            .TextFrame2.TextRange.text = labels(i)
            .TextFrame2.TextRange.Font.Size = 18
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            
            ' Assign click action
            .OnAction = "modSellerEssentials.Essentials_Btn" & i
        End With
    Next i
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Error creating buttons: " & Err.Description, vbExclamation, "Error"
End Sub

'-- Individual button macros (assigned via .OnAction) ---------------
Sub Essentials_Btn1(): Call Essentials_ShowSection(1): End Sub
Sub Essentials_Btn2(): Call Essentials_ShowSection(2): End Sub
Sub Essentials_Btn3(): Call Essentials_ShowSection(3): End Sub
Sub Essentials_Btn4(): Call Essentials_ShowSection(4): End Sub
Sub Essentials_Btn5(): Call Essentials_ShowSection(5): End Sub
Sub Essentials_Btn6(): Call Essentials_ShowSection(6): End Sub
Sub Essentials_Btn7(): Call Essentials_ShowSection(7): End Sub

'-- Helper: section block row ranges -------------------------------
Private Function Essentials_BlockStart(sNum As Integer) As Long
    Select Case sNum
        Case 1: Essentials_BlockStart = S1_BLOCK_START
        Case 2: Essentials_BlockStart = S2_BLOCK_START
        Case 3: Essentials_BlockStart = S3_BLOCK_START
        Case 4: Essentials_BlockStart = S4_BLOCK_START
        Case 5: Essentials_BlockStart = S5_BLOCK_START
        Case 6: Essentials_BlockStart = S6_BLOCK_START
        Case 7: Essentials_BlockStart = S7_BLOCK_START
    End Select
End Function

Private Function Essentials_BlockEnd(sNum As Integer) As Long
    Select Case sNum
        Case 1: Essentials_BlockEnd = S1_BLOCK_END
        Case 2: Essentials_BlockEnd = S2_BLOCK_END
        Case 3: Essentials_BlockEnd = S3_BLOCK_END
        Case 4: Essentials_BlockEnd = S4_BLOCK_END
        Case 5: Essentials_BlockEnd = S5_BLOCK_END
        Case 6: Essentials_BlockEnd = S6_BLOCK_END
        Case 7: Essentials_BlockEnd = S7_BLOCK_END
    End Select
End Function

Private Function Essentials_Accent(sNum As Integer) As String
    Select Case sNum
        Case 1: Essentials_Accent = S1_ACCENT
        Case 2: Essentials_Accent = S2_ACCENT
        Case 3: Essentials_Accent = S3_ACCENT
        Case 4: Essentials_Accent = S4_ACCENT
        Case 5: Essentials_Accent = S5_ACCENT
        Case 6: Essentials_Accent = S6_ACCENT
        Case 7: Essentials_Accent = S7_ACCENT
    End Select
End Function

Private Function Essentials_Muted(sNum As Integer) As String
    Select Case sNum
        Case 1: Essentials_Muted = S1_MUTED
        Case 2: Essentials_Muted = S2_MUTED
        Case 3: Essentials_Muted = S3_MUTED
        Case 4: Essentials_Muted = S4_MUTED
        Case 5: Essentials_Muted = S5_MUTED
        Case 6: Essentials_Muted = S6_MUTED
        Case 7: Essentials_Muted = S7_MUTED
    End Select
End Function

'-- Manual test: Show Photography and verify hiding ---------------
Public Sub TestShowPhotography()
    On Error Resume Next
    MsgBox "Testing: Will show only Photography section (rows 7-41)", vbInformation, "Test"
    Call Essentials_ShowSection(1)
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description & " (" & Err.Number & ")", vbExclamation, "Test Failed"
    Else
        MsgBox "Photography section should now be visible. Check if other sections are hidden.", vbInformation, "Test Complete"
    End If
    On Error GoTo 0
End Sub

'-- Hex color string to RGB Long ---------------------------------
Private Function HexToRGB(hexColor As String) As Long
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Left(hexColor, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Right(hexColor, 2))
    HexToRGB = RGB(r, g, b)
End Function
