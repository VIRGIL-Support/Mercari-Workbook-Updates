Attribute VB_Name = "modControlFactory"
Option Explicit

Public Function CreateStandardLabel( _
    pg As MSForms.Page, _
    controlName As String, _
    caption As String, _
    topPos As Single) As MSForms.Label

    Dim lbl As MSForms.Label

    Set lbl = pg.Controls.Add("Forms.Label.1")

    With lbl

        .Name = controlName
        .caption = caption

        .Left = UI_LEFT_MARGIN
        .Top = topPos + 3

        .width = UI_LABEL_WIDTH
        .height = 18

        .Font.Size = UI_FONT_SIZE
        .Font.Bold = True

    End With

    Set CreateStandardLabel = lbl

End Function


Public Function CreateStandardTextbox( _
    frm As frmItemEditor, _
    pg As MSForms.Page, _
    controlName As String, _
    topPos As Single) As MSForms.TextBox

    Dim txt As MSForms.TextBox

    Dim controlLeft As Single
    Dim controlWidth As Single

    Dim fieldName As String
    Dim placeholderText As String

    controlLeft = UI_LEFT_MARGIN + UI_LABEL_WIDTH + 10

    controlWidth = _
        frm.mpItemTabs.width _
        - controlLeft _
        - UI_RIGHT_MARGIN _
        - 10

    Set txt = pg.Controls.Add("Forms.TextBox.1")

    fieldName = Replace(controlName, "txt_", "")

    placeholderText = GetPlaceholderText(fieldName)

    With txt

        .Name = controlName

        .Left = controlLeft
        .Top = topPos

        .width = controlWidth
        .height = UI_STANDARD_HEIGHT

        .Font.Size = UI_FONT_SIZE

        If placeholderText <> "" Then

            .text = placeholderText
            .ForeColor = UI_PLACEHOLDER_FORECOLOR

        Else

            .ForeColor = UI_NORMAL_FORECOLOR

        End If

    End With

    Set CreateStandardTextbox = txt

End Function


Public Function CreateMultilineTextbox( _
    frm As frmItemEditor, _
    pg As MSForms.Page, _
    controlName As String, _
    topPos As Single) As MSForms.TextBox

    Dim txt As MSForms.TextBox

    Dim controlWidth As Single

    Dim fieldName As String
    Dim placeholderText As String

    controlWidth = _
        frm.mpItemTabs.width _
        - UI_LEFT_MARGIN _
        - UI_RIGHT_MARGIN _
        - 10

    Set txt = pg.Controls.Add("Forms.TextBox.1")

    fieldName = Replace(controlName, "txt_", "")

    placeholderText = GetPlaceholderText(fieldName)

    With txt

        .Name = controlName

        .Left = UI_LEFT_MARGIN
        .Top = topPos + UI_MULTILINE_TOP_OFFSET

        .width = controlWidth
        .height = UI_MULTILINE_HEIGHT

        .MultiLine = True

        .Font.Size = UI_FONT_SIZE

        If placeholderText <> "" Then

            .text = placeholderText
            .ForeColor = UI_PLACEHOLDER_FORECOLOR

        Else

            .ForeColor = UI_NORMAL_FORECOLOR

        End If

    End With

    Set CreateMultilineTextbox = txt

End Function


Public Function CreateEditableComboBox( _
    frm As frmItemEditor, _
    pg As MSForms.Page, _
    controlName As String, _
    topPos As Single) As MSForms.ComboBox

    Dim cbo As MSForms.ComboBox

    Dim controlLeft As Single
    Dim controlWidth As Single

    Dim fieldName As String
    Dim placeholderText As String

    controlLeft = UI_LEFT_MARGIN + UI_LABEL_WIDTH + 10

    controlWidth = _
        frm.mpItemTabs.width _
        - controlLeft _
        - UI_RIGHT_MARGIN _
        - 10

    Set cbo = pg.Controls.Add("Forms.ComboBox.1")

    fieldName = Replace(controlName, "cbo_", "")

    placeholderText = GetPlaceholderText(fieldName)

    With cbo

        .Name = controlName

        .Left = controlLeft
        .Top = topPos

        .width = controlWidth
        .height = UI_STANDARD_HEIGHT

        ConfigureLookupComboBox cbo

        .Font.Size = UI_FONT_SIZE

        If placeholderText <> "" Then

            .text = placeholderText
            .ForeColor = UI_PLACEHOLDER_FORECOLOR

        Else

            .ForeColor = UI_NORMAL_FORECOLOR

        End If

    End With

    Set CreateEditableComboBox = cbo

End Function


Public Sub CreateBooleanWithNotesControls( _
    frm As frmItemEditor, _
    pg As MSForms.Page, _
    fieldName As String, _
    topPos As Single, _
    Optional ByVal includeNA As Boolean = False)

    Dim fra As MSForms.Frame

    Dim optYes As MSForms.OptionButton
    Dim optNo As MSForms.OptionButton
    Dim optNA As MSForms.OptionButton

    Dim txt As MSForms.TextBox

    Dim controlLeft As Single
    Dim notesLeft As Single
    Dim notesWidth As Single

    Dim placeholderText As String

    controlLeft = UI_LEFT_MARGIN + UI_LABEL_WIDTH + 10

    notesLeft = controlLeft + 190

    notesWidth = _
        frm.mpItemTabs.width _
        - notesLeft _
        - UI_RIGHT_MARGIN _
        - 10

    Set fra = pg.Controls.Add("Forms.Frame.1")

    With fra

        .Name = "fra_" & fieldName

        .Left = controlLeft
        .Top = topPos - 2

        .width = 180
        .height = 24

        .caption = ""

        .SpecialEffect = fmSpecialEffectFlat

    End With

    Set optYes = fra.Controls.Add("Forms.OptionButton.1")

    With optYes

        .Name = "opt_" & fieldName & "_YES"

        .caption = "Yes"

        .Left = 5
        .Top = 3

        .width = 50

        .Font.Size = UI_FONT_SIZE

    End With

    Set optNo = fra.Controls.Add("Forms.OptionButton.1")

    With optNo

        .Name = "opt_" & fieldName & "_NO"

        .caption = "No"

        .Left = 60
        .Top = 3

        .width = 50

        .Font.Size = UI_FONT_SIZE

    End With

    If includeNA = True Then

        Set optNA = fra.Controls.Add("Forms.OptionButton.1")

        With optNA

            .Name = "opt_" & fieldName & "_NA"

            .caption = "N/A"

            .Left = 115
            .Top = 3

            .width = 60

            .Font.Size = UI_FONT_SIZE

        End With

    End If

    Set txt = pg.Controls.Add("Forms.TextBox.1")

    placeholderText = GetPlaceholderText(fieldName & "_NOTES")

    With txt

        .Name = "txt_" & fieldName & "_NOTES"

        .Left = notesLeft
        .Top = topPos

        .width = notesWidth
        .height = UI_STANDARD_HEIGHT

        .Font.Size = UI_FONT_SIZE

        If placeholderText <> "" Then

            .text = placeholderText
            .ForeColor = UI_PLACEHOLDER_FORECOLOR

        Else

            .ForeColor = UI_NORMAL_FORECOLOR

        End If

    End With

    frm.RegisterTextboxEvent _
        txt, _
        fieldName & "_NOTES"

End Sub


Public Sub CreateDimensionsControl( _
    frm As frmItemEditor, _
    pg As MSForms.Page, _
    fieldPrefix As String, _
    topPos As Single)

    Dim txt1 As MSForms.TextBox
    Dim txt2 As MSForms.TextBox
    Dim txt3 As MSForms.TextBox

    Dim lblX1 As MSForms.Label
    Dim lblX2 As MSForms.Label

    Dim startLeft As Single

    startLeft = UI_LEFT_MARGIN + UI_LABEL_WIDTH + 10

    Set txt1 = pg.Controls.Add("Forms.TextBox.1")
    Set txt2 = pg.Controls.Add("Forms.TextBox.1")
    Set txt3 = pg.Controls.Add("Forms.TextBox.1")

    With txt1

        .Name = "txt_" & fieldPrefix & "_HEIGHT"

        .Left = startLeft
        .Top = topPos + 2

        .width = 55
        .height = UI_STANDARD_HEIGHT

        .Font.Size = UI_FONT_SIZE

        .text = "W"
        .ForeColor = UI_PLACEHOLDER_FORECOLOR

    End With

    frm.RegisterTextboxEvent txt1, fieldPrefix & "_HEIGHT"

    Set lblX1 = pg.Controls.Add("Forms.Label.1")

    With lblX1

        .caption = "x"

        .Left = startLeft + 62
        .Top = topPos + 4

        .width = 10

        .Font.Size = UI_FONT_SIZE
        .Font.Bold = True

    End With

    With txt2

        .Name = "txt_" & fieldPrefix & "_WIDTH"

        .Left = startLeft + 78
        .Top = topPos + 2

        .width = 55
        .height = UI_STANDARD_HEIGHT

        .Font.Size = UI_FONT_SIZE

        .text = "H"
        .ForeColor = UI_PLACEHOLDER_FORECOLOR

    End With

    frm.RegisterTextboxEvent txt2, fieldPrefix & "_WIDTH"

    Set lblX2 = pg.Controls.Add("Forms.Label.1")

    With lblX2

        .caption = "x"

        .Left = startLeft + 140
        .Top = topPos + 4

        .width = 10

        .Font.Size = UI_FONT_SIZE
        .Font.Bold = True

    End With

    With txt3

        .Name = "txt_" & fieldPrefix & "_LENGTH"

        .Left = startLeft + 156
        .Top = topPos + 2

        .width = 55
        .height = UI_STANDARD_HEIGHT

        .Font.Size = UI_FONT_SIZE

        .text = "L"
        .ForeColor = UI_PLACEHOLDER_FORECOLOR

    End With

    frm.RegisterTextboxEvent txt3, fieldPrefix & "_LENGTH"

    RegisterDynamicControl fieldPrefix & "_HEIGHT", txt1
    RegisterDynamicControl fieldPrefix & "_WIDTH", txt2
    RegisterDynamicControl fieldPrefix & "_LENGTH", txt3

End Sub


Public Sub CreateWeightControl( _
    frm As frmItemEditor, _
    pg As MSForms.Page, _
    fieldPrefix As String, _
    topPos As Single)

    Dim txtLbs As MSForms.TextBox
    Dim txtOz As MSForms.TextBox

    Dim lblLbs As MSForms.Label
    Dim lblOz As MSForms.Label

    Dim startLeft As Single

    startLeft = UI_LEFT_MARGIN + UI_LABEL_WIDTH + 10

    Set txtLbs = pg.Controls.Add("Forms.TextBox.1")
    Set txtOz = pg.Controls.Add("Forms.TextBox.1")

    With txtLbs

        .Name = "txt_" & fieldPrefix & "_POUNDS"

        .Left = startLeft
        .Top = topPos + 2

        .width = 55
        .height = UI_STANDARD_HEIGHT

        .Font.Size = UI_FONT_SIZE

        .text = "lbs"
        .ForeColor = UI_PLACEHOLDER_FORECOLOR

    End With

    frm.RegisterTextboxEvent txtLbs, fieldPrefix & "_POUNDS"

    Set lblLbs = pg.Controls.Add("Forms.Label.1")

    With lblLbs

        .caption = "lbs"

        .Left = startLeft + 62
        .Top = topPos + 4

        .width = 22

        .Font.Size = UI_FONT_SIZE

    End With

    With txtOz

        .Name = "txt_" & fieldPrefix & "_OUNCES"

        .Left = startLeft + 95
        .Top = topPos + 2

        .width = 55
        .height = UI_STANDARD_HEIGHT

        .Font.Size = UI_FONT_SIZE

        .text = "oz"
        .ForeColor = UI_PLACEHOLDER_FORECOLOR

    End With

    frm.RegisterTextboxEvent txtOz, fieldPrefix & "_OUNCES"

    Set lblOz = pg.Controls.Add("Forms.Label.1")

    With lblOz

        .caption = "oz"

        .Left = startLeft + 157
        .Top = topPos + 4

        .width = 20

        .Font.Size = UI_FONT_SIZE

    End With

    RegisterDynamicControl fieldPrefix & "_POUNDS", txtLbs
    RegisterDynamicControl fieldPrefix & "_OUNCES", txtOz

End Sub

