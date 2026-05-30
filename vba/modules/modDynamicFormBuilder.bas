Attribute VB_Name = "modDynamicFormBuilder"
Option Explicit

Public DynamicControls As Object


Public Sub BuildDynamicForm(frm As frmItemEditor)

    ClearDynamicControls frm

    Set DynamicControls = CreateObject("Scripting.Dictionary")

    Dim i As Long

    Dim pg As MSForms.Page

    Dim currentTop As Single

    Dim tabPages As Object
    Dim tabTopPositions As Object

    Dim controlName As String

    Dim ctrl As Object

    Dim lookupType As String

    Set tabPages = CreateObject("Scripting.Dictionary")
    Set tabTopPositions = CreateObject("Scripting.Dictionary")

    For i = 0 To frm.mpItemTabs.Pages.Count - 1

        tabPages.Add _
            Trim(frm.mpItemTabs.Pages(i).caption), _
            frm.mpItemTabs.Pages(i)

    Next i

    LoadFieldDefinitions

    For Each pg In frm.mpItemTabs.Pages

        tabTopPositions.Add Trim(pg.caption), 35

    Next pg

    For i = LBound(FieldDefinitions) To UBound(FieldDefinitions)

        Set pg = tabPages(FieldDefinitions(i).tabName)

        currentTop = tabTopPositions(FieldDefinitions(i).tabName)

        If _
            FieldDefinitions(i).fieldName = "ITEM_WIDTH" _
            Or FieldDefinitions(i).fieldName = "ITEM_LENGTH" _
            Or FieldDefinitions(i).fieldName = "TOTAL_CONTENT_WIDTH" _
            Or FieldDefinitions(i).fieldName = "TOTAL_CONTENT_LENGTH" _
            Or FieldDefinitions(i).fieldName = "WEIGHT_OUNCES" Then

            GoTo SkipField

        End If

        CreateStandardLabel _
            pg, _
            "lblDynamic_" & FieldDefinitions(i).fieldName, _
            FieldDefinitions(i).fieldLabel, _
            currentTop

        If FieldDefinitions(i).ControlStyle = STANDARD_TEXT Then

            If FieldDefinitions(i).fieldName = "ITEM_HEIGHT" Then

                CreateDimensionsControl _
                    frm, _
                    pg, _
                    "ITEM", _
                    currentTop

                currentTop = currentTop + UI_ROW_SPACING

            ElseIf FieldDefinitions(i).fieldName = "TOTAL_CONTENT_HEIGHT" Then

                CreateDimensionsControl _
                    frm, _
                    pg, _
                    "TOTAL_CONTENT", _
                    currentTop

                currentTop = currentTop + UI_ROW_SPACING

            ElseIf _
                FieldDefinitions(i).fieldName = "WEIGHT_POUNDS" _
                Or FieldDefinitions(i).fieldName = "TOTAL_CONTENT_WEIGHT" Then

                CreateWeightControl _
                    frm, _
                    pg, _
                    "WEIGHT", _
                    currentTop

                currentTop = currentTop + UI_ROW_SPACING

            Else

                controlName = "txt_" & FieldDefinitions(i).fieldName

                Set ctrl = CreateStandardTextbox( _
                    frm, _
                    pg, _
                    controlName, _
                    currentTop)

                RegisterDynamicControl _
                    FieldDefinitions(i).fieldName, _
                    ctrl

                frm.RegisterTextboxEvent _
                    ctrl, _
                    FieldDefinitions(i).fieldName

                currentTop = currentTop + UI_ROW_SPACING

            End If

        End If

        If FieldDefinitions(i).ControlStyle = MULTILINE_TEXT Then

            controlName = "txt_" & FieldDefinitions(i).fieldName

            Set ctrl = CreateMultilineTextbox( _
                frm, _
                pg, _
                controlName, _
                currentTop)

            RegisterDynamicControl _
                FieldDefinitions(i).fieldName, _
                ctrl

            frm.RegisterTextboxEvent _
                ctrl, _
                FieldDefinitions(i).fieldName

            currentTop = currentTop + UI_MULTILINE_SPACING

        End If

        If FieldDefinitions(i).ControlStyle = EDITABLE_DROPDOWN Then

            controlName = "cbo_" & FieldDefinitions(i).fieldName

            CreateEditableComboBox _
                frm, _
                pg, _
                controlName, _
                currentTop

            lookupType = ""

            If Trim(FieldDefinitions(i).DropdownSource) <> "" Then

                lookupType = _
                    GetLookupTypeFromSource( _
                        FieldDefinitions(i).DropdownSource)

                PopulateComboBox _
                    pg.Controls(controlName), _
                    lookupType

            End If

            RegisterDynamicControl _
                FieldDefinitions(i).fieldName, _
                pg.Controls(controlName)

            frm.RegisterComboBoxEvent _
                pg.Controls(controlName), _
                FieldDefinitions(i).fieldName

            currentTop = currentTop + UI_ROW_SPACING

        End If

        If FieldDefinitions(i).ControlStyle = BOOLEAN_WITH_NOTES Then

            CreateBooleanWithNotesControls _
                frm, _
                pg, _
                FieldDefinitions(i).fieldName, _
                currentTop, _
                False

            currentTop = currentTop + UI_ROW_SPACING

        End If

        If FieldDefinitions(i).ControlStyle = BOOLEAN_WITH_NOTES_NA Then

            CreateBooleanWithNotesControls _
                frm, _
                pg, _
                FieldDefinitions(i).fieldName, _
                currentTop, _
                True

            currentTop = currentTop + UI_ROW_SPACING

        End If

SkipField:

        tabTopPositions(FieldDefinitions(i).tabName) = currentTop

    Next i

    CreatePhotosTabControls frm

End Sub


Public Sub RegisterDynamicControl( _
    ByVal fieldName As String, _
    ByVal ctrl As Object _
)

    If DynamicControls Is Nothing Then
        Set DynamicControls = CreateObject("Scripting.Dictionary")
    End If

    If DynamicControls.Exists(fieldName) Then
        DynamicControls.Remove fieldName
    End If

    Set DynamicControls(fieldName) = ctrl

End Sub


Public Function GetDynamicControl( _
    ByVal fieldName As String _
) As Object

    If DynamicControls Is Nothing Then Exit Function

    If DynamicControls.Exists(fieldName) Then

        Set GetDynamicControl = DynamicControls(fieldName)

    End If

End Function


Public Sub ClearDynamicControls(frm As frmItemEditor)

    Dim pg As MSForms.Page
    Dim i As Long

    Dim ctrlName As String

    For Each pg In frm.mpItemTabs.Pages

        For i = pg.Controls.Count - 1 To 0 Step -1

            ctrlName = pg.Controls(i).Name

            If Left(ctrlName, 4) = "txt_" _
            Or Left(ctrlName, 4) = "cbo_" _
            Or Left(ctrlName, 4) = "chk_" _
            Or Left(ctrlName, 4) = "opt_" _
            Or Left(ctrlName, 4) = "fra_" _
            Or Left(ctrlName, 11) = "lblDynamic" Then

                pg.Controls.Remove ctrlName

            End If

        Next i

    Next pg

End Sub

