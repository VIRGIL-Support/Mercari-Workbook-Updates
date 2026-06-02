VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemEditor 
   Caption         =   "Item Editor"
   ClientHeight    =   13440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17784
   OleObjectBlob   =   "frmItemEditor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mCurrentDataRow As Long

Private mItemNumber As String
Private mItemName As String

Private mIsDirty As Boolean

Private mTextBoxEvents As Collection
Private mComboBoxEvents As Collection
Private mPhotoButtonEvents As Collection
Private mSelectedPhotoPaths As Collection
Private mRemovedPhotoPaths As Collection
Private mPhotoTabLoaded As Boolean
Private mPhotosChanged As Boolean

Private mCurrentTabIndex As Long
Private mSuppressTabChange As Boolean


' =====================================================
' SAVE + CLOSE
' =====================================================

Private Sub btnSaveClose_Click()

    If SaveFormData = False Then Exit Sub

    ThisWorkbook.Save
    Unload Me

End Sub

' =====================================================
' MANAGE DROPDOWN LISTS
' =====================================================

Private Sub btnManageLists_Click()

    Dim wasDirty As Boolean
    
    wasDirty = mIsDirty
    
    ' Save current data before opening manager
    SaveAllFieldValues Me
    
    ' Hide form temporarily while managing lists
    Me.Hide
    
    ' Open the lookup manager
    ManageLookupValues
    
    ' Reload form data to refresh dropdowns
    LoadFormData
    
    ' Show form again
    Me.Show
    
    If wasDirty Then MarkDirty
    
    lblStatus.caption = "Dropdown lists updated " & Format(Now, "h:mm:ss AM/PM")

End Sub


' =====================================================
' INITIALIZE
' =====================================================

Private Sub UserForm_Initialize()

    Dim i As Long

    Set mTextBoxEvents = New Collection
    Set mComboBoxEvents = New Collection
    Set mPhotoButtonEvents = New Collection
    Set mSelectedPhotoPaths = New Collection
    Set mRemovedPhotoPaths = New Collection
    mPhotoTabLoaded = False
    mPhotosChanged = False

    For i = 0 To mpItemTabs.Pages.Count - 1

        mpItemTabs.Pages(i).caption = _
            "   " & Trim(mpItemTabs.Pages(i).caption) & "   "

    Next i

    mpItemTabs.Font.Size = UI_TAB_FONT_SIZE
    mpItemTabs.Value = GetTabIndexByCaption("IDENTITY")
    mCurrentTabIndex = mpItemTabs.Value

    BuildDynamicForm Me

    RefreshHeaderDisplay

End Sub


' =====================================================
' ACTIVATE
' =====================================================

Private Sub UserForm_Activate()

    Static alreadyLoaded As Boolean

    If alreadyLoaded = True Then Exit Sub

    LoadFormData

    alreadyLoaded = True

End Sub


' =====================================================
' TAB CHANGE AUTOSAVE
' =====================================================

Private Sub mpItemTabs_Change()

    Dim requestedTabIndex As Long

    If mSuppressTabChange Then Exit Sub

    requestedTabIndex = mpItemTabs.Value

    If requestedTabIndex = mCurrentTabIndex Then Exit Sub

    lblStatus.caption = "Autosaving..."

    SaveAllFieldValues Me

    lblStatus.caption = _
        "Autosaved " & Format(Now, "h:mm:ss AM/PM")

    mIsDirty = False
    mCurrentTabIndex = requestedTabIndex
    If UCase$(Trim$(mpItemTabs.Pages(requestedTabIndex).caption)) = "PHOTOS" Then
        RefreshPhotoTab
        mPhotoTabLoaded = True
    End If

End Sub


' =====================================================
' LOAD FORM DATA
' =====================================================

Public Sub LoadFormData()

    lblStatus.caption = "Loading..."

    LoadAllFieldValues Me
    Set mSelectedPhotoPaths = LoadPhotoSourcePaths(mCurrentDataRow)
    Set mRemovedPhotoPaths = New Collection
    mPhotosChanged = False
    If UCase$(Trim$(mpItemTabs.Pages(mpItemTabs.Value).caption)) = "PHOTOS" Then
        RefreshPhotoTab
        mPhotoTabLoaded = True
    End If

    lblStatus.caption = _
        "Loaded " & Format(Now, "h:mm:ss AM/PM")

    mIsDirty = False

End Sub


' =====================================================
' SAVE FORM DATA
' =====================================================

Public Function SaveFormData() As Boolean

    Dim copiedPhotoPaths As Collection
    Dim itemFolderPath As String

    If ValidateItemEditorForm(Me) = False Then Exit Function
    If mSelectedPhotoPaths.Count > DEFAULT_MAX_PHOTOS Then
        MsgBox "Mercari allows a maximum of 12 photos per listing. Please remove photos until the listing has 12 photos or fewer.", vbExclamation
        Exit Function
    End If

    lblStatus.caption = "Saving..."

    SaveAllFieldValues Me
    itemFolderPath = GetItemPhotoFolderPath(mCurrentDataRow)
    CreateFolderIfMissing itemFolderPath
    If mSelectedPhotoPaths.Count > 0 Then
        Set copiedPhotoPaths = CopySelectedPhotosToItemFolder(mCurrentDataRow, mSelectedPhotoPaths)
        Set mSelectedPhotoPaths = copiedPhotoPaths
    End If
    If mRemovedPhotoPaths.Count > 0 Then
        MoveRemovedPhotosToRemovedFolder mCurrentDataRow, mRemovedPhotoPaths
        Set mRemovedPhotoPaths = New Collection
    End If
    SavePhotoSourcePaths mCurrentDataRow, mSelectedPhotoPaths
    ThisWorkbook.Worksheets(WS_INVENTORY).Cells(mCurrentDataRow, COL_DETAILS_FOLDER).ClearContents

    ThisWorkbook.Save
    lblStatus.caption = "Saved " & Format(Now, "h:mm:ss AM/PM")

    mIsDirty = False
    mPhotosChanged = False

    SaveFormData = True

End Function


' =====================================================
' REGISTER TEXTBOX EVENT
' =====================================================

Public Sub RegisterTextboxEvent( _
    ByVal txt As MSForms.TextBox, _
    ByVal fieldName As String)

    Dim evt As clsTextBoxEvents

    Set evt = New clsTextBoxEvents

    evt.Initialize txt, fieldName, Me

    mTextBoxEvents.Add evt

End Sub


Public Sub RestoreEmptyTextBoxPlaceholders()

    Dim evt As clsTextBoxEvents

    For Each evt In mTextBoxEvents

        evt.RestoreIfEmpty

    Next evt

End Sub


Public Sub RestoreEmptyComboBoxPlaceholders()

    Dim evt As clsComboBoxEvents

    For Each evt In mComboBoxEvents

        evt.RestoreIfEmpty

    Next evt

End Sub


Private Sub UserForm_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    RestoreEmptyTextBoxPlaceholders
    RestoreEmptyComboBoxPlaceholders

End Sub


' =====================================================
' REGISTER COMBOBOX EVENT
' =====================================================

Public Sub RegisterComboBoxEvent( _
    ByVal cbo As MSForms.ComboBox, _
    ByVal fieldName As String)

    Dim evt As clsComboBoxEvents

    Set evt = New clsComboBoxEvents

    evt.Initialize cbo, fieldName, Me

    mComboBoxEvents.Add evt

End Sub


' =====================================================
' PHOTO TAB EVENTS
' =====================================================

Public Sub RegisterPhotoButtonEvent( _
    ByVal btn As MSForms.CommandButton, _
    ByVal actionName As String, _
    Optional ByVal photoIndex As Long = 0)

    Dim evt As clsPhotoButtonEvents

    Set evt = New clsPhotoButtonEvents
    evt.Initialize btn, Me, actionName, photoIndex
    mPhotoButtonEvents.Add evt

End Sub


Public Sub SelectPhotos()

    Dim fd As FileDialog
    Dim selectedItem As Variant
    Dim currentPhotoCount As Long
    Dim newPhotoCount As Long
    Dim totalPhotoCount As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .AllowMultiSelect = True
        .Title = "Select up to 12 item photos"
        .Filters.Clear
        .Filters.Add "Image Files", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.webp;*.heic;*.heif;*.tif;*.tiff"
        If .Show <> -1 Then Exit Sub
        newPhotoCount = .SelectedItems.Count
        currentPhotoCount = mSelectedPhotoPaths.Count
        totalPhotoCount = currentPhotoCount + newPhotoCount
        If totalPhotoCount > DEFAULT_MAX_PHOTOS Then
            ShowPhotoLimitMessage currentPhotoCount, newPhotoCount
            Exit Sub
        End If
        For Each selectedItem In .SelectedItems
            If IsSupportedPhotoFile(CStr(selectedItem)) And IsPhotoAlreadySelected(CStr(selectedItem)) = False Then
                mSelectedPhotoPaths.Add CStr(selectedItem)
            End If
        Next selectedItem
    End With

    RefreshPhotoTab
    mPhotosChanged = True
    MarkDirty

End Sub


Private Function IsPhotoAlreadySelected(ByVal photoPath As String) As Boolean

    Dim i As Long

    For i = 1 To mSelectedPhotoPaths.Count
        If LCase$(CStr(mSelectedPhotoPaths(i))) = LCase$(photoPath) Then
            IsPhotoAlreadySelected = True
            Exit Function
        End If
    Next i

End Function


Private Sub ShowPhotoLimitMessage(ByVal previousCount As Long, ByVal importCount As Long)

    Dim totalCount As Long

    totalCount = previousCount + importCount

    MsgBox _
        "Mercari allows a maximum of 12 photos per listing." & vbCrLf & vbCrLf & _
        "Previously Imported: " & previousCount & vbCrLf & _
        "This Import: " & importCount & vbCrLf & _
        "Total: " & totalCount & vbCrLf & _
        "Variance: " & (totalCount - DEFAULT_MAX_PHOTOS) & vbCrLf & vbCrLf & _
        "Please remove existing photos or choose fewer new photos so the listing stays within Mercari's 12-photo limit.", _
        vbExclamation

End Sub


Public Sub RemoveAllPhotos()

    Dim i As Long
    For i = 1 To mSelectedPhotoPaths.Count
        mRemovedPhotoPaths.Add CStr(mSelectedPhotoPaths(i))
    Next i

    Set mSelectedPhotoPaths = New Collection
    SavePhotoSourcePaths mCurrentDataRow, mSelectedPhotoPaths
    RefreshPhotoTab
    mPhotosChanged = True
    MarkDirty

End Sub


Public Sub RemovePhotoAt(ByVal photoIndex As Long)

    Dim newPaths As New Collection
    Dim i As Long

    If photoIndex < 1 Or photoIndex > mSelectedPhotoPaths.Count Then Exit Sub
    mRemovedPhotoPaths.Add CStr(mSelectedPhotoPaths(photoIndex))
    For i = 1 To mSelectedPhotoPaths.Count
        If i <> photoIndex Then newPaths.Add mSelectedPhotoPaths(i)
    Next i

    Set mSelectedPhotoPaths = newPaths
    SavePhotoSourcePaths mCurrentDataRow, mSelectedPhotoPaths
    RefreshPhotoTab
    mPhotosChanged = True
    MarkDirty

End Sub


Public Sub RefreshPhotoTab()

    Dim pg As MSForms.Page
    Dim i As Long
    Dim img As MSForms.Image
    Dim lbl As MSForms.Label
    Dim btn As MSForms.CommandButton
    Dim pageItem As MSForms.Page
    Dim rowNum As Long
    Dim colNum As Long
    Dim leftPos As Single
    Dim topPos As Single

    For Each pageItem In mpItemTabs.Pages
        If UCase$(Trim$(pageItem.caption)) = "PHOTOS" Then
            Set pg = pageItem
            Exit For
        End If
    Next pageItem
    If pg Is Nothing Then Exit Sub

    For i = pg.Controls.Count - 1 To 0 Step -1
        If Left$(pg.Controls(i).Name, 9) = "imgPhoto_" _
            Or Left$(pg.Controls(i).Name, 9) = "lblPhoto_" _
            Or Left$(pg.Controls(i).Name, 15) = "btnRemovePhoto_" Then
            pg.Controls.Remove pg.Controls(i).Name
        End If
    Next i

    For i = 1 To mSelectedPhotoPaths.Count
        rowNum = Int((i - 1) / 4)
        colNum = (i - 1) Mod 4
        leftPos = 18 + (colNum * 145)
        topPos = 82 + (rowNum * 92)
        Set img = pg.Controls.Add("Forms.Image.1", "imgPhoto_" & i)
        With img
            .Left = leftPos
            .Top = topPos
            .Width = 92
            .Height = 64
            .PictureSizeMode = fmPictureSizeModeZoom
            On Error Resume Next
            Set .Picture = LoadPicture(CStr(mSelectedPhotoPaths(i)))
            On Error GoTo 0
        End With
        Set lbl = pg.Controls.Add("Forms.Label.1", "lblPhoto_" & i)
        With lbl
            .caption = Mid$(CStr(mSelectedPhotoPaths(i)), InStrRev(CStr(mSelectedPhotoPaths(i)), "\") + 1)
            .Left = leftPos
            .Top = topPos + 66
            .Width = 130
            .Height = 12
            .Font.Size = 7
        End With
        Set btn = pg.Controls.Add("Forms.CommandButton.1", "btnRemovePhoto_" & i)
        With btn
            .caption = "X"
            .Left = leftPos + 74
            .Top = topPos
            .Width = 18
            .Height = 16
        End With
        RegisterPhotoButtonEvent btn, "REMOVE_ONE", i
    Next i

End Sub


' =====================================================
' MARK DIRTY
' =====================================================

Public Sub MarkDirty()

    mIsDirty = True

    lblStatus.caption = "Unsaved Changes"

End Sub


' =====================================================
' HEADER DISPLAY
' =====================================================

Private Sub RefreshHeaderDisplay()

    Dim displayName As String

    lblItemNumber.caption = mItemNumber
    lblItemNumber.Width = 72

    displayName = mItemName
    If Len(displayName) > 72 Then displayName = Left$(displayName, 72) & "..."
    lblItemName.caption = displayName
    lblItemName.Left = lblItemNumber.Left + lblItemNumber.Width + 6
    lblItemName.Width = 480
    lblItemName.Height = lblItemNumber.Height
    lblItemName.WordWrap = False

End Sub

Private Function GetTabIndexByCaption(ByVal tabCaption As String) As Long

    Dim i As Long

    For i = 0 To mpItemTabs.Pages.Count - 1
        If UCase$(Trim$(mpItemTabs.Pages(i).caption)) = UCase$(Trim$(tabCaption)) Then
            GetTabIndexByCaption = i
            Exit Function
        End If
    Next i

End Function


' =====================================================
' CURRENT DATA ROW
' =====================================================

Public Property Get CurrentDataRow() As Long

    CurrentDataRow = mCurrentDataRow

End Property

Public Property Let CurrentDataRow(ByVal Value As Long)

    mCurrentDataRow = Value

End Property


' =====================================================
' ITEM NUMBER
' =====================================================

Public Property Get itemNumber() As String

    itemNumber = mItemNumber

End Property

Public Property Let itemNumber(ByVal Value As String)

    mItemNumber = Value

    RefreshHeaderDisplay

End Property


' =====================================================
' ITEM NAME
' =====================================================

Public Property Get itemName() As String

    itemName = mItemName

End Property

Public Property Let itemName(ByVal Value As String)

    mItemName = Value

    RefreshHeaderDisplay

End Property


' =====================================================
' DIRTY STATE
' =====================================================

Public Property Get IsDirty() As Boolean

    IsDirty = mIsDirty

End Property

