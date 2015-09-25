VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyingSettingsForm 
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   OleObjectBlob   =   "CopyingSettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CopyingSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_CurrentStoredConfigName As String
Private m_IsNeededAdding As Boolean
Private m_PreviousPairedItemIndex As Long
Private m_IsNeededProceedChangeEvent As Boolean

Private Sub CopyingSettingsListBox_Change()

'prevent infinite loop
If m_IsNeededProceedChangeEvent = False Then
    m_IsNeededProceedChangeEvent = True
    Exit Sub
End If

Dim FirstPairedItemIndex As Long, SecondPairedItemIndex As Long
Dim SelectionMode As Boolean

'select two listbox rows based on the current click
FirstPairedItemIndex = Me.CopyingSettingsListBox.ListIndex
SelectionMode = Me.CopyingSettingsListBox.Selected(FirstPairedItemIndex)

If FirstPairedItemIndex Mod 2 = 0 Then
    SecondPairedItemIndex = FirstPairedItemIndex + 1
Else
    SecondPairedItemIndex = FirstPairedItemIndex - 1
End If

'change the paired item state
If Me.CopyingSettingsListBox.Selected(SecondPairedItemIndex) <> SelectionMode Then
    'we don't need change event to be invoked at this stage
    m_IsNeededProceedChangeEvent = False
    Me.CopyingSettingsListBox.Selected(SecondPairedItemIndex) = SelectionMode
End If

'check if change event is invoked for the first time
If m_PreviousPairedItemIndex = -1 Then
    m_PreviousPairedItemIndex = FirstPairedItemIndex
    Exit Sub
End If

'clear the previous settings out
If m_PreviousPairedItemIndex <> FirstPairedItemIndex And m_PreviousPairedItemIndex <> SecondPairedItemIndex Then
    m_IsNeededProceedChangeEvent = False
    Me.CopyingSettingsListBox.Selected(m_PreviousPairedItemIndex) = False
    If m_PreviousPairedItemIndex Mod 2 = 0 Then
        m_IsNeededProceedChangeEvent = False
        Me.CopyingSettingsListBox.Selected(m_PreviousPairedItemIndex + 1) = False
    Else
        m_IsNeededProceedChangeEvent = False
        Me.CopyingSettingsListBox.Selected(m_PreviousPairedItemIndex - 1) = False
    End If

    m_PreviousPairedItemIndex = FirstPairedItemIndex
End If

End Sub

Private Sub UserForm_Initialize()

Me.Caption = COPYING_SETTINGS_FORM_TITLE

m_IsNeededAdding = False
m_IsNeededProceedChangeEvent = True
m_PreviousPairedItemIndex = -1

'labels
Me.InputCopyingConfigNameLabel.Caption = INPUT_COPYING_CONFIG_NAME_LABEL
Me.CopyingListBoxDescriptionLabel.Caption = COPYING_LIST_BOX_DESCRIPTION_LABEL
Me.CopyingDirectionLabel.Caption = COPYING_DIRECTION_LABEL
Me.CopyingColumnLabel.Caption = COPYING_COLUMN_LABEL
Me.CopyingOffsetsLabel.Caption = COPYING_OFFSETS_LABEL
Me.CopyingWorksheetLabel.Caption = COPYING_WORKSHEET_LABEL
Me.CopyingWorkbookLabel.Caption = COPYING_WORKBOOK_LABEL
Me.CopyingPasteParametersLabel.Caption = COPYING_PASTE_PARAMETERS_LABEL
Me.CopyingColorLabel.Caption = COPYING_COLOR_LABEL

'buttons
Me.CancelCopyingConfigsButton.Caption = CANCEL_BUTTON_TITLE
Me.SaveCopyingConfigsButton.Caption = SAVE_BUTTON_TITLE
Me.AddConfigSettingButton.Caption = ADD_BUTTON_TITLE
Me.DeleteConfigSettingButton.Caption = DELETE_BUTTON_TITLE
Me.EditConfigSettingButton.Caption = EDIT_BUTTON_TITLE
Me.CopyingGlobalSettingsButton.Caption = GLOBAL_SETTINGS_BUTTON_TITLE

End Sub

Public Property Get CurrentStoredConfigName() As String

CurrentStoredConfigName = m_CurrentStoredConfigName
    
End Property

Public Property Get IsNeededAdding() As Boolean

IsNeededAdding = m_IsNeededAdding
    
End Property

Public Property Let CurrentStoredConfigName(ByVal ConfigName As String)

m_CurrentStoredConfigName = ConfigName
    
End Property

Public Property Let IsNeededAdding(ByVal IsNeededAddingFlag As Boolean)

m_IsNeededAdding = IsNeededAddingFlag
    
End Property

Public Function UpdateCurrentConfigName() As Boolean

UpdateCurrentConfigName = False

If Len(m_CurrentStoredConfigName) > 0 Then
    If m_CurrentStoredConfigName <> Me.InputCopyingConfigNameTextBox.text Then
        If RenameSetting(COPYING_SECTION_KEY & SECTION_DELIMITER & m_CurrentStoredConfigName, Me.InputCopyingConfigNameTextBox.text) = False Then Exit Function
    End If
Else
    If AddSetting(Me.InputCopyingConfigNameTextBox.text) = False Then Exit Function
End If

m_IsNeededAdding = False
m_CurrentStoredConfigName = Me.InputCopyingConfigNameTextBox.text

UpdateCurrentConfigName = True

End Function

Private Sub CopyingGlobalSettingsButton_Click()

If Len(Me.InputCopyingConfigNameTextBox.text) = 0 Then
    MsgBox INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
End If

FormsHelpers.ChangeStateOfAllControlsOnForm CopyingSettingsForm, False
If CopyingGlobalSettingsForm.Visible = False Then CopyingGlobalSettingsForm.Show

End Sub

Private Sub AddConfigSettingButton_Click()

If Len(Me.InputCopyingConfigNameTextBox.text) = 0 Then
    MsgBox INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
Else
    RecoverGlobalSettings
    FormsHelpers.PrepareAndShowCopyingForm EditCopyingConfigForm, CopyingSettingsForm, _
            True, False
End If

End Sub

Private Sub DeleteConfigSettingButton_Click()

Dim ReturnedChoice As Integer

ReturnedChoice = MsgBox(DELETE_CONFIRMATION, vbOKCancel, ATTENTION_TITLE)
If ReturnedChoice = vbCancel Then
    Exit Sub
End If

If ListBoxFunctions.IsListBoxHasSelectedItems(Me.CopyingSettingsListBox) = False Then
    MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
Else
    If Len(m_CurrentStoredConfigName) = 0 Then
        MsgBox INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
End If

'delete from the listbox
Dim BaseIndex As Long
If m_PreviousPairedItemIndex Mod 2 = 0 Then
    BaseIndex = m_PreviousPairedItemIndex
    Me.CopyingSettingsListBox.RemoveItem m_PreviousPairedItemIndex
    'the next item will have the same index after removing
    Me.CopyingSettingsListBox.RemoveItem m_PreviousPairedItemIndex
Else
    BaseIndex = m_PreviousPairedItemIndex - 1
    Me.CopyingSettingsListBox.RemoveItem m_PreviousPairedItemIndex
    Me.CopyingSettingsListBox.RemoveItem m_PreviousPairedItemIndex - 1
End If

'delete from the storage
ExcelUtilsMainWindow.MainStorage.DeleteSection COPYING_SECTION_KEY & SECTION_DELIMITER & m_CurrentStoredConfigName & SECTION_DELIMITER & CStr(CLng(BaseIndex / 2 + 1))

If Me.CopyingSettingsListBox.ListCount > 0 Then
    Dim i As Long
    For i = BaseIndex To Me.CopyingSettingsListBox.ListCount - 2 Step 2
        ExcelUtilsMainWindow.MainStorage.RenameSection COPYING_SECTION_KEY & SECTION_DELIMITER & m_CurrentStoredConfigName & SECTION_DELIMITER & _
                Me.CopyingSettingsListBox.List(i, 0), Me.CopyingSettingsListBox.List(i, 0) - 1
        Me.CopyingSettingsListBox.List(i, 0) = Me.CopyingSettingsListBox.List(i, 0) - 1
    Next i
End If

m_PreviousPairedItemIndex = -1

End Sub

Private Sub RecoverGlobalSettings(Optional ByVal SettingNumber As Long = -1)

If Len(m_CurrentStoredConfigName) > 0 Then

    Dim SectionName As String
    SectionName = COPYING_SECTION_KEY & SECTION_DELIMITER & m_CurrentStoredConfigName
    
    'from-worksheet
    Dim FromWorksheetName As String: FromWorksheetName = ""
    Dim GlobalFWSIsUsed As Boolean
    GlobalFWSIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FWS_IS_USED_KEY, False)
    
    If GlobalFWSIsUsed = True Then
        'use global worksheet name
        FromWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FROMWORKSHEET_NAME_KEY, "")
    Else
        If SettingNumber <> -1 Then
            'use local worksheet name
            FromWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName & SECTION_DELIMITER & CStr(SettingNumber), COPYING_FROMWORKSHEET_KEY, "")
        End If
    End If
    
    EditCopyingConfigForm.InputFromWorksheetNameTextBox.text = FromWorksheetName
    EditCopyingConfigForm.InputFromWorksheetNameTextBox.Enabled = Not GlobalFWSIsUsed
    
    'from-workbook
    Dim FromWorkbookName As String: FromWorkbookName = ""
    Dim GlobalFWBIsUsed As Boolean
    GlobalFWBIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FWB_IS_USED_KEY, False)
    
    If GlobalFWBIsUsed = True Then
        'use global worksheet name
        FromWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FROMWORKBOOK_NAME_KEY, "")
    Else
        If SettingNumber <> -1 Then
            'use local worksheet name
            FromWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName & SECTION_DELIMITER & CStr(SettingNumber), COPYING_FROMWORKBOOK_KEY, "")
        End If
    End If
    
    EditCopyingConfigForm.InputFromWorkbookNameTextBox.text = FromWorkbookName
    EditCopyingConfigForm.InputFromWorkbookNameTextBox.Enabled = Not GlobalFWBIsUsed
    EditCopyingConfigForm.FromWBBrowseButton.Enabled = Not GlobalFWBIsUsed
    
    'to-worksheet
    Dim ToWorksheetName As String: ToWorksheetName = ""
    Dim GlobalTWSIsUsed As Boolean
    GlobalTWSIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TWS_IS_USED_KEY, False)
    
    If GlobalTWSIsUsed = True Then
        'use global worksheet name
        ToWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TOWORKSHEET_NAME_KEY, "")
    Else
        If SettingNumber <> -1 Then
            'use local worksheet name
            ToWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName & SECTION_DELIMITER & CStr(SettingNumber), COPYING_TOWORKSHEET_KEY, "")
        End If
    End If
    
    EditCopyingConfigForm.InputToWorksheetNameTextBox.text = ToWorksheetName
    EditCopyingConfigForm.InputToWorksheetNameTextBox.Enabled = Not GlobalTWSIsUsed
    
    'to-workbook
    Dim ToWorkbookName As String: ToWorkbookName = ""
    Dim GlobalTWBIsUsed As Boolean
    GlobalTWBIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TWB_IS_USED_KEY, False)
    
    If GlobalTWBIsUsed = True Then
        'use global worksheet name
        ToWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TOWORKBOOK_NAME_KEY, "")
    Else
        If SettingNumber <> -1 Then
            'use local worksheet name
            ToWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName & SECTION_DELIMITER & CStr(SettingNumber), COPYING_TOWORKBOOK_KEY, "")
        End If
    End If
    
    EditCopyingConfigForm.InputToWorkbookNameTextBox.text = ToWorkbookName
    EditCopyingConfigForm.InputToWorkbookNameTextBox.Enabled = Not GlobalTWBIsUsed
    EditCopyingConfigForm.ToWBBrowseButton.Enabled = Not GlobalTWBIsUsed

Else
    
    EditCopyingConfigForm.InputFromWorksheetNameTextBox.text = ""
    EditCopyingConfigForm.InputToWorksheetNameTextBox.text = ""
    EditCopyingConfigForm.InputFromWorkbookNameTextBox.text = ""
    EditCopyingConfigForm.InputToWorkbookNameTextBox.text = ""

End If

End Sub

Private Sub EditConfigSettingButton_Click()

If Len(Me.InputCopyingConfigNameTextBox.text) = 0 Then
    MsgBox INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
End If

If ListBoxFunctions.IsListBoxHasSelectedItems(Me.CopyingSettingsListBox) = False Then
    MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
End If

Dim BaseIndex As Long

'm_PreviousPairedItemIndex is the currently selected item
If m_PreviousPairedItemIndex Mod 2 = 0 Then
    BaseIndex = m_PreviousPairedItemIndex
Else
    BaseIndex = m_PreviousPairedItemIndex - 1
End If

'BaseIndex is an even index of the currently selected item in CopyingSettingsListBox
EditCopyingConfigForm.UpdatedListBoxRow = BaseIndex

'recover global settings (sheets, books etc.)
RecoverGlobalSettings (BaseIndex + 2) / 2

'recover base cell and offsets
Dim StageNumber As Integer
'we need two stages to process: 1) from-worksheet settings; 2) to-worksheet settings
For StageNumber = 0 To 1

    Dim OffsetsCell As String, BaseCell As String
    OffsetsCell = Me.CopyingSettingsListBox.List(BaseIndex + StageNumber, 3)
    
    Dim NestedCells() As String
    NestedCells = Split(OffsetsCell, DELIMITER:="(")
    NestedCells = Split(NestedCells(0), DELIMITER:=";")
    
    BaseCell = NestedCells(0)
    
    Dim TLCell As String, RBCell As String
    If InStr(2, NestedCells(1), ":") = 0 Then
        TLCell = NestedCells(1)
        RBCell = ""
    Else
        NestedCells = Split(NestedCells(1), DELIMITER:=":")
        TLCell = NestedCells(0)
        RBCell = NestedCells(1)
    End If
    
    Select Case StageNumber

    Case SettingsType.FROM_WORKSHEET_TYPE:
        EditCopyingConfigForm.InputFromBaseCopyingCellTextBox.text = BaseCell
        EditCopyingConfigForm.InputFromTopLeftCopyingCellTextBox.text = TLCell
        EditCopyingConfigForm.InputFromRightBottomCopyingCellTextBox.text = RBCell
        
    Case SettingsType.TO_WORKSHEET_TYPE:
        EditCopyingConfigForm.InputToBaseCopyingCellTextBox.text = BaseCell
        EditCopyingConfigForm.InputToTopLeftCopyingCellTextBox.text = TLCell
        EditCopyingConfigForm.InputToRightBottomCopyingCellTextBox.text = RBCell
        
    End Select
    
Next StageNumber
    
'recover comboboxes
Dim NestedPasteParameters() As String
NestedPasteParameters = Split(CStr(Me.CopyingSettingsListBox.List(BaseIndex, 6)), DELIMITER:=PASTE_PARAMETERS_DELIMITER)
    
Dim SpecialOperationComboBoxSelectedItem As Integer
SpecialOperationComboBoxSelectedItem = CopyingSettingsHelpers.GetSpecialOperationsComboBoxItemSerialNumberFromString("xlPasteSpecialOperation" & NestedPasteParameters(1))

If SpecialOperationComboBoxSelectedItem = -1 Then
    Debug.Print "Error! SpecialOperationComboBoxSelectedItem in [EditConfigSettingButton_Click] has uncorrect value!"
    Exit Sub
End If

EditCopyingConfigForm.CopyingSpecialOperationsComboBox.text = EditCopyingConfigForm.CopyingSpecialOperationsComboBox.List(SpecialOperationComboBoxSelectedItem)

Dim PasteTypeComboBoxSelectedItem As Integer
PasteTypeComboBoxSelectedItem = CopyingSettingsHelpers.GetPasteTypesComboBoxItemSerialNumberFromString("xlPaste" & NestedPasteParameters(0))

If PasteTypeComboBoxSelectedItem = -1 Then
    Debug.Print "Error! PasteTypeComboBoxSelectedItem in [EditConfigSettingButton_Click] has uncorrect value!"
    Exit Sub
End If

EditCopyingConfigForm.CopyingPasteTypesComboBox.text = EditCopyingConfigForm.CopyingPasteTypesComboBox.List(PasteTypeComboBoxSelectedItem)

FormsHelpers.PrepareAndShowCopyingForm EditCopyingConfigForm, CopyingSettingsForm, _
        False, False
End Sub

Private Sub SaveCopyingConfigsButton_Click()

If Len(Me.InputCopyingConfigNameTextBox.text) = 0 Then
    MsgBox INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
End If

If m_IsNeededAdding = True Then
    If AddSetting(Me.InputCopyingConfigNameTextBox.text) = False Then Exit Sub
Else
    If Len(m_CurrentStoredConfigName) > 0 Then
        If m_CurrentStoredConfigName <> Me.InputCopyingConfigNameTextBox.text Then
            If RenameSetting(COPYING_SECTION_KEY & SECTION_DELIMITER & m_CurrentStoredConfigName, Me.InputCopyingConfigNameTextBox.text) = False Then Exit Sub
        End If
    Else
        MsgBox CANNOT_REPLACE_ITEM_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
End If

DestroyObject

End Sub

'SettingName is a main SubSection name (without nested subsections)
Private Function AddSetting(ByVal SettingName As String) As Boolean

AddSetting = False

If ListBoxFunctions.IsListBoxHasItem(SettingName, _
    ExcelUtilsMainWindow.MainMultiPageObject.Pages("CopyingPage").CopyingConfigsListBox) = True Then
    MsgBox LISTBOX_ALREADY_HAS_ITEM_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Function
End If

With ExcelUtilsMainWindow.MainMultiPageObject.Pages("CopyingPage")
    ExcelUtilsMainWindow.MainStorage.AddSection COPYING_SECTION_KEY & SECTION_DELIMITER & SettingName
    .CopyingConfigsListBox.AddItem SettingName
    AddSetting = True
End With

End Function

'FromSettingName must be in a full form: "MainSection/SubSection/etc..."
'ToSettingName must contain only last subsection's title
Private Function RenameSetting(ByVal FromSettingName As String, ByVal ToSettingName As String) As Boolean

RenameSetting = False

If ListBoxFunctions.IsListBoxHasItem(ToSettingName, _
    ExcelUtilsMainWindow.MainMultiPageObject.Pages("CopyingPage").CopyingConfigsListBox) = True Then
    MsgBox LISTBOX_ALREADY_HAS_ITEM_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Function
End If

'ListBox displays a name without top-level sections (i.e. only the inner-most subsection)
Dim NestedSections() As String
Dim FromSettingNameInListBox As String
NestedSections = Split(FromSettingName, DELIMITER:=SECTION_DELIMITER)
FromSettingNameInListBox = NestedSections(UBound(NestedSections))

With ExcelUtilsMainWindow.MainMultiPageObject.Pages("CopyingPage")
    If ExcelUtilsMainWindow.MainStorage.RenameSection(FromSettingName, ToSettingName) = True And _
        ListBoxFunctions.RenameItem(FromSettingNameInListBox, ToSettingName, _
                                .CopyingConfigsListBox) = True Then
        RenameSetting = True
    End If
End With

End Function

Private Sub DestroyObject()

m_CurrentStoredConfigName = ""

'enable the controls on the parent form
FormsHelpers.ChangeStateOfAllControlsOnForm ExcelUtilsMainWindow, True

'unload child forms
If FormsHelpers.IsUserFormLoaded("EditCopyingConfigForm") = True Then Unload EditCopyingConfigForm
If FormsHelpers.IsUserFormLoaded("CopyingGlobalSettingsForm") = True Then Unload CopyingGlobalSettingsForm

'unload current form
Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub

Private Sub CancelCopyingConfigsButton_Click()

DestroyObject

End Sub
