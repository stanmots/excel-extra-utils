VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditCopyingConfigForm 
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   OleObjectBlob   =   "EditCopyingConfigForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditCopyingConfigForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_IsNeededAdding As Boolean
Private m_UpdatedListBoxRow As Long

Public Property Get IsNeededAdding() As Boolean

IsNeededAdding = m_IsNeededAdding
    
End Property

Public Property Get UpdatedListBoxRow() As Long

UpdatedListBoxRow = m_UpdatedListBoxRow
    
End Property

Public Property Let IsNeededAdding(ByVal IsNeededAddingFlag As Boolean)

m_IsNeededAdding = IsNeededAddingFlag
    
End Property

Public Property Let UpdatedListBoxRow(ByVal UpdatedListBoxRow As Long)

m_UpdatedListBoxRow = UpdatedListBoxRow
    
End Property

Private Sub UserForm_Initialize()

m_UpdatedListBoxRow = -1

Me.Caption = EDIT_COPYING_CONFIG_FORM_TITLE

'labels
Me.FromWorksheetCopyingSettingsLabel.Caption = FROM_WORKSHEET_COPYING_SETTINGS_LABEL
Me.FromWorksheetNameLabel.Caption = WORKSHEET_NAME_LABEL
Me.FromWorkbookNameLabel.Caption = WORKBOOK_NAME_LABEL
Me.FromBaseCopyingCellLabel.Caption = BASE_CELL_LABEL
Me.FromCopyingRangeLabel.Caption = COPYING_RANGE_LABEL
Me.ToWorksheetCopyingSettingsLabel.Caption = TO_WORKSHEET_COPYING_SETTINGS_LABEL
Me.ToWorksheetNameLabel.Caption = WORKSHEET_NAME_LABEL
Me.ToWorkbookNameLabel.Caption = WORKBOOK_NAME_LABEL
Me.ToBaseCopyingCellLabel.Caption = BASE_CELL_LABEL
Me.ToCopyingRangeLabel.Caption = COPYING_RANGE_LABEL
Me.CopyingColorLabel.Caption = COPYING_COLOR_LABEL & ":"
Me.CommonCopyingSettingLabel.Caption = COMMON_COPYING_LABEL
Me.CopyingPasteTypeLabel.Caption = COPYING_PASTE_TYPE_LABEL & ":"
Me.CopyingSpecialOperationLabel.Caption = COPYING_SPECIAL_OPERATION_LABEL & ":"

'fill special operations combo-box
Me.CopyingSpecialOperationsComboBox.AddItem "xlPasteSpecialOperationNone"
Me.CopyingSpecialOperationsComboBox.AddItem "xlPasteSpecialOperationAdd"
Me.CopyingSpecialOperationsComboBox.AddItem "xlPasteSpecialOperationSubtract"
Me.CopyingSpecialOperationsComboBox.AddItem "xlPasteSpecialOperationMultiply"
Me.CopyingSpecialOperationsComboBox.AddItem "xlPasteSpecialOperationDivide"

'fill paste types combo-box
Me.CopyingPasteTypesComboBox.AddItem "xlPasteAll"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteAllExceptBorders"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteAllMergingConditionalFormats"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteAllUsingSourceTheme"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteColumnWidths"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteComments"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteFormats"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteFormulas"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteFormulasAndNumberFormats"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteValidation"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteValues"
Me.CopyingPasteTypesComboBox.AddItem "xlPasteValuesAndNumberFormats"

'buttons
Me.CancelCopyingSettingButton.Caption = CANCEL_BUTTON_TITLE
Me.SaveCopyingSettingButton.Caption = SAVE_BUTTON_TITLE
Me.SetCopyingColorButton.Caption = SET_BUTTON_TITLE
Me.ClearCopyingColorButton.Caption = CLEAR_BUTTON_TITLE
Me.FromWBBrowseButton.Caption = BROWSE_BUTTON_TITLE & "..."
Me.ToWBBrowseButton.Caption = BROWSE_BUTTON_TITLE & "..."

'combo-boxes default values
Me.CopyingSpecialOperationsComboBox.text = Me.CopyingSpecialOperationsComboBox.List(0)
Me.CopyingPasteTypesComboBox.text = Me.CopyingPasteTypesComboBox.List(0)

End Sub

Private Sub FromWBBrowseButton_Click()

Dim TestVar As Variant
TestVar = Application.GetOpenFilename(FileFilter:=EXCEL_FILE_FILTER, Title:=GET_PATH_TO_FILE_TITLE, MultiSelect:=False)
If TestVar <> False Then
    Me.InputFromWorkbookNameTextBox.text = TestVar
End If

End Sub

Private Sub ToWBBrowseButton_Click()

Dim TestVar As Variant
TestVar = Application.GetOpenFilename(FileFilter:=EXCEL_FILE_FILTER, Title:=GET_PATH_TO_FILE_TITLE, MultiSelect:=False)
If TestVar <> False Then
    Me.InputToWorkbookNameTextBox.text = TestVar
End If

End Sub

Private Sub SetCopyingColorButton_Click()

Dim ColorCode As Long

If Len(Me.CopyingColorTextBox.text) > 0 Then
ColorCode = CommonHelpers.ShowEditColorDialog(CLng(Me.CopyingColorTextBox.text))
Else
ColorCode = CommonHelpers.ShowEditColorDialog
End If

If ColorCode <> -1 Then

Me.CopyingColorTextBox.text = ColorCode

End If

End Sub

Private Sub ClearCopyingColorButton_Click()

Me.CopyingColorTextBox.text = ""

End Sub

Private Function IsInputValuesCorrect() As Boolean

IsInputValuesCorrect = True

'check FromWorksheet Settings
If Len(Me.InputFromWorksheetNameTextBox.text) = 0 Or Len(Me.InputFromBaseCopyingCellTextBox.text) = 0 Or _
    Len(Me.InputFromTopLeftCopyingCellTextBox.text) = 0 Or IsValidRange(Me.InputFromTopLeftCopyingCellTextBox.text) = False _
    Then
    IsInputValuesCorrect = False
    Exit Function
End If

'check ToWorksheet Settings
If Len(Me.InputToWorksheetNameTextBox.text) = 0 Or Len(Me.InputToBaseCopyingCellTextBox.text) = 0 Or _
    Len(Me.InputToTopLeftCopyingCellTextBox.text) = 0 Or IsValidRange(Me.InputToTopLeftCopyingCellTextBox.text) = False _
    Then
    IsInputValuesCorrect = False
    Exit Function
End If

'check ranges
Dim FromRangeStr As String
FromRangeStr = Me.InputFromTopLeftCopyingCellTextBox.text

If Len(Me.InputFromRightBottomCopyingCellTextBox.text) > 0 And IsValidRange(Me.InputFromRightBottomCopyingCellTextBox.text) = True Then
    FromRangeStr = FromRangeStr & ":" & Me.InputFromRightBottomCopyingCellTextBox.text
End If

Dim ToRangeStr As String
ToRangeStr = Me.InputToTopLeftCopyingCellTextBox.text

If Len(Me.InputToRightBottomCopyingCellTextBox.text) > 0 And IsValidRange(Me.InputToRightBottomCopyingCellTextBox.text) = True Then
    ToRangeStr = ToRangeStr & ":" & Me.InputToRightBottomCopyingCellTextBox.text
End If

If CopyingSettingsHelpers.AreCopyingRangesValid(FromRangeStr, ToRangeStr) = CopyingRangesType.INCORRECT_RANGES Then
    IsInputValuesCorrect = False
    Exit Function
End If

End Function

Private Sub SaveCopyingSettingButton_Click()

If IsInputValuesCorrect = False Then
    MsgBox INCORRECT_INPUT_VALUES_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
Else
    SaveCurrentSettings
    DestroyObject
End If

End Sub

Private Sub SaveCurrentSettings()

'[1] update current config name
If CopyingSettingsForm.UpdateCurrentConfigName() = False Then
    Debug.Print "Error in [SaveCurrentSettings]! Cannot update the config name."
    Exit Sub
End If

'[2] save current settings
'[2][1] prepare list-box
Dim SettingNumber As Long

If m_IsNeededAdding = True Then
    
    'SettingNumber is a logical listbox row
    'm_UpdatedListBoxRow is a real listbox row
    SettingNumber = CLng(CopyingSettingsForm.CopyingSettingsListBox.ListCount / 2 + 1)

    'from-row
    CopyingSettingsForm.CopyingSettingsListBox.AddItem
    m_UpdatedListBoxRow = CopyingSettingsForm.CopyingSettingsListBox.ListCount - 1
    CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 0) = SettingNumber
    CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 1) = COPYING_DESTINATION_FROMTYPE
    
    'to-row
    CopyingSettingsForm.CopyingSettingsListBox.AddItem
    CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow + 1, 1) = COPYING_DESTINATION_TOTYPE
Else
    If m_UpdatedListBoxRow = -1 Then
        'cannot rename nonexistent item
        Debug.Print "Error! m_UpdatedListBoxRow variable in [SaveCurrentSettings] has uncorrect value!"
        Exit Sub
    End If
    
    'm_UpdatedListBoxRow must be set in the calling form,
    'as it is an even index,(m_UpdatedListBoxRow + 2) will be the list count (index is starting from 0)
    SettingNumber = CLng((m_UpdatedListBoxRow + 2) / 2)
End If

'[2][2] update main storage
'two stages to process: 1) from-worksheet settings; 2) to-worksheet settings
Dim StageNumber As Integer
Dim SectionName As String
SectionName = COPYING_SECTION_KEY & SECTION_DELIMITER & CopyingSettingsForm.CurrentStoredConfigName & SECTION_DELIMITER & CStr(SettingNumber)

With ExcelUtilsMainWindow.MainStorage
    
    '1 - FROM_SETTINGS
    'worksheet
    .StoreValue SectionName, COPYING_FROMWORKSHEET_KEY, Me.InputFromWorksheetNameTextBox.text
    
    'workbook
    .StoreValue SectionName, COPYING_FROMWORKBOOK_KEY, Me.InputFromWorkbookNameTextBox.text

    'basecell
    .StoreValue SectionName, COPYING_FROMBASECELL_KEY, Me.InputFromBaseCopyingCellTextBox.text
    
    'range
    Dim FromRange As String
    FromRange = Me.InputFromTopLeftCopyingCellTextBox.text

    If Len(Me.InputFromRightBottomCopyingCellTextBox.text) > 0 Then
        FromRange = FromRange & ":" & Me.InputFromRightBottomCopyingCellTextBox.text
    End If
    
    .StoreValue SectionName, COPYING_FROMRANGE_KEY, FromRange
    
    'column
    Dim FromColumn As String
    FromColumn = ColumnNumberToLetter(Range(Me.InputFromBaseCopyingCellTextBox.text).Column)
    .StoreValue SectionName, COPYING_FROMCOLUMN_KEY, FromColumn
    
    'offsets
    Dim TopLeftCellRowOffset As Long, TopLeftCellColumnOffset As Long
    Dim FromCopyingOffsets As String
    
    TopLeftCellColumnOffset = Range(Me.InputFromTopLeftCopyingCellTextBox.text).Column - Range(Me.InputFromBaseCopyingCellTextBox.text).Column
    .StoreValue SectionName, COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY, TopLeftCellColumnOffset
    
    TopLeftCellRowOffset = Range(Me.InputFromTopLeftCopyingCellTextBox.text).Row - Range(Me.InputFromBaseCopyingCellTextBox.text).Row
    .StoreValue SectionName, COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY, TopLeftCellRowOffset
    
    Dim FromRightBottomCellRowOffset As Variant, FromRightBottomCellColumnOffset As Variant
    
    If Len(Me.InputFromRightBottomCopyingCellTextBox.text) > 0 Then
        
        FromRightBottomCellColumnOffset = Range(Me.InputFromRightBottomCopyingCellTextBox.text).Column - Range(Me.InputFromBaseCopyingCellTextBox.text).Column
        .StoreValue SectionName, COPYING_FROM_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY, FromRightBottomCellColumnOffset
        
        FromRightBottomCellRowOffset = Range(Me.InputFromRightBottomCopyingCellTextBox.text).Row - Range(Me.InputFromBaseCopyingCellTextBox.text).Row
        .StoreValue SectionName, COPYING_FROM_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY, FromRightBottomCellRowOffset
    Else
        FromRightBottomCellColumnOffset = ""
        FromRightBottomCellRowOffset = ""
    End If
    
    FromCopyingOffsets = CommonHelpers.GetFormattedStringFromOffsets(TopLeftCellColumnOffset, TopLeftCellRowOffset, _
        FromRightBottomCellColumnOffset, FromRightBottomCellRowOffset)
        
    '2 - TO_SETTINGS
    'worksheet
    .StoreValue SectionName, COPYING_TOWORKSHEET_KEY, Me.InputToWorksheetNameTextBox.text
    
    'workbook
    .StoreValue SectionName, COPYING_TOWORKBOOK_KEY, Me.InputToWorkbookNameTextBox.text
 
    'basecell
    .StoreValue SectionName, COPYING_TOBASECELL_KEY, Me.InputToBaseCopyingCellTextBox.text
    
    'range
    Dim ToRange As String
    ToRange = Me.InputToTopLeftCopyingCellTextBox.text
    
    If Len(Me.InputToRightBottomCopyingCellTextBox.text) > 0 Then
        ToRange = ToRange & ":" & Me.InputToRightBottomCopyingCellTextBox.text
    End If
    
    .StoreValue SectionName, COPYING_TORANGE_KEY, ToRange
    
    'column
    Dim ToColumn As String
    ToColumn = ColumnNumberToLetter(Range(Me.InputToBaseCopyingCellTextBox.text).Column)
    .StoreValue SectionName, COPYING_TOCOLUMN_KEY, ToColumn
    
    'offsets
    Dim ToCopyingOffsets As String
    
    TopLeftCellColumnOffset = Range(Me.InputToTopLeftCopyingCellTextBox.text).Column - Range(Me.InputToBaseCopyingCellTextBox.text).Column
    .StoreValue SectionName, COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY, TopLeftCellColumnOffset
    
    TopLeftCellRowOffset = Range(Me.InputToTopLeftCopyingCellTextBox.text).Row - Range(Me.InputToBaseCopyingCellTextBox.text).Row
    .StoreValue SectionName, COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY, TopLeftCellRowOffset
    
    Dim ToRightBottomCellRowOffset As Variant, ToRightBottomCellColumnOffset As Variant
    
    If Len(Me.InputToRightBottomCopyingCellTextBox.text) > 0 Then
        
        ToRightBottomCellColumnOffset = Range(Me.InputToRightBottomCopyingCellTextBox.text).Column - Range(Me.InputToBaseCopyingCellTextBox.text).Column
        .StoreValue SectionName, COPYING_TO_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY, ToRightBottomCellColumnOffset
        
        ToRightBottomCellRowOffset = Range(Me.InputToRightBottomCopyingCellTextBox.text).Row - Range(Me.InputToBaseCopyingCellTextBox.text).Row
        .StoreValue SectionName, COPYING_TO_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY, ToRightBottomCellRowOffset
    Else
        ToRightBottomCellColumnOffset = ""
        ToRightBottomCellRowOffset = ""
    End If
    
    ToCopyingOffsets = CommonHelpers.GetFormattedStringFromOffsets(TopLeftCellColumnOffset, TopLeftCellRowOffset, _
        ToRightBottomCellColumnOffset, ToRightBottomCellRowOffset)
    
    'common settings
    'paste special operation
    Dim xlPasteSpecialOperation As Long
    xlPasteSpecialOperation = CopyingSettingsHelpers.GetXlPasteSpecialOperationNumber(Me.CopyingSpecialOperationsComboBox.text)
    .StoreValue SectionName, XL_SPECIAL_OPERATION_KEY, xlPasteSpecialOperation
    
    'paste type
    Dim xlPasteType As Long
    xlPasteType = CopyingSettingsHelpers.GetXlPasteTypeNumber(Me.CopyingPasteTypesComboBox.text)
    .StoreValue SectionName, XL_PASTE_TYPE_KEY, xlPasteType
    
    'color
    If Len(Me.CopyingColorTextBox.text) > 0 Then
        .StoreValue SectionName, COPYING_COLOR_KEY, Me.CopyingColorTextBox.text
    End If
     
    Dim PasteParameters As String
    PasteParameters = CopyingSettingsHelpers.GetPasteParametersString(Me.CopyingPasteTypesComboBox.text, Me.CopyingSpecialOperationsComboBox.text)

End With

'[3] update list-box
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 2) = FromColumn
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 3) = Me.InputFromBaseCopyingCellTextBox.text & ";" & FromRange & FromCopyingOffsets
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 4) = Me.InputFromWorksheetNameTextBox.text
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 5) = CommonHelpers.ExtractFileNameFromPath(Me.InputToWorkbookNameTextBox.text)
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 6) = PasteParameters
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow, 7) = Me.CopyingColorTextBox.text

CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow + 1, 2) = ToColumn
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow + 1, 3) = Me.InputToBaseCopyingCellTextBox.text & ";" & ToRange & ToCopyingOffsets
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow + 1, 4) = Me.InputToWorksheetNameTextBox.text
CopyingSettingsForm.CopyingSettingsListBox.List(m_UpdatedListBoxRow + 1, 5) = CommonHelpers.ExtractFileNameFromPath(Me.InputToWorkbookNameTextBox.text)

End Sub

Private Sub DestroyObject()

FormsHelpers.ChangeStateOfAllControlsOnForm CopyingSettingsForm, True
Unload Me

End Sub

Private Sub CancelCopyingSettingButton_Click()

DestroyObject

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub
