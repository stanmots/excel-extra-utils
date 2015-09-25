VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyingGlobalSettingsForm 
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   OleObjectBlob   =   "CopyingGlobalSettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CopyingGlobalSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()

Me.Caption = COPYING_GLOBAL_SETTINGS_TITLE

'buttons
Me.CancelCopyingGlobalSettingsButton.Caption = CANCEL_BUTTON_TITLE
Me.SaveCopyingGlobalSettingsButton.Caption = SAVE_BUTTON_TITLE
Me.FromGlobalWBBrowseButton.Caption = BROWSE_BUTTON_TITLE & "..."
Me.ToGlobalWBBrowseButton.Caption = BROWSE_BUTTON_TITLE & "..."

'labels
Me.FromGlobalCopyingSettingsLabel.Caption = FROM_WORKSHEET_COPYING_SETTINGS_LABEL
Me.ToGlobalCopyingSettingsLabel.Caption = TO_WORKSHEET_COPYING_SETTINGS_LABEL
Me.FromGlobalWorkbookNameLabel.Caption = WORKBOOK_NAME_LABEL
Me.ToGlobalWorkbookNameLabel.Caption = WORKBOOK_NAME_LABEL
Me.UseCurrentGlobalSettingLabel.Caption = USE_CURRENT_GLOBAL_SETTING_LABEL
Me.GlobalSettingIsRemovedAfterCopyingLabel.Caption = IS_REMOVED_GLOBAL_SETTING_LABEL
Me.FromGlobalWorksheetNameLabel.Caption = WORKSHEET_NAME_LABEL
Me.ToGlobalWorksheetNameLabel.Caption = WORKSHEET_NAME_LABEL

'restore settings
If Len(CopyingSettingsForm.CurrentStoredConfigName) > 0 Then

    Dim SectionName As String
    SectionName = COPYING_SECTION_KEY & SECTION_DELIMITER & CopyingSettingsForm.CurrentStoredConfigName
    
    'from-workbook
    Dim GlobalFWBIsUsed As Boolean
    GlobalFWBIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FWB_IS_USED_KEY, False)
    
    Me.CurrentGlobalFWBIsUsedCheckBox.Value = GlobalFWBIsUsed
    CurrentGlobalFWBIsUsedCheckBox_CustomChangeEvent
       
    Me.GlobalFWBIsRemovedAfterCopyingCheckBox.Value = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FWB_IS_REMOVED_AFTER_COPYING_KEY, False)
    Me.InputGlobalFWBNameTextBox.text = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FROMWORKBOOK_NAME_KEY, "")
    
    'from-worksheet
    Dim GlobalFWSIsUsed As Boolean
    GlobalFWSIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FWS_IS_USED_KEY, False)
    
    Me.CurrentGlobalFWSIsUsedCheckBox.Value = GlobalFWSIsUsed
    CurrentGlobalFWSIsUsedCheckBox_CustomChangeEvent
    
    Me.InputGlobalFWSNameTextBox.text = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_FROMWORKSHEET_NAME_KEY, "")
    
    'to-workbook
    Dim GlobalTWBIsUsed As Boolean
    GlobalTWBIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TWB_IS_USED_KEY, False)
    
    Me.CurrentGlobalTWBIsUsedCheckBox.Value = GlobalTWBIsUsed
    CurrentGlobalTWBIsUsedCheckBox_CustomChangeEvent
    
    Me.GlobalTWBIsRemovedAfterCopyingCheckBox.Value = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TWB_IS_REMOVED_AFTER_COPYING_KEY, False)
    Me.InputGlobalTWBNameTextBox.text = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TOWORKBOOK_NAME_KEY, "")
    
    'to-worksheet
    Dim GlobalTWSIsUsed As Boolean
    GlobalTWSIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TWS_IS_USED_KEY, False)
    
    Me.CurrentGlobalTWSIsUsedCheckBox.Value = GlobalTWSIsUsed
    CurrentGlobalTWSIsUsedCheckBox_CustomChangeEvent
    
    Me.InputGlobalTWSNameTextBox.text = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, GLOBAL_TOWORKSHEET_NAME_KEY, "")
    
Else
    ChangeStateOfGlobalSettingCheckBoxes False
End If

End Sub

Private Sub SaveCopyingGlobalSettingsButton_Click()

'[1] errors-checking
If Me.CurrentGlobalFWSIsUsedCheckBox.Value = True And Len(Me.InputGlobalFWSNameTextBox.text) = 0 Or _
    Me.CurrentGlobalTWSIsUsedCheckBox.Value = True And Len(Me.InputGlobalTWSNameTextBox.text) = 0 Then
        
    MsgBox INCORRECT_INPUT_VALUES_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
End If

'[1][1] update current config name
If CopyingSettingsForm.UpdateCurrentConfigName() = False Then
    Debug.Print "Error in [SaveCopyingGlobalSettingsButton_Click]! Cannot update the config name."
    Exit Sub
End If

'set section name only after updating the config name
Dim SectionName As String
SectionName = COPYING_SECTION_KEY & SECTION_DELIMITER & CopyingSettingsForm.CurrentStoredConfigName


With ExcelUtilsMainWindow.MainStorage
    '[2] Save FROM_SETTINGS
    'workbook
    .StoreValue SectionName, GLOBAL_FWB_IS_USED_KEY, Me.CurrentGlobalFWBIsUsedCheckBox.Value
    
    .StoreValue SectionName, GLOBAL_FWB_IS_REMOVED_AFTER_COPYING_KEY, Me.GlobalFWBIsRemovedAfterCopyingCheckBox.Value
    .StoreValue SectionName, GLOBAL_FROMWORKBOOK_NAME_KEY, Me.InputGlobalFWBNameTextBox.text
    
    'worksheet
    .StoreValue SectionName, GLOBAL_FWS_IS_USED_KEY, Me.CurrentGlobalFWSIsUsedCheckBox.Value
    
    .StoreValue SectionName, GLOBAL_FROMWORKSHEET_NAME_KEY, Me.InputGlobalFWSNameTextBox.text
    
    '[3] Save TO_SETTINGS
    'workbook
    .StoreValue SectionName, GLOBAL_TWB_IS_USED_KEY, Me.CurrentGlobalTWBIsUsedCheckBox.Value
    
    .StoreValue SectionName, GLOBAL_TWB_IS_REMOVED_AFTER_COPYING_KEY, Me.GlobalTWBIsRemovedAfterCopyingCheckBox.Value
    .StoreValue SectionName, GLOBAL_TOWORKBOOK_NAME_KEY, Me.InputGlobalTWBNameTextBox.text
    
    'worksheet
    .StoreValue SectionName, GLOBAL_TWS_IS_USED_KEY, Me.CurrentGlobalTWSIsUsedCheckBox.Value
    
    .StoreValue SectionName, GLOBAL_TOWORKSHEET_NAME_KEY, Me.InputGlobalTWSNameTextBox.text

    '[4] update copying settings listbox
    Dim i As Long
    Dim CurrentSubSectionName As String
                
    For i = 0 To CopyingSettingsForm.CopyingSettingsListBox.ListCount - 1
        If ((i Mod 2) = 0) Then
            'i is even, update from-settings
            CurrentSubSectionName = CopyingSettingsForm.CopyingSettingsListBox.List(i, 0)
                
            CopyingSettingsForm.CopyingSettingsListBox.List(i, 5) = CommonHelpers.ExtractFileNameFromPath(Me.InputGlobalFWBNameTextBox.text)
            .StoreValue SectionName & SECTION_DELIMITER & CurrentSubSectionName, COPYING_FROMWORKBOOK_KEY, Me.InputGlobalFWBNameTextBox.text
            
            If Me.CurrentGlobalFWSIsUsedCheckBox.Value = True Then
                CopyingSettingsForm.CopyingSettingsListBox.List(i, 4) = Me.InputGlobalFWSNameTextBox.text
                .StoreValue SectionName & SECTION_DELIMITER & CurrentSubSectionName, COPYING_FROMWORKSHEET_KEY, Me.InputGlobalFWSNameTextBox.text
            End If
        Else
            'i is odd, update to-settings
            CurrentSubSectionName = CopyingSettingsForm.CopyingSettingsListBox.List(i - 1, 0)
            
            CopyingSettingsForm.CopyingSettingsListBox.List(i, 5) = CommonHelpers.ExtractFileNameFromPath(Me.InputGlobalTWBNameTextBox.text)
            .StoreValue SectionName & SECTION_DELIMITER & CurrentSubSectionName, COPYING_TOWORKBOOK_KEY, Me.InputGlobalTWBNameTextBox.text
            
            If Me.CurrentGlobalTWSIsUsedCheckBox.Value = True Then
                CopyingSettingsForm.CopyingSettingsListBox.List(i, 4) = Me.InputGlobalTWSNameTextBox.text
                .StoreValue SectionName & SECTION_DELIMITER & CurrentSubSectionName, COPYING_TOWORKSHEET_KEY, Me.InputGlobalTWSNameTextBox.text
            End If
        End If
    Next i

End With

'[5] close current form
DestroyObject

End Sub

Private Sub FromGlobalWBBrowseButton_Click()

Dim TestVar As Variant
TestVar = Application.GetOpenFilename(FileFilter:=EXCEL_FILE_FILTER, Title:=GET_PATH_TO_FILE_TITLE, MultiSelect:=False)
If TestVar <> False Then
    Me.InputGlobalFWBNameTextBox.text = TestVar
End If

End Sub

Private Sub ToGlobalWBBrowseButton_Click()

Dim TestVar As Variant
TestVar = Application.GetOpenFilename(FileFilter:=EXCEL_FILE_FILTER, Title:=GET_PATH_TO_FILE_TITLE, MultiSelect:=False)
If TestVar <> False Then
    Me.InputGlobalTWBNameTextBox.text = TestVar
End If

End Sub

Private Sub CurrentGlobalFWBIsUsedCheckBox_Click()

CurrentGlobalFWBIsUsedCheckBox_CustomChangeEvent

End Sub

Private Sub CurrentGlobalFWSIsUsedCheckBox_Click()

CurrentGlobalFWSIsUsedCheckBox_CustomChangeEvent

End Sub

Private Sub CurrentGlobalTWBIsUsedCheckBox_Click()

CurrentGlobalTWBIsUsedCheckBox_CustomChangeEvent

End Sub

Private Sub CurrentGlobalTWSIsUsedCheckBox_Click()

CurrentGlobalTWSIsUsedCheckBox_CustomChangeEvent

End Sub

Private Sub CurrentGlobalFWBIsUsedCheckBox_CustomChangeEvent()

Me.InputGlobalFWBNameTextBox.Enabled = Me.CurrentGlobalFWBIsUsedCheckBox.Value
Me.FromGlobalWBBrowseButton.Enabled = Me.CurrentGlobalFWBIsUsedCheckBox.Value
Me.GlobalFWBIsRemovedAfterCopyingCheckBox.Enabled = Me.CurrentGlobalFWBIsUsedCheckBox.Value

End Sub

Private Sub CurrentGlobalFWSIsUsedCheckBox_CustomChangeEvent()

Me.InputGlobalFWSNameTextBox.Enabled = Me.CurrentGlobalFWSIsUsedCheckBox.Value

End Sub

Private Sub CurrentGlobalTWBIsUsedCheckBox_CustomChangeEvent()

Me.InputGlobalTWBNameTextBox.Enabled = Me.CurrentGlobalTWBIsUsedCheckBox.Value
Me.ToGlobalWBBrowseButton.Enabled = Me.CurrentGlobalTWBIsUsedCheckBox.Value
Me.GlobalTWBIsRemovedAfterCopyingCheckBox.Enabled = Me.CurrentGlobalTWBIsUsedCheckBox.Value

End Sub

Private Sub CurrentGlobalTWSIsUsedCheckBox_CustomChangeEvent()

Me.InputGlobalTWSNameTextBox.Enabled = Me.CurrentGlobalTWSIsUsedCheckBox.Value

End Sub

Private Sub ChangeStateOfGlobalSettingCheckBoxes(ByVal State As Boolean)

Me.CurrentGlobalFWBIsUsedCheckBox.Value = State
CurrentGlobalFWBIsUsedCheckBox_CustomChangeEvent

Me.CurrentGlobalFWSIsUsedCheckBox.Value = State
CurrentGlobalFWSIsUsedCheckBox_CustomChangeEvent

Me.CurrentGlobalTWBIsUsedCheckBox.Value = State
CurrentGlobalTWBIsUsedCheckBox_CustomChangeEvent

Me.CurrentGlobalTWSIsUsedCheckBox.Value = State
CurrentGlobalTWSIsUsedCheckBox_CustomChangeEvent

End Sub

Private Sub DestroyObject()

'enable the parent form
FormsHelpers.ChangeStateOfAllControlsOnForm CopyingSettingsForm, True

'unload the current form
Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 
'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub

Private Sub CancelCopyingGlobalSettingsButton_Click()

DestroyObject

End Sub
