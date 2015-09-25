VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColoringSettingsForm 
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   OleObjectBlob   =   "ColoringSettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColoringSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()

Me.Caption = COLORING_SETTINGS_FORM_TITLE
Me.ColoringListBoxDescriptionLabel.Caption = COLORING_LIST_BOX_DESCRIPTION_LABEL
Me.ColoringSheetNameLabel.Caption = WORKSHEETNAME_LABEL
Me.ColoringOffsetsLabel.Caption = OFFSETS_LABEL
Me.ColoringBaseRangeLabel.Caption = COLORING_BASERANGE_LABEL
Me.CurrentColorLabel.Caption = COLOR_LABEL

Me.EditColoringSettingsButton.Caption = EDIT_BUTTON_TITLE
Me.SelectAllButton.Caption = SELECTALL_BUTTON_TITLE
Me.UnselectAllButton.Caption = UNSELECTALL_BUTTON_TITLE
Me.CancelColoringSettingsButton.Caption = CANCEL_BUTTON_TITLE

End Sub

Private Sub SelectAllButton_Click()

ListBoxFunctions.SetSelectionModeOfListBox Me.ColoringSettingsListBox, True

End Sub

Private Sub UnselectAllButton_Click()

ListBoxFunctions.SetSelectionModeOfListBox Me.ColoringSettingsListBox, False

End Sub

Private Sub EditColoringSettingsButton_Click()

If ListBoxFunctions.IsListBoxHasSelectedItems(Me.ColoringSettingsListBox) = True Then

    Dim BaseCell As String, BaseRange As String, SoughtForRange As String
    Dim Color As Long
    
    'all settings will be the same for all selected worksheets
    ColoringSettingsHelpers.RetrieveCurrentWorksheetSettings Me.ColoringSettingsListBox.List(0, 0), BaseCell:=BaseCell, BaseRange:=BaseRange, SoughtForRange:=SoughtForRange, Color:=Color
    
    If Len(BaseCell) > 0 And Len(BaseRange) > 0 And Len(SoughtForRange) > 0 Then
    
        SelectedSheetsColorSettingsForm.InputBaseCellTextBox.text = BaseCell
        
        Dim SplittedBaseRange() As String
        SplittedBaseRange = Split(BaseRange, DELIMITER:=":")
        
        SelectedSheetsColorSettingsForm.BaseRangeTopLeftCellTextBox.text = SplittedBaseRange(0)
        SelectedSheetsColorSettingsForm.BaseRangeRightBottomCellTextBox.text = SplittedBaseRange(1)
        
        Dim SplittedSoughtForRange() As String
        SplittedSoughtForRange = Split(SoughtForRange, DELIMITER:=":")
        
        SelectedSheetsColorSettingsForm.SoughtForRangeTopLeftCellTextBox.text = SplittedSoughtForRange(0)
        SelectedSheetsColorSettingsForm.SoughtForRightBottomCellTextBox.text = SplittedSoughtForRange(1)
    
    End If
    
    'color = 0 is default value
    SelectedSheetsColorSettingsForm.ColorTextBox.text = Color
 
    FormsHelpers.ChangeStateOfAllControlsOnForm Me, False
    SelectedSheetsColorSettingsForm.Show
Else
    MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
End If

End Sub

Public Sub RetrieveWorksheetsSettings()

Dim RowIndex As Long
Dim ColoringOffsetsString As String, BaseCell As String, BaseRange As String, SoughtForRange As String
Dim ColoringOffsets As Collection
Dim Color As Long

For RowIndex = 0 To Me.ColoringSettingsListBox.ListCount - 1
    Set ColoringOffsets = New Collection
    If ColoringSettingsHelpers.RetrieveCurrentWorksheetSettings(WorksheetName:=Me.ColoringSettingsListBox.List(RowIndex, 0), _
        ColoringOffsets:=ColoringOffsets, BaseRange:=BaseRange, BaseCell:=BaseCell, SoughtForRange:=SoughtForRange, Color:=Color) = True Then
        
        ColoringOffsetsString = CommonHelpers.GetFormattedStringFromOffsets(ColoringOffsets.item(COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), ColoringOffsets.item(COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
            ColoringOffsets.item(COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), ColoringOffsets.item(COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
       
        Me.ColoringSettingsListBox.List(RowIndex, 1) = BaseCell & ";" & SoughtForRange & ";" & ColoringOffsetsString
        Me.ColoringSettingsListBox.List(RowIndex, 2) = BaseRange
        Me.ColoringSettingsListBox.List(RowIndex, 3) = CStr(Color)
    End If
    Set ColoringOffsets = Nothing
Next RowIndex

End Sub

Private Sub DestroyObject()

'enable the controls on the parent form
FormsHelpers.ChangeStateOfAllControlsOnForm ExcelUtilsMainWindow, True

'unload child forms
If FormsHelpers.IsUserFormLoaded("SelectedSheetsColorSettingsForm") = True Then Unload SelectedSheetsColorSettingsForm

'unload current form
Unload Me

End Sub

Private Sub CancelColoringSettingsButton_Click()

DestroyObject

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub
