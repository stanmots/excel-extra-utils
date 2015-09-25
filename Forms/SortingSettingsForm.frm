VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SortingSettingsForm 
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   OleObjectBlob   =   "SortingSettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SortingSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()

Me.Caption = SORTING_SETTINGS_FORM_TITLE
Me.SortingListBoxDescriptionLabel = SORTING_LIST_BOX_DESCRIPTION_LABEL
Me.SortingSheetNameLabel = SORTING_WORKSHEETS_NAME_LABEL
Me.SortingColumnLabel = SORTING_COLUMN_LABEL
Me.SortingOffsetsLabel = SORTING_OFFSETS_LABEL
Me.EditSortingSettingsButton.Caption = EDIT_BUTTON_TITLE
Me.SelectAllButton.Caption = SELECTALL_BUTTON_TITLE
Me.UnselectAllButton.Caption = UNSELECTALL_BUTTON_TITLE
Me.CancelSortingSettingsButton.Caption = CANCEL_BUTTON_TITLE

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub

Private Sub CancelSortingSettingsButton_Click()

DestroyObject

End Sub

Private Sub EditSortingSettingsButton_Click()

If ListBoxFunctions.IsListBoxHasSelectedItems(Me.SortingSettingsListBox) = True Then

    Dim BaseCell As String, SortingRange As String, SerialCell As String
    SortingSettingsHelpers.RetrieveCurrentWorksheetSettings Me.SortingSettingsListBox.List(0, 0), BaseCell:=BaseCell, SortingRange:=SortingRange, SerialCell:=SerialCell
    
    If Len(BaseCell) > 0 And Len(SortingRange) > 0 Then
    
        SelectedSheetsSortSettingsForm.InputCellFromSortColumnTextBox.text = BaseCell
        
        Dim SplittedRange() As String
        SplittedRange = Split(SortingRange, DELIMITER:=":")
        
        SelectedSheetsSortSettingsForm.InputTopLeftCellTextBox.text = SplittedRange(0)
        SelectedSheetsSortSettingsForm.InputRightBottomCellTextBox.text = SplittedRange(1)
    
    End If
    
    If Len(SerialCell) > 0 Then
        SelectedSheetsSortSettingsForm.InputSerialCellTextBox.text = SerialCell
    End If
     
    FormsHelpers.ChangeStateOfAllControlsOnForm Me, False
    SelectedSheetsSortSettingsForm.Show
Else
    MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
End If

End Sub

Private Sub SelectAllButton_Click()

ListBoxFunctions.SetSelectionModeOfListBox Me.SortingSettingsListBox, True

End Sub

Private Sub UnselectAllButton_Click()

ListBoxFunctions.SetSelectionModeOfListBox Me.SortingSettingsListBox, False

End Sub

Public Sub RetrieveWorksheetsSettings()

Dim RowIndex As Long
Dim SortingColumn As String, SortingOffsetsString As String, BaseCell As String, SortingRange As String, SerialCell As String
Dim SortingOffsets As Collection

For RowIndex = 0 To Me.SortingSettingsListBox.ListCount - 1
    Set SortingOffsets = New Collection
    If SortingSettingsHelpers.RetrieveCurrentWorksheetSettings(Me.SortingSettingsListBox.List(RowIndex, 0), _
        SortingColumn, SortingOffsets, SortingRange, BaseCell, SerialCell) = True Then
        
        Me.SortingSettingsListBox.List(RowIndex, 1) = SortingColumn

        SortingOffsetsString = CommonHelpers.GetFormattedStringFromOffsets(SortingOffsets.item(SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
            SortingOffsets.item(SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY), _
            SortingOffsets.item(SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_SERIAL_CELL_ROW_OFFSET_KEY))
       
        Me.SortingSettingsListBox.List(RowIndex, 2) = BaseCell & ";" & SortingRange & ";" & SerialCell & SortingOffsetsString
        
    End If
    Set SortingOffsets = Nothing
Next RowIndex

End Sub

Private Sub DestroyObject()

'enable the controls on the parent form
FormsHelpers.ChangeStateOfAllControlsOnForm ExcelUtilsMainWindow, True

'unload child forms
If FormsHelpers.IsUserFormLoaded("SelectedSheetsSortSettingsForm") = True Then Unload SelectedSheetsSortSettingsForm

'unload current form
Unload Me

End Sub
