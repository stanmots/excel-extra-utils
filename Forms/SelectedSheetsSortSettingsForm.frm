VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectedSheetsSortSettingsForm 
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   OleObjectBlob   =   "SelectedSheetsSortSettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectedSheetsSortSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()

Me.Caption = SELECTED_SHEETS_SORTING_SETTINGS_FORM_TITLE
Me.InputCellFromSortColumnLabel.Caption = INPUT_SORTING_COLUMN_LABEL
Me.InputRangeOfSortEntryLabel.Caption = INPUT_SORTING_OFFSETS_LABEL
Me.CancelButton.Caption = CANCEL_BUTTON_TITLE
Me.SaveButton.Caption = SAVE_BUTTON_TITLE
Me.InputSerialCellLabel.Caption = INPUT_SERIAL_NUMBERS_LABEL

End Sub

Private Sub SaveButton_Click()

If IsInputValuesCorrect = False Then
    MsgBox INCORRECT_INPUT_VALUES_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
Else
    SaveCurrentValues
    DestroyObject
End If

End Sub

Private Sub SaveCurrentValues()

Dim SortingColumn As String
SortingColumn = ColumnNumberToLetter(Range(Me.InputCellFromSortColumnTextBox.text).Column)

Dim TopLeftCellRowOffset As Long, TopLeftCellColumnOffset As Long
Dim RightBottomCellRowOffset As Long, RightBottomCellColumnOffset As Long
Dim SerialCellRowOffset As Variant, SerialCellColumnOffset As Variant

'calculate the offsets from the chosen cell with a sorting number
'topleftcell and rightbottomcell are required to be entered into the appropriate textfield
TopLeftCellRowOffset = Range(Me.InputTopLeftCellTextBox.text).Row - Range(Me.InputCellFromSortColumnTextBox.text).Row
TopLeftCellColumnOffset = Range(Me.InputTopLeftCellTextBox.text).Column - Range(Me.InputCellFromSortColumnTextBox.text).Column
RightBottomCellRowOffset = Range(Me.InputRightBottomCellTextBox.text).Row - Range(Me.InputCellFromSortColumnTextBox.text).Row
RightBottomCellColumnOffset = Range(Me.InputRightBottomCellTextBox.text).Column - Range(Me.InputCellFromSortColumnTextBox.text).Column

'serial-number cell
If Len(Me.InputSerialCellTextBox.text) > 0 Then
    SerialCellRowOffset = Range(Me.InputSerialCellTextBox.text).Row - Range(Me.InputCellFromSortColumnTextBox.text).Row
    SerialCellColumnOffset = Range(Me.InputSerialCellTextBox.text).Column - Range(Me.InputCellFromSortColumnTextBox.text).Column
Else
    SerialCellRowOffset = ""
    SerialCellColumnOffset = ""
End If

Dim i As Long
Dim SortingOffsets As String, SortingRange As String, BaseCell As String, SectionName As String

SortingRange = Me.InputTopLeftCellTextBox.text & ":" & Me.InputRightBottomCellTextBox.text
BaseCell = Me.InputCellFromSortColumnTextBox.text

'for each selected worksheet
For i = 0 To SortingSettingsForm.SortingSettingsListBox.ListCount - 1
    If SortingSettingsForm.SortingSettingsListBox.Selected(i) = True Then

        SectionName = SORTING_SECTION_KEY & SECTION_DELIMITER & SortingSettingsForm.SortingSettingsListBox.List(i, 0)
                    
        SortingSettingsForm.SortingSettingsListBox.List(i, 1) = SortingColumn
        
        With ExcelUtilsMainWindow.MainStorage
            .StoreValue SectionName, SORTING_COLUMN_KEY, SortingColumn
            .StoreValue SectionName, SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY, TopLeftCellRowOffset
            .StoreValue SectionName, SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY, TopLeftCellColumnOffset
            .StoreValue SectionName, SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY, RightBottomCellRowOffset
            .StoreValue SectionName, SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY, RightBottomCellColumnOffset
        
            .StoreValue SectionName, SORTING_RANGE_KEY, SortingRange
            .StoreValue SectionName, SORTING_BASECELL_KEY, Me.InputCellFromSortColumnTextBox.text
            .StoreValue SectionName, SORTING_SERIAL_CELL_KEY, Me.InputSerialCellTextBox.text
              
            If Len(Me.InputSerialCellTextBox.text) > 0 Then
                .StoreValue SectionName, SORTING_SERIAL_CELL_ROW_OFFSET_KEY, SerialCellRowOffset
                .StoreValue SectionName, SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY, SerialCellColumnOffset
            End If
        End With
        
        SortingOffsets = CommonHelpers.GetFormattedStringFromOffsets(TopLeftCellColumnOffset, TopLeftCellRowOffset, _
            RightBottomCellColumnOffset, RightBottomCellRowOffset, _
            SerialCellColumnOffset, SerialCellRowOffset)
         
        SortingSettingsForm.SortingSettingsListBox.List(i, 2) = BaseCell & ";" & SortingRange & ";" & Me.InputSerialCellTextBox.text & SortingOffsets
    End If
Next i

End Sub

Private Function IsInputValuesCorrect() As Boolean

IsInputValuesCorrect = True

Dim ctl As Control

For Each ctl In Me.Controls
If TypeOf ctl Is MSForms.TextBox And ctl.Name <> "InputSerialCellTextBox" Then
    If IsValidRange(ctl.Value) = False Then
        IsInputValuesCorrect = False
        Exit Function
    End If
End If
Next

End Function

Private Sub DestroyObject()

FormsHelpers.ChangeStateOfAllControlsOnForm SortingSettingsForm, True
Unload Me

End Sub

Private Sub CancelButton_Click()

DestroyObject

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub
