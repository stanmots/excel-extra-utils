VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectedSheetsColorSettingsForm 
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   OleObjectBlob   =   "SelectedSheetsColorSettingsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectedSheetsColorSettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()

Me.Caption = SELECTED_SHEETS_COLORING_SETTINGS_FORM_TITLE
Me.CancelButton.Caption = CANCEL_BUTTON_TITLE
Me.SaveButton.Caption = SAVE_BUTTON_TITLE
Me.BaseCellLabel.Caption = BASECELL_LABEL
Me.SoughtForRangeLabel.Caption = SOUGHTFOR_RANGE_LABEL
Me.BaseRangeLabel.Caption = BASERANGE_LABEL
Me.ColorLabel.Caption = INPUT_COLOR_LABEL
Me.SetColorButton.Caption = SET_BUTTON_TITLE
Me.ClearColorButton.Caption = CLEAR_BUTTON_TITLE

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

Dim ColoringColumn As String
ColoringColumn = ColumnNumberToLetter(Range(Me.InputBaseCellTextBox.text).Column)

Dim TopLeftCellRowOffset As Long, TopLeftCellColumnOffset As Long
Dim RightBottomCellRowOffset As Long, RightBottomCellColumnOffset As Long

TopLeftCellRowOffset = Range(Me.SoughtForRangeTopLeftCellTextBox.text).Row - Range(Me.InputBaseCellTextBox.text).Row
TopLeftCellColumnOffset = Range(Me.SoughtForRangeTopLeftCellTextBox.text).Column - Range(Me.InputBaseCellTextBox.text).Column
RightBottomCellRowOffset = Range(Me.SoughtForRightBottomCellTextBox.text).Row - Range(Me.InputBaseCellTextBox.text).Row
RightBottomCellColumnOffset = Range(Me.SoughtForRightBottomCellTextBox.text).Column - Range(Me.InputBaseCellTextBox.text).Column


Dim i As Long
Dim ColoringOffsetsStr As String, SoughtForRange As String, BaseRange As String, BaseCell As String, SectionName As String

SoughtForRange = Me.SoughtForRangeTopLeftCellTextBox.text & ":" & Me.SoughtForRightBottomCellTextBox.text
BaseRange = Me.BaseRangeTopLeftCellTextBox.text & ":" & Me.BaseRangeRightBottomCellTextBox.text
BaseCell = Me.InputBaseCellTextBox.text

'for each selected worksheet
For i = 0 To ColoringSettingsForm.ColoringSettingsListBox.ListCount - 1
    If ColoringSettingsForm.ColoringSettingsListBox.Selected(i) = True Then

        SectionName = COLORING_SECTION_KEY & SECTION_DELIMITER & ColoringSettingsForm.ColoringSettingsListBox.List(i, 0)
        
        With ExcelUtilsMainWindow.MainStorage
            .StoreValue SectionName, COLORING_COLUMN_KEY, ColoringColumn
            .StoreValue SectionName, COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_ROW_OFFSET_KEY, TopLeftCellRowOffset
            .StoreValue SectionName, COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_COLUMN_OFFSET_KEY, TopLeftCellColumnOffset
            .StoreValue SectionName, COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY, RightBottomCellRowOffset
            .StoreValue SectionName, COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY, RightBottomCellColumnOffset
        
            .StoreValue SectionName, COLORING_SOUGHTFORRANGE_KEY, SoughtForRange
            .StoreValue SectionName, COLORING_BASERANGE_KEY, BaseRange
            .StoreValue SectionName, COLORING_BASECELL_KEY, BaseCell
            .StoreValue SectionName, COLORING_BASECOLOR_KEY, Me.ColorTextBox.text
        End With
        
        ColoringOffsetsStr = CommonHelpers.GetFormattedStringFromOffsets(TopLeftCellColumnOffset, TopLeftCellRowOffset, _
            RightBottomCellColumnOffset, RightBottomCellRowOffset)
         
        ColoringSettingsForm.ColoringSettingsListBox.List(i, 1) = BaseCell & ";" & SoughtForRange & ";" & ColoringOffsetsStr
        ColoringSettingsForm.ColoringSettingsListBox.List(i, 2) = BaseRange
        ColoringSettingsForm.ColoringSettingsListBox.List(i, 3) = Me.ColorTextBox.text
    
    End If
Next i

End Sub

Private Function IsInputValuesCorrect() As Boolean

IsInputValuesCorrect = True

Dim ctl As Control

For Each ctl In Me.Controls
If TypeOf ctl Is MSForms.TextBox And ctl.Name <> "ColorTextBox" Then
    If IsValidRange(ctl.Value) = False Then
        IsInputValuesCorrect = False
        Exit Function
    End If
End If
Next

If Range(SoughtForRangeTopLeftCellTextBox & ":" & SoughtForRightBottomCellTextBox).Count <> Range(BaseRangeTopLeftCellTextBox & ":" & BaseRangeRightBottomCellTextBox).Count Then
    IsInputValuesCorrect = False
    Exit Function
End If

End Function

Private Sub ClearColorButton_Click()

Me.ColorTextBox.text = ""

End Sub

Private Sub SetColorButton_Click()

Dim ColorCode As Long

If Len(Me.ColorTextBox.text) > 0 Then
ColorCode = CommonHelpers.ShowEditColorDialog(CLng(Me.ColorTextBox.text))
Else
ColorCode = CommonHelpers.ShowEditColorDialog
End If

If ColorCode <> -1 Then

Me.ColorTextBox.text = ColorCode

End If

End Sub

Private Sub DestroyObject()

FormsHelpers.ChangeStateOfAllControlsOnForm ColoringSettingsForm, True
Unload Me

End Sub

Private Sub CancelButton_Click()

DestroyObject

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub
