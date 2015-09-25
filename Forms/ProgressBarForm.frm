VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBarForm 
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   -15405
   ClientWidth     =   7725
   OleObjectBlob   =   "ProgressBarForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Const DetailsButtonOpenStateMark As String = ">>"
Private Const DetailsButtonClosedStateMark As String = "<<"

Private DetailsFrameState As DetailsFrameStateType

'byte because we need store only percent values (0-100)
Private CurrentLoopPercent As Byte
Private IgnoredLoopsCount As Long
Private WithEvents CurrentPercentValue As IntegerWrapper
Attribute CurrentPercentValue.VB_VarHelpID = -1

Private Enum DetailsFrameStateType

DETAILS_FRAME_OPEN = 0
DETAILS_FRAME_CLOSED = 1

End Enum

Private Sub CurrentPercentValue_ValueChanged(ByVal Value As Integer)

If Value = 100 Then
    Me.CloseProgressBarButton.Enabled = True
End If

End Sub

Public Property Get GetCurrentPercentValue() As Byte

GetCurrentPercentValue = CurrentPercentValue.Value
    
End Property
Public Sub SetMainLabelText(ByVal text As String)

Me.MainProcessLabel.Caption = text
Me.Repaint

End Sub

Public Sub SetCurrentOperationLabelText(ByVal text As String)

Me.CurrentOperationLabel.Caption = text
Me.Repaint

End Sub

Public Sub AddMessageToDetailsBox(ByVal msg As String)

Me.DetailsTextBox.text = Me.DetailsTextBox.text & "[" & Format(Now(), "hh:mm:ss") & "] - " & msg & vbCrLf

Me.Repaint

End Sub

Public Sub SetLoopsParameters(ByVal AddedOverallPercentValue As Byte, ByVal LoopsNumber As Long)

CurrentLoopPercent = CByte(Int(AddedOverallPercentValue / LoopsNumber))

'check if LoopsNumber > AddedOverallPercentValue
If CurrentLoopPercent = 0 Then
    CurrentLoopPercent = 1
End If

'correct the round error
Dim PercentWithRoundError As Long
PercentWithRoundError = CurrentLoopPercent * LoopsNumber
If PercentWithRoundError < AddedOverallPercentValue Then
    IncreaseProgressByPercent AddedOverallPercentValue - PercentWithRoundError
ElseIf PercentWithRoundError > AddedOverallPercentValue Then
    IgnoredLoopsCount = PercentWithRoundError - AddedOverallPercentValue
End If

End Sub

Public Sub ClearLoopsParameters()

CurrentLoopPercent = 0

End Sub

Public Sub IncreaseProgressInsideLoop()

If IgnoredLoopsCount = 0 Then
    IncreaseProgressByPercent CurrentLoopPercent
Else
    IgnoredLoopsCount = IgnoredLoopsCount - 1
End If

End Sub

Public Sub FinishProgress()

CurrentPercentValue.Value = 100
Me.CurrentProgressTextLabel.Caption = CStr(CurrentPercentValue.Value) & " %"
Me.IndicatorLabel.Width = Me.PlaceholderLabel.Width

Me.Repaint

End Sub

Public Sub ResetProgress()

CurrentLoopPercent = 0
IgnoredLoopsCount = 0
CurrentPercentValue.Value = 0
Me.CurrentProgressTextLabel.Caption = CStr(CurrentPercentValue.Value) & " %"
Me.IndicatorLabel.Width = 0
Me.MainProcessLabel.Caption = ""
Me.CurrentOperationLabel.Caption = ""

Me.Repaint

End Sub

Public Sub IncreaseProgressByPercent(ByVal AddedPercentValue As Byte)

If AddedPercentValue <= 0 Or AddedPercentValue > 100 Then
    Debug.Print "Error! Cannot increase the current progress, because the new percent's value is incorrect (" & CStr(AddedPercentValue) & "). It must be between next interval: [0,100]."
    Exit Sub
End If

If CurrentPercentValue.Value + AddedPercentValue > 100 Then
    'Debug.Print "Warning! Cannot increase the current progress, because the new percent's value (" & CStr(AddedPercentValue) & ") is bigger than expected."
    CurrentPercentValue.Value = 100
    Me.IndicatorLabel.Width = Me.PlaceholderLabel.Width
    Me.Repaint
    Exit Sub
End If

CurrentPercentValue.Value = CurrentPercentValue.Value + AddedPercentValue
Me.CurrentProgressTextLabel.Caption = CStr(CurrentPercentValue.Value) & " %"

Dim AddedIndicatorWidth As Double
AddedIndicatorWidth = Me.PlaceholderLabel.Width * AddedPercentValue / 100

If Me.IndicatorLabel.Width + AddedIndicatorWidth > Me.PlaceholderLabel.Width Then
    Me.IndicatorLabel.Width = Me.PlaceholderLabel.Width
    Me.Repaint
    Exit Sub
End If

Me.IndicatorLabel.Width = Me.IndicatorLabel.Width + AddedIndicatorWidth

Me.Repaint

End Sub

Private Sub UserForm_Initialize()

CurrentLoopPercent = 0
IgnoredLoopsCount = 0
Set CurrentPercentValue = New IntegerWrapper
CurrentPercentValue.Value = 0

Me.Caption = PROGRESS_BAR_TITLE
Me.CloseProgressBarButton.Caption = CLOSE_PROGRESS_BAR_BUTTON_TITLE
Me.DetailsFrame.Caption = DETAILS_FRAME_TITLE

SetInitialDetailsFrameMode DETAILS_FRAME_OPEN

End Sub

Private Sub SetInitialDetailsFrameMode(ByVal mode As DetailsFrameStateType)

Select Case DetailsFrameState
    Case DETAILS_FRAME_OPEN
        DetailsFrameState = DETAILS_FRAME_OPEN
        Me.DetailsButton.Caption = DetailsButtonOpenStateMark & "  " & HIDE_DETAILS_BUTTON_TITLE
    Case DETAILS_FRAME_CLOSED
        CloseDetailsFrame
End Select

End Sub

Private Sub DetailsButton_Click()

Select Case DetailsFrameState
    Case DETAILS_FRAME_OPEN
        CloseDetailsFrame
    Case DETAILS_FRAME_CLOSED
        OpenDetailsFrame
    Case Else
        Debug.Print "Warning! The variable 'DetailsFrameState' has an inappropriate value (" & CStr(DetailsFrameState) & ")."
End Select

End Sub

Private Sub CloseDetailsFrame()

Me.Height = Me.Height - Me.DetailsFrame.Height
DetailsFrameState = DETAILS_FRAME_CLOSED
Me.DetailsButton.Caption = DetailsButtonClosedStateMark & "  " & SHOW_DETAILS_BUTTON_TITLE
Me.Repaint

End Sub

Private Sub OpenDetailsFrame()

Me.Height = Me.Height + Me.DetailsFrame.Height
DetailsFrameState = DETAILS_FRAME_OPEN
Me.DetailsButton.Caption = DetailsButtonOpenStateMark & "  " & HIDE_DETAILS_BUTTON_TITLE
Me.Repaint

End Sub

Private Sub CloseProgressBarButton_Click()

Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'Prevent user from closing with the Close box in the title bar
'CloseMode vbFormControlMenu(0) - The user has chosen the Close command from the Control menu on the UserForm.
'CloseMode vbFormCode(1) - the Unload statement is invoked from code
'CloseMode vbAppWindows(2) - The current Windows operating environment session is ending
'CloseMode vbAppTaskManager(3) - The Windows Task Manager is closing the application

If CloseMode < 2 Then
    If CurrentPercentValue.Value <> 100 Then
        Cancel = 1
    End If
End If
    
End Sub
