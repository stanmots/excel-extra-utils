Attribute VB_Name = "SwapItemsInArrayModule"
Option Explicit
Option Private Module

Public Sub SwapItemsInArray(ByRef Arr As Variant, _
  FirstIndex As Long, SecondIndex As Long)

If IsArray(Arr) = True Then
    If FirstIndex < LBound(Arr) Or SecondIndex > UBound(Arr) Then
        Debug.Print "There was an error during swaping items!"
        Exit Sub
    End If
Else
    Debug.Print "SwapItemsInArray function needs an array as a parameter!"
    Exit Sub
End If

If FirstIndex = SecondIndex Then
    Debug.Print "Warning! Items in [SwapItemsInArray] are identical. There won't be any swapping."
    Exit Sub
End If

Dim tmpVar As Variant

tmpVar = Arr(FirstIndex)
Arr(FirstIndex) = Arr(SecondIndex)
Arr(SecondIndex) = tmpVar

End Sub

