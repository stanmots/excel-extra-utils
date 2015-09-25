Attribute VB_Name = "RandomizedQuickSortModule"
Option Explicit
Option Private Module

Public Sub RandomizedQuickSort(ByRef Arr As Variant, ByVal leftmost As Long, ByVal rightmost As Long)

Dim pivot As Variant
Dim tmpVar As Variant
Dim tmpLo As Long
Dim tmpHi As Long
  
tmpLo = leftmost
tmpHi = rightmost

' Generate random pivot
Call SetRandomSeed
Dim randomPivotIndex As Long
randomPivotIndex = CLng(Int((rightmost - leftmost + 1) * Rnd + leftmost))

pivot = Arr(randomPivotIndex)

Do While tmpLo <= tmpHi

    'Starting at the bottom of the list move up until we find the value that is bigger than pivot
    Do While Arr(tmpLo) < pivot And tmpLo < rightmost
      tmpLo = tmpLo + 1
    Loop

    'Starting at the top of the list move down until we find the first value that is smaller than pivot
    Do While pivot < Arr(tmpHi) And tmpHi > leftmost
      tmpHi = tmpHi - 1
    Loop
    
    If tmpLo < tmpHi Then
      SwapItemsInArray Arr, tmpLo, tmpHi
    End If
    
    If tmpLo <= tmpHi Then
      tmpLo = tmpLo + 1
      tmpHi = tmpHi - 1
    End If
Loop

  If leftmost < tmpHi Then RandomizedQuickSort Arr, leftmost, tmpHi
  If tmpLo < rightmost Then RandomizedQuickSort Arr, tmpLo, rightmost

End Sub

Private Sub UnitTest_RandomizedQuickSort()

'Initialize an array with the test values
Dim TestArray() As Variant
TestArray = Array(23, 12, 123, 16, 57, 48, 78, 3, 149, 963, 28, 125, 37, 77, 77)

Debug.Print "Content of the array before sorting: "
PrintArrayToConsole TestArray

Debug.Print "Content of the array after sorting: "
RandomizedQuickSort TestArray, LBound(TestArray), UBound(TestArray)
PrintArrayToConsole TestArray

End Sub


