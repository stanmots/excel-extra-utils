Attribute VB_Name = "GetArrayDimensionsNumberModule"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetArrayDimensionsNumber
' This function returns the number of dimensions of an array.
' An unallocated dynamic array has 0 dimensions.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Option Private Module

Public Function GetArrayDimensionsNumber(var As Variant) As Integer
On Error GoTo Err:
    Dim i As Integer
    Dim tmp As Integer
    i = 0
    Do While True:
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    GetArrayDimensionsNumber = i - 1
End Function

Private Sub UnitTest_GetArrayDimensionsNumber()
  
Dim OneDimArray As Variant
OneDimArray = Array(23, 12, 123)

Dim TwoDimArray(1 To 1, 1 To 3) As Integer
TwoDimArray(1, 1) = 10
TwoDimArray(1, 2) = 11
TwoDimArray(1, 3) = 12

Dim UnallocArray() As Integer

Debug.Print GetArrayDimensionsNumber(OneDimArray)
Debug.Print GetArrayDimensionsNumber(TwoDimArray)
Debug.Print GetArrayDimensionsNumber(UnallocArray)

End Sub
