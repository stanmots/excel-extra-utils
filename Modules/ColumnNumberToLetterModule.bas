Attribute VB_Name = "ColumnNumberToLetterModule"
Option Explicit
Option Private Module

Public Function ColumnNumberToLetter(ByVal ColumnNumber As Long) As String

Dim cLetter As Variant
cLetter = Split(Cells(1, ColumnNumber).Address(True, False), "$")
ColumnNumberToLetter = cLetter(0)

End Function

Private Sub UnitTest_ColumnNumberToLetter()
  
Debug.Print ColumnNumberToLetter(3)

End Sub
