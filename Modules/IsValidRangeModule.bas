Attribute VB_Name = "IsValidRangeModule"
Option Explicit
Option Private Module

Public Function IsValidRange(ByVal str As String) As Boolean

IsValidRange = False

Dim rg As Range
Dim errCode As Long

Err.Clear
On Error Resume Next

Set rg = Range(str)
errCode = CLng(Err.Number)

On Error GoTo 0

If errCode = 0 Then IsValidRange = True

End Function
