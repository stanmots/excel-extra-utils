Attribute VB_Name = "PrintArrayToConsoleModule"
Option Explicit
Option Private Module

'This function also works thiw collections
Public Sub PrintArrayToConsole(ByVal Arr As Variant)

Debug.Print "[";
Dim item As Variant
For Each item In Arr
    Debug.Print item;
Next
Debug.Print "]";
Debug.Print vbCrLf

End Sub
