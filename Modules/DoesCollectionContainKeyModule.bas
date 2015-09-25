Attribute VB_Name = "DoesCollectionContainKeyModule"
Option Explicit
Option Private Module

Public Function DoesCollectionContainKey(ByVal col As Collection, ByVal Key As String) As Boolean

If col Is Nothing Then
    DoesCollectionContainKey = False
    Exit Function
End If

Dim var As Variant
Dim errCode As Long

Set var = Nothing

Err.Clear
On Error Resume Next

var = col.item(Key)
errCode = CLng(Err.Number)

On Error GoTo 0

Select Case errCode

'438 means that var is Object
Case 0, 438:    DoesCollectionContainKey = True
Case 5, 3265:   DoesCollectionContainKey = False
Case Else: Error errCode

End Select

End Function

Private Sub UnitTest_DoesCollectionContainKey()

Dim nc As New Collection
nc.Add "TestItem1", "TestKey"

Debug.Print DoesCollectionContainKey(nc, "TestKey")
Debug.Print DoesCollectionContainKey(nc, "aaaa")

Set nc = Nothing

End Sub
