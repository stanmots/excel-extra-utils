Attribute VB_Name = "SwapItemsInCollectionModule"
Option Explicit
Option Private Module

Public Function SwapItemsInCollection(ByVal c As Collection, ByVal FirstIndex As Long, ByVal SecondIndex As Long)

If FirstIndex = SecondIndex Or c.Count < 2 Then
    Debug.Print "Error! Provided items in [SwapItemsInCollection]cannot be swapped!"
    Exit Function
End If

Dim i As Long
Dim j As Long

If FirstIndex < SecondIndex Then
    i = FirstIndex
    j = SecondIndex
Else
    i = SecondIndex
    j = FirstIndex
End If


c.Add c.item(j), Before:=i
c.Add c.item(i + 1), Before:=j + 1

c.Remove (i + 1)
c.Remove (j + 1)

End Function
