Attribute VB_Name = "CollectionToArrayModule"
Option Explicit
Option Private Module

Public Function CollectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Integer
    For i = 1 To c.Count
        a(i - 1) = c.item(i)
    Next
    CollectionToArray = a
End Function


Private Sub UnitTest_CollectionToArray()
  
Dim c As New Collection
Dim a As Variant

c.Add 20
c.Add 30

a = CollectionToArray(c)

Debug.Print TypeName(a)
Debug.Print a(1)

Set c = Nothing

End Sub
