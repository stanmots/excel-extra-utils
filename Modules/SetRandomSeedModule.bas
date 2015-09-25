Attribute VB_Name = "SetRandomSeedModule"
Option Explicit
Option Private Module

Public Sub SetRandomSeed()

Static firstCallFlag As Boolean
If firstCallFlag = False Then
    'Initialize the random-number generator
    Randomize
    firstCallFlag = True
End If

End Sub
