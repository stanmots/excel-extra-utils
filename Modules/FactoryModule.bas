Attribute VB_Name = "FactoryModule"
Option Explicit
Option Private Module

Public Function CreateObjectOfTypeVbaSettings(ByVal WorksheetName As String) As VbaSettings

    Set CreateObjectOfTypeVbaSettings = New VbaSettings
    
    CreateObjectOfTypeVbaSettings.InitiateProperties WorksheetName

End Function

