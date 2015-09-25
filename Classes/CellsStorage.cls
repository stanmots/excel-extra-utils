VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellsStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CellsValues As Collection
Private m_CellsAddresses As Collection

Public Property Get CellsValues() As Collection

Set CellsValues = m_CellsValues
    
End Property

Public Property Get CellsAddresses() As Collection

Set CellsAddresses = m_CellsAddresses
    
End Property

Public Property Set CellsValues(ByVal ncv As Collection)

Set m_CellsValues = ncv
    
End Property

Public Property Set CellsAddresses(ByVal nca As Collection)

Set m_CellsAddresses = nca
    
End Property

Private Sub Class_Initialize()

Set m_CellsValues = New Collection
Set m_CellsAddresses = New Collection

End Sub

Private Sub Class_Terminate()

Set m_CellsValues = Nothing
Set m_CellsAddresses = Nothing

End Sub
