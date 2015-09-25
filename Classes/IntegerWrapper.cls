VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IntegerWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Value As Integer

Public Event ValueChanged(ByVal Value As Integer)

Public Property Get Value() As Integer

Value = m_Value
    
End Property


Public Property Let Value(ByVal nValue As Integer)

m_Value = nValue
RaiseEvent ValueChanged(m_Value)
    
End Property
