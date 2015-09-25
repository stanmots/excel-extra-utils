VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'use this class with primitive types only
Private m_First As Variant
Private m_Second As Variant

Public Property Get First() As Variant

First = m_First
    
End Property

Public Property Get Second() As Variant

Second = m_Second
    
End Property

Public Property Let First(ByVal p_First As Variant)

m_First = p_First

End Property

Public Property Let Second(ByVal p_Second As Variant)

m_Second = p_Second

End Property
