VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Combo1 As Collection
'private member

' Get/Set Properties for Combo 1
Property Get Combo1() As Collection
    Set Combo1 = m_Combo1
End Property

Property Set Combo1(newCombo As Collection)
    Set m_Combo1 = newCombo
End Property

