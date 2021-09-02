VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_Label As Collection
'private member

' Get/Set Properties for Labeling
Property Get Labels() As Collection
    Set Labels = m_Label
End Property

Property Set Labels(newLabels As Collection)
    Set m_Label = newLabels
End Property
