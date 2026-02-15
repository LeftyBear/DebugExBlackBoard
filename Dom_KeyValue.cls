VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_KeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.VO"
Option Explicit

Private Type Member
    Name  As String
    Value As Variant
End Type

Private This As Member

Public Sub Initialize(ByVal KeyName As String, ByVal KeyValue As Variant)
    This.Name = KeyName
    This.Value = KeyValue
End Sub

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Get Value() As Variant
    Value = This.Value
End Property
