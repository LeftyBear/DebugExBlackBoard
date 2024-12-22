VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDayDuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IgnoreModule WriteOnlyProperty, EncapsulatePublicField
'@Exposed
'@Folder("Sheet")
Option Explicit
Public FirstDate As Date
Public Property Set RecordList(ByVal RecordList As Scripting.IDictionary)
End Property
Public Property Let Order(ByRef DataBody() As Variant)
End Property
Public Property Let TargetName(ByVal Target As String)
End Property
Public Property Let CanInputDate(ByVal State As Boolean)
End Property
Public Function DataBody() As Variant()
End Function
