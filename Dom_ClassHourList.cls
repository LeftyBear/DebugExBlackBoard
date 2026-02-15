VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Domain.EntityList"
Option Explicit

Private Type Member
    Items As VBA.Collection
End Type
Private This As Member

Public Sub Initialize(ByVal Items As VBA.Collection)
    Set This.Items = Items
End Sub

Public Property Get Items() As VBA.Collection
    Set Items = This.Items
End Property

Public Property Get Count() As Long
    Count = This.Items.Count
End Property
