VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    DateID  As Dom_DateID
    Value   As Boolean
End Type

Private This As Member

Public Sub Initialize(ByVal DateID As Dom_DateID, ByVal Value As Boolean)
    Set This.DateID = DateID
    This.Value = Value
End Sub

Friend Sub ChangeValueInternal(ByVal NewValue As Boolean)
    This.Value = NewValue
End Sub

Public Property Get DateID() As Dom_DateID
    DateID = This.DateID
End Property

Public Property Get Value() As Boolean
    Value = This.Value
End Property
