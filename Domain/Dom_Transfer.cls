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
    Value As Boolean
End Type

Private This As Member

Public Sub Initialize(ByVal Value As Boolean)
    This.Value = Value
End Sub

Public Property Get IsTrue() As Boolean
    IsTrue = This.Value
End Property

Public Function Equals(ByVal Other As Dom_Transfer) As Boolean
    Equals = (Me.IsTrue = Other.IsTrue)
End Function
