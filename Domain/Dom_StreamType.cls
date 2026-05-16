VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_StreamType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Public Enum Dom_StreamTypePolicy
    Main = 1
    Special
End Enum

Private Type Member
    Value As Long
End Type

Private This As Member

Public Sub Initialize(ByVal Value As Long)
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_StreamType", "学級型は0以上である必要があります。"
    This.Value = Value
End Sub

Public Function IsMain() As Boolean
    IsMain = (This.Value = Main)
End Function

Public Function IsSpecial() As Boolean
    IsSpecial = (This.Value = Special)
End Function

Public Function Equals(ByVal Other As Dom_StreamType) As Boolean
    Equals = (Me.IsMain = Other.IsMain)
End Function
