VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Grade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Value As Long
End Type

Private This As Member

Public Sub Initialize(ByVal Value As Long)
    If Not IsTarget(Value) Then Err.Raise DomErrOutOfRange, "Dom_Grade", "学年は3以上9以下である必要があります。"
    This.Value = Value
End Sub

Public Property Get Value() As Long
    Value = This.Value
End Property

Public Function Equals(ByVal Other As Dom_Grade) As Boolean
    Equals = (This.Value = Other.Value)
End Function

Private Function IsTarget(ByVal Value As Long) As Boolean
    If Value < 3 Then Exit Function
    If 9 < Value Then Exit Function
    IsTarget = True
End Function
