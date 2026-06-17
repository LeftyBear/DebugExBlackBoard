VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Value As String
End Type

Private This As Member

Public Sub Initialize(ByVal Value As String)
    If Value = VbNullString Then Err.Raise DomErrOutOfRange, "Dom_ClassName", "学級名は文字列である必要があります。"
    This.Value = Value
End Sub

Public Property Get Value() As String
    Value = This.Value
End Property

Public Function Equals(ByVal Other As Dom_ClassName) As Boolean
    Equals = (This.Value = Other.Value)
End Function
