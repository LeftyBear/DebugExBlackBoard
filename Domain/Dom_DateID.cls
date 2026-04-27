VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DateID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Value As Date
End Type

Private This As Member

Public Sub Initialize(ByVal Value As Date)
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_DateID", "日付型は0以上である必要があります。"
    This.Value = Value
End Sub

Public Property Get Value() As Date
    Value = This.Value
End Property

Public Function ToKey() As String
    ToKey = CStr(This.Value)
End Function

Public Function Equals(ByVal Other As Dom_DateID) As Boolean
    If This.Value <> Other.Value Then Exit Function
    Equals = True
End Function
