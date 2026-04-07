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

Friend Sub InitializeInternal(ByVal Value As Date)
    This.Value = Value
End Sub

Public Function Create(ByVal Value As Date) As Dom_DateID
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_DateID", "日付型は0以上である必要があります。"
    Dim Result As Dom_DateID
    Set Result = New Dom_DateID
    Result.InitializeInternal Value
    Set Create = Result
End Function

Public Function Equals(ByVal Other As Dom_DateID) As Boolean
    If This.Value <> Other.Value Then Exit Function
    Equals = True
End Function

Public Function ToKey() As String
    ToKey = CStr(This.Value)
End Function

Public Property Get Value() As Date
    Value = This.Value
End Property
