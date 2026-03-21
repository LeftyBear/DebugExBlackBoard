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

Friend Sub InitializeInternal(ByVal Value As Long)
    This.Value = Value
End Sub

Public Function Create(ByVal Value As Long) As Dom_Grade
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_Grade", "学年は0以上である必要があります。"
    Dim Result As Dom_Grade
    Set Result = New Dom_Grade
    Result.InitializeInternal Value
    Set Create = Result
End Function

Public Function Equals(ByVal Other As Dom_Grade) As Boolean
    If This.Value <> Other.Value Then Exit Function
    Equals = True
End Function

Public Function ToKey() As String
    ToKey = CStr(This.Value) & jpnNen
End Function

Public Property Get Value() As Long
    Value = This.Value
End Property
