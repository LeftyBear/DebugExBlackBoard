VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourValue"
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
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_ClassHourValue", "数値は0以上である必要があります。"
    This.Value = Value
End Sub

Public Property Get Value() As Long
    Value = This.Value
End Property

Public Function Add(ByVal Other As Dom_ClassHourValue) As Dom_ClassHourValue
    Dim Result As Dom_ClassHourValue
    Set Result = New Dom_ClassHourValue
    Result.Initialize This.Value + Other.Value
    Set Add = Result
End Function

Public Function Equals(ByVal Other As Dom_ClassHourValue) As Boolean
    Equals = (This.Value = Other.Value)
End Function
