VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_EnrollmentValue"
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
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_EnrollmentValue", "数値は0以上である必要があります。"
    This.Value = Value
End Sub

Public Property Get Value() As Long
    Value = This.Value
End Property

Public Function Add(ByVal Other As Dom_EnrollmentValue) As Dom_EnrollmentValue
    Dim Result As Dom_EnrollmentValue
    Set Result = New Dom_EnrollmentValue
    Result.Initialize This.Value + Other.Value
    Set Add = Result
End Function

Public Function Equals(ByVal Other As Dom_EnrollmentValue) As Boolean
    If Other Is Nothing Then Exit Function
    Equals = (This.Value = Other.Value)
End Function
