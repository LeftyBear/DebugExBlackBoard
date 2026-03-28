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

Friend Sub InitializeInternal(ByVal Value As Long)
    This.Value = Value
End Sub

Public Function Create(ByVal Value As Long) As Dom_EnrollmentValue
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_EnrollmentValue", "数値は0以上である必要があります。"
    Dim Result As Dom_EnrollmentValue
    Set Result = New Dom_EnrollmentValue
    Result.InitializeInternal Value
    Set Create = Result
End Function

Public Property Get Value() As Long
    Value = This.Value
End Property
