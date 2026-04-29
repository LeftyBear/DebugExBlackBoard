VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_SchoolYear"
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
    If Not App_DateUtility.IsFourDigitNumber(Value) Then Err.Raise DomErrNotFourDigitNumber, "Dom_SchoolYear", "年度は数字４桁である必要があります。"
    This.Value = Value
End Sub

Public Property Get Value() As Long
    Value = This.Value
End Property

Public Function ToKey() As String
    ToKey = CStr(This.Value)
End Function

Public Function Equals(ByVal Other As Dom_SchoolYear) As Boolean
    Equals = (This.Value = Other.Value)
End Function
