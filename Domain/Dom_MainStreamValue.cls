VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_MainStreamValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    UpperGrade      As Long
    UpperClassNo    As Long
End Type

Private This As Member

Public Sub Initialize(ByVal UpperGrade As Long, ByVal UpperClassNo As Long)
    If UpperGrade < 1 Then Err.Raise DomErrInvalidValue, "Dom_MainStreamValue", "学年は正の数である必要があります。"
    If UpperClassNo < 1 Then Err.Raise DomErrInvalidValue, "Dom_MainStreamValue", "学級は正の数である必要があります。"
    This.UpperGrade = UpperGrade
    This.UpperClassNo = UpperClassNo
End Sub

Public Property Get UpperGrade() As String
    UpperGrade = This.UpperGrade
End Property

Public Property Get UpperClassNo() As String
    UpperClassNo = This.UpperClassNo
End Property

Public Function Equals(ByVal Other As Dom_MainStreamValue) As Boolean
    If This.UpperGrade <> Other.UpperGrade Then Exit Function
    If This.UpperClassNo <> Other.UpperClassNo Then Exit Function
    Equals = True
End Function
