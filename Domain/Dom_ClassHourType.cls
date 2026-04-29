VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ClassHourType"
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
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_ClassHourType", "数値は0以上である必要があります。"
    This.Value = Value
End Sub

Public Property Get Value() As Long
    Value = This.Value
End Property

Public Function ToKey() As String
    If This.Value = Dom_ClassHourTypeEnum.Plan Then
        ToKey = jpnKeikaku
    ElseIf This.Value = Dom_ClassHourTypeEnum.Execution Then
        ToKey = jpnJisshi
    End If
End Function

Public Function Equals(ByVal Other As Dom_ClassHourType) As Boolean
    Equals = (This.Value = Other.Value)
End Function
