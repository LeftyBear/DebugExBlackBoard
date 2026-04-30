VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_Gender"
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
    If Value < 0 Then Err.Raise DomErrNegativeNumber, "Dom_Gender", "«•ÊŒ^‚Í0ˆÈã‚Å‚ ‚é•K—v‚ª‚ ‚è‚Ü‚·B"
    This.Value = Value
End Sub

Public Property Get Value() As Long
    Value = This.Value
End Property

Public Function ToKey() As String
    If This.Value = Dom_GenderTypePolicy.Male Then
        ToKey = jpnOtoko
    ElseIf This.Value = Dom_GenderTypePolicy.Female Then
        ToKey = jpnOnna
    End If
End Function

Public Function Equals(ByVal Other As Dom_Gender) As Boolean
    Equals = (This.Value = Other.Value)
End Function
