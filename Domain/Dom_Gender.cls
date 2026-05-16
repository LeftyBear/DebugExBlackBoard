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

Public Enum Dom_GenderTypePolicy
    Male = 1
    Female
End Enum

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

Public Function IsMale() As Boolean
    IsMale = (This.Value = Male)
End Function

Public Function IsFemale() As Boolean
    IsFemale = (This.Value = Female)
End Function

Public Function Equals(ByVal Other As Dom_Gender) As Boolean
    Equals = (This.Value = Other.Value)
End Function
