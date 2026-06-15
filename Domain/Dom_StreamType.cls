VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_StreamType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Public Enum Dom_StreamTypeCode
    MainCode = 1
    SpecialCode
End Enum

Private Type Member
    Value As Dom_StreamTypeCode
End Type

Private This As Member

Public Sub Initialize(ByVal Value As Dom_StreamTypeCode)
    This.Value = Value
End Sub

Public Function Value() As Dom_StreamTypeCode
    Value = This.Value
End Function

Public Function IsMain() As Boolean
    IsMain = (This.Value = MainCode)
End Function

Public Function IsSpecial() As Boolean
    IsSpecial = (This.Value = SpecialCode)
End Function

Public Function Equals(ByVal Other As Dom_StreamType) As Boolean
    Equals = (Me.IsMain = Other.IsMain)
End Function
