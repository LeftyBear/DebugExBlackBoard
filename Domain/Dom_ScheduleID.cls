VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.ValueObject")
Option Explicit

Private Type Member
    Value As String
End Type

Private This As Member

Public Sub Initialize(ByVal Value As String)
    If Value = vbNullString Then Err.Raise DomErrNullString, "Dom_ScheduleID", "IDÇÕï∂éöóÒÇ≈Ç†ÇÈïKóvÇ™Ç†ÇËÇ‹Ç∑ÅB"
    This.Value = Value
End Sub

Public Function Value() As String
    Value = This.Value
End Function

Public Function ToKey() As String
    ToKey = This.Value
End Function

Public Function Equals(ByVal Other As Dom_ScheduleID) As Boolean
    Equals = (This.Value = Other.Value)
End Function
