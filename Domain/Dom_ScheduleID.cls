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
    Key As String
End Type

Private This As Member

Public Sub Initialize(ByVal Key As String)
    If Key = vbNullString Then Err.Raise DomErrNullString, "Dom_ScheduleID", "IDÇÕï∂éöóÒÇ≈Ç†ÇÈïKóvÇ™Ç†ÇËÇ‹Ç∑ÅB"
    This.Key = Key
End Sub

Public Function ToKey() As String
    ToKey = This.Key
End Function

Public Function Equals(ByVal Other As Dom_ScheduleID) As Boolean
    If Other Is Nothing Then Exit Function
    If Me.ToKey <> Other.ToKey Then Exit Function
    Equals = True
End Function
