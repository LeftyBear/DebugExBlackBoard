Attribute VB_Name = "Dom_Util_Boolean"
'@Folder("Domain.Utility")
Option Explicit
Option Private Module

Public Function NormalizeToBoolean(ByVal RawText As String) As Boolean
    Dim TextValue As String
    TextValue = VBA.Trim$(RawText)
    If TextValue = vbNullString Then TextValue = "FALSE"
    '@Ignore AssignmentNotUsed
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    Select Case VBA.UCase$(TextValue)
    Case "TRUE"
        NormalizeToBoolean = True
    Case "FALSE"
        NormalizeToBoolean = False
    Case "1"
        NormalizeToBoolean = True
    Case "0"
        NormalizeToBoolean = False
    Case Else
        Err.Raise DomErrNotBoolean, "Dom_Util_Boolesn", "Booleanに変換できる値である必要があります。値: " & RawText
    End Select
End Function
