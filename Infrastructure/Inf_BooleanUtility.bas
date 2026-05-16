Attribute VB_Name = "Inf_BooleanUtility"
'@Folder("Application.Service")
Option Explicit
Option Private Module

Public Function ToBoolean(ByVal RawText As String) As Boolean
    Dim TextValue As String
    TextValue = VBA.Trim$(RawText)
    If TextValue = vbNullString Then TextValue = "FALSE"
    '@Ignore AssignmentNotUsed
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    Select Case VBA.UCase$(TextValue)
    Case "TRUE"
        ToBoolean = True
    Case "FALSE"
        ToBoolean = False
    Case "1"
        ToBoolean = True
    Case "0"
        ToBoolean = False
    Case Else
        Err.Raise DomErrNotBoolean, "Inf_BooleanUtility", "Booleanに変換できる値である必要があります。値: " & RawText
    End Select
End Function
