Attribute VB_Name = "Dom_Util_Long"
'@Folder("Domain.Utility")
Option Explicit
Option Private Module

Public Function NormalizeToLong(ByVal RawText As String) As Long
    Dim TextValue As String
    TextValue = Trim$(RawText)
    If TextValue = vbNullString Then TextValue = "0"
    TextValue = VBA.Replace(TextValue, charComma, vbNullString)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    If Not VBA.IsNumeric(TextValue) Then Err.Raise DomErrNotNumeric, "Dom_Util_Long", "数値に変換できる値である必要があります。値: " & RawText
    If 0 < VBA.InStr(1, TextValue, charPeriod) Then Err.Raise DomErrNotInteger, "Dom_Util_Long", "値は整数である必要があります。値: " & RawText
    NormalizeToLong = CLng(TextValue)
End Function
