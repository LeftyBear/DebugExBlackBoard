Attribute VB_Name = "App_NumericUtility"
'@Folder("Application.Utility")
Option Explicit
Option Private Module

Public Function ExtractNumber(ByVal RawText As String) As Long
    Dim TextValue As String
    TextValue = VBA.Trim$(RawText)
    If TextValue = vbNullString Then TextValue = "0"
    TextValue = VBA.Replace(TextValue, charComma, vbNullString)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    If Not VBA.IsNumeric(TextValue) Then Err.Raise DomErrNotNumeric, "App_NumericUtility", "数値に変換できる値である必要があります。値: " & RawText
    If 0 < VBA.InStr(1, TextValue, charPeriod) Then Err.Raise DomErrNotInteger, "App_NumericUtility", "値は整数である必要があります。値: " & RawText
    ExtractNumber = CLng(TextValue)
End Function
