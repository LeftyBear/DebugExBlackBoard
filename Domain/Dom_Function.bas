Attribute VB_Name = "Dom_Function"
'@Folder("Domain.Function")
Option Explicit
Option Private Module

Public Function GetDateIndex(ByVal TargetDate As Date) As Long
    Dim SchoolYear As Long
    SchoolYear = GetSchoolYear(TargetDate)
    GetDateIndex = TargetDate - CDate(VBA.DateSerial(CInt(SchoolYear), 4, 1)) + 1
End Function

Public Function GetSchoolYear(ByVal TargetDate As Date) As Long
    If VBA.Month(TargetDate) < 4 Then
        GetSchoolYear = VBA.Year(TargetDate) - 1
    Else
        GetSchoolYear = VBA.Year(TargetDate)
    End If
End Function

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
        Err.Raise DomErrNotBoolean, "Dom_Function", "Booleanに変換できる値である必要があります。値: " & RawText
    End Select
End Function

Public Function NormalizeToDate(ByVal RawText As String) As Date
    Dim TextValue As String
    TextValue = VBA.Trim$(RawText)
    If TextValue = vbNullString Then Err.Raise DomErrEmptyDate, "Dom_Function", "日付に変換する値が必要です。"
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.Replace(TextValue, charHyphen, charBackSlash)
    If 0 < VBA.InStr(1, TextValue, charSlash) Then
        If UBound(VBA.Split(TextValue, charSlash)) <> 2 Then Err.Raise DomErrInvalidDateFormat, "Dom_Function", "日付に変換できる形式である必要があります。値: " & RawText
        Dim YearPart As Long
        YearPart = CLng(VBA.Split(TextValue, charSlash)(0))
        Dim MonthPart As Long
        MonthPart = CLng(VBA.Split(TextValue, charSlash)(1))
        Dim DayPart As Long
        DayPart = CLng(VBA.Split(TextValue, charSlash)(2))
    ElseIf VBA.Len(TextValue) = 8 And VBA.IsNumeric(TextValue) Then
        YearPart = CLng(VBA.Left$(TextValue, 4))
        MonthPart = CLng(VBA.Mid$(TextValue, 5, 2))
        DayPart = CLng(VBA.Right$(TextValue, 2))
    Else
        Err.Raise DomErrInvalidDateFormat, "Dom_Function", "日付に変換できる形式である必要があります。値: " & RawText
    End If
    On Error GoTo InvalidDate
    NormalizeToDate = CDate(VBA.DateSerial(YearPart, MonthPart, DayPart))
    On Error GoTo 0
    Exit Function
InvalidDate:
    Err.Raise DomErrInvalidDateValue, "Dom_Function", "日付に変換できる値である必要があります。値: " & RawText
End Function

Public Function NormalizeToLong(ByVal RawText As String) As Long
    Dim TextValue As String
    TextValue = Trim$(RawText)
    If TextValue = vbNullString Then TextValue = "0"
    TextValue = VBA.Replace(TextValue, charComma, vbNullString)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    If Not VBA.IsNumeric(TextValue) Then Err.Raise DomErrNotNumeric, "Dom_Function", "数値に変換できる値である必要があります。値: " & RawText
    If 0 < VBA.InStr(1, TextValue, ".") Then Err.Raise DomErrNotInteger, "Dom_Function", "値は整数である必要があります。値: " & RawText
    NormalizeToLong = CLng(TextValue)
End Function
