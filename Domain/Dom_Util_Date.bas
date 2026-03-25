Attribute VB_Name = "Dom_Util_Date"
'@Folder("Domain.Utility")
Option Explicit
Option Private Module

Public Function DiffApr1st(ByVal TargetDate As Date) As Long
    DiffApr1st = TargetDate - GetApr1st(TargetDate) + 1
End Function

Public Function GetApr1st(ByVal TargetDate As Date) As Date
    Dim SchoolYear As Long
    SchoolYear = GetSchoolYear(TargetDate)
    GetApr1st = VBA.DateSerial(CInt(SchoolYear), 4, 1)
End Function

Public Function GetBeginOfMonth(ByVal TargetDate As Date) As Date
    GetBeginOfMonth = VBA.DateSerial(VBA.Year(TargetDate), VBA.Month(TargetDate), 1)
End Function

Public Function GetEndOfMonth(ByVal TargetDate As Date) As Date
    GetEndOfMonth = VBA.DateSerial(VBA.Year(TargetDate), VBA.Month(TargetDate) + 1, 0)
End Function

Public Function GetMar31th(ByVal TargetDate As Date) As Date
    Dim SchoolYear As Long
    SchoolYear = GetSchoolYear(TargetDate)
    If 3 < VBA.Month(TargetDate) Then
        Dim TargetYear As Long
        TargetYear = SchoolYear + 1
    Else
        TargetYear = SchoolYear
    End If
    GetMar31th = VBA.DateSerial(CInt(TargetYear), 3, 31)
End Function

Public Function GetMonday(ByVal TargetDate As Date) As Date
    Dim WeekDayIndex As Long
    WeekDayIndex = VBA.Weekday(TargetDate, vbMonday)
    Dim Buffer As Date
    Buffer = TargetDate - (WeekDayIndex - 1)
    If OutOfSchoolYear(Buffer) Then
        GetMonday = GetApr1st(TargetDate)
    Else
        GetMonday = Buffer
    End If
End Function

Public Function GetSchoolYear(ByVal TargetDate As Date) As Long
    If VBA.Month(TargetDate) < 4 Then
        GetSchoolYear = VBA.Year(TargetDate) - 1
    Else
        GetSchoolYear = VBA.Year(TargetDate)
    End If
End Function

Public Function GetSunday(ByVal TargetDate As Date) As Date
    Dim WeekDayIndex As Long
    WeekDayIndex = VBA.Weekday(TargetDate, vbMonday)
    Dim Buffer As Date
    Buffer = TargetDate + (7 - WeekDayIndex)
    If OutOfSchoolYear(Buffer) Then
        GetSunday = GetMar31th(TargetDate)
    Else
        GetSunday = Buffer
    End If
End Function

Public Function GetSchoolWeek(ByVal TargetDate As Date) As Long
    Dim Apr1st As Date
    Apr1st = GetApr1st(TargetDate)
    Dim BaseMonday As Date
    BaseMonday = Apr1st - (VBA.Weekday(Apr1st, vbMonday) - 1)
    Dim TargetMonday As Date
    TargetMonday = TargetDate - (VBA.Weekday(TargetDate, vbMonday) - 1)
    GetSchoolWeek = ((TargetMonday - BaseMonday) / 7) + 1
End Function

Public Function IsFourDigitNumber(ByVal Value As String) As Boolean
    If VBA.Len(Value) <> 4 Then Exit Function
    If Not VBA.IsNumeric(Value) Then Exit Function
    IsFourDigitNumber = True
End Function

Public Function NormalizeToDate(ByVal RawText As String) As Date
    Dim TextValue As String
    TextValue = VBA.Trim$(RawText)
    If TextValue = vbNullString Then Err.Raise DomErrEmptyDate, "Dom_Util_Date", "日付に変換する値が必要です。"
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.Replace(TextValue, charHyphen, charBackSlash)
    If 0 < VBA.InStr(1, TextValue, charSlash) Then
        If UBound(VBA.Split(TextValue, charSlash)) <> 2 Then Err.Raise DomErrInvalidDateFormat, "Dom_Util_Date", "日付に変換できる形式である必要があります。値: " & RawText
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
        Err.Raise DomErrInvalidDateFormat, "Dom_Util_Date", "日付に変換できる形式である必要があります。値: " & RawText
    End If
    On Error GoTo InvalidDate
    NormalizeToDate = VBA.DateSerial(YearPart, MonthPart, DayPart)
    On Error GoTo 0
    Exit Function
InvalidDate:
    Err.Raise DomErrInvalidDateValue, "Dom_Util_Date", "日付に変換できる値である必要があります。値: " & RawText
End Function

Public Function OutOfSchoolYear(ByVal TargetDate As Date) As Boolean
    Dim Apr1st As Date
    Apr1st = GetApr1st(TargetDate)
    Dim Mar31th As Date
    Mar31th = GetMar31th(TargetDate)
    If TargetDate < Apr1st Then OutOfSchoolYear = True
    If Mar31th < TargetDate Then OutOfSchoolYear = True
End Function
