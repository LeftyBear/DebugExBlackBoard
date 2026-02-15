Attribute VB_Name = "Util_Function"
'@Folder "Utility.Function"
Option Explicit

Public Function ResolveSchoolYear(ByVal TargetDate As Date) As Long
    Select Case VBA.Month(TargetDate)
    Case 4, 5, 6, 7, 8, 9, 10, 11, 12
        ResolveSchoolYear = VBA.Year(TargetDate)
    Case 1, 2, 3
        ResolveSchoolYear = VBA.Year(TargetDate) - 1
    End Select
End Function

Public Function ParseSearchKeywords(ByVal InputText As String) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim Normalized As String
    Normalized = InputText
    Normalized = VBA.Replace$(Normalized, Util_Character.WideSpace, Util_Character.HalfSpace)
    Normalized = VBA.Replace$(Normalized, VBA.vbTab, Util_Character.HalfSpace)
    Normalized = VBA.Replace$(Normalized, VBA.vbCr, Util_Character.HalfSpace)
    Normalized = VBA.Replace$(Normalized, VBA.vbLf, Util_Character.HalfSpace)
    Normalized = VBA.Trim$(Normalized)
    If VBA.Len(Normalized) = 0 Then
        Set ParseSearchKeywords = Result
        Exit Function
    End If
    Do While 0 < VBA.InStr(1, Normalized, Util_Character.WideSpace)
        Normalized = VBA.Replace$(Normalized, Util_Character.WideSpace, Util_Character.HalfSpace)
    Loop
    Dim Parts() As String
    Parts = VBA.Split(Normalized, Util_Character.HalfSpace)
    Dim Index As Long
    For Index = LBound(Parts) To UBound(Parts)
        If 0 < VBA.Len(Parts(Index)) Then Result.Add Parts(Index)
    Next
    Set ParseSearchKeywords = Result
End Function

Public Function NormalizeToLong(ByVal RawText As String) As Long
    Dim TextValue As String
    TextValue = Trim$(RawText)
    If TextValue = vbNullString Then TextValue = "0"
    TextValue = VBA.Replace(TextValue, Util_Character.Comma, vbNullString)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    If Not VBA.IsNumeric(TextValue) Then Err.Raise Util_ErrNum.NotNumeric, "Util_Function", "数値変換不可: " & RawText
    If 0 < VBA.InStr(1, TextValue, ".") Then Err.Raise Util_ErrNum.NotInteger, "Util_Function", "整数ではない値: " & RawText
    NormalizeToLong = CLng(TextValue)
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
        Err.Raise Util_ErrNum.NotBoolean, "Util_Function", "Boolean変換不可: " & RawText
    End Select
End Function

Public Function NormalizeToDate(ByVal RawText As String) As Date
    Dim TextValue As String
    TextValue = VBA.Trim$(RawText)
    If TextValue = vbNullString Then Err.Raise Util_ErrNum.EmptyDate, "Util_Function", "空値は許可されていません"
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.Replace(TextValue, "-", "/")
    If 0 < VBA.InStr(1, TextValue, "/") Then
        If UBound(VBA.Split(TextValue, "/")) <> 2 Then Err.Raise Util_ErrNum.InvalidDateFormat, "Util_Function", "日付形式不正: " & RawText
        Dim YearPart As Long
        YearPart = CLng(VBA.Split(TextValue, "/")(0))
        Dim MonthPart As Long
        MonthPart = CLng(VBA.Split(TextValue, "/")(1))
        Dim DayPart As Long
        DayPart = CLng(VBA.Split(TextValue, "/")(2))
    ElseIf VBA.Len(TextValue) = 8 And VBA.IsNumeric(TextValue) Then
        YearPart = CLng(VBA.Left$(TextValue, 4))
        MonthPart = CLng(VBA.Mid$(TextValue, 5, 2))
        DayPart = CLng(VBA.Right$(TextValue, 2))
    Else
        Err.Raise Util_ErrNum.InvalidDateFormat, "Util_Function", "日付形式不正: " & RawText
    End If
    On Error GoTo InvalidDate
    NormalizeToDate = VBA.DateSerial(YearPart, MonthPart, DayPart)
    On Error GoTo 0
    Exit Function
InvalidDate:
    Err.Raise Util_ErrNum.InvalidDateValue, "Util_Function", "日付値不正: " & RawText
End Function

