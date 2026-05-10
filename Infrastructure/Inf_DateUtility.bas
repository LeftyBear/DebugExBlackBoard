Attribute VB_Name = "Inf_DateUtility"
'@Folder("Infrastructure.Service")
Option Explicit
Option Private Module

Public Function DiffApr1st(ByVal TargetDate As Date) As Long
    DiffApr1st = TargetDate - GetApr1st(TargetDate) + 1
End Function

Public Function GetSchoolYear(ByVal TargetDate As Date) As Long
    If VBA.Month(TargetDate) < 4 Then
        GetSchoolYear = VBA.Year(TargetDate) - 1
    Else
        GetSchoolYear = VBA.Year(TargetDate)
    End If
End Function

Public Function NormalizeToDate(ByVal RawText As String) As Date
    Dim TextValue As String
    TextValue = VBA.Trim$(RawText)
    If TextValue = vbNullString Then Exit Function
    TextValue = VBA.StrConv(TextValue, vbNarrow)
    '@Ignore AssignmentNotUsed
    TextValue = VBA.Replace(TextValue, charHyphen, charBackSlash)
    If 0 < VBA.InStr(1, TextValue, charSlash) Then
        If UBound(VBA.Split(TextValue, charSlash)) <> 2 Then Exit Function
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
        Exit Function
    End If
    On Error Resume Next
    NormalizeToDate = VBA.DateSerial(YearPart, MonthPart, DayPart)
    On Error GoTo 0
End Function
