Attribute VB_Name = "Inf_ScheduleRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal Column As Inf_ScheduleColumn, ByVal RawText As String) As Inf_ScheduleRow
    Dim Result As Inf_ScheduleRow
    Set Result = New Inf_ScheduleRow
    If Column.RawDate <> vbNullString Then
        Result.NormDate = Inf_DateUtility.NormalizeToDate(RawText)
    ElseIf Column.RawID <> vbNullString Then
        Result.RawID = Column.RawID
        Result.Value = RawText
    End If
    Set Create = Result
End Function
