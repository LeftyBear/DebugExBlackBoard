Attribute VB_Name = "Inf_ScheduleRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawDate As String, ByVal Column As Inf_ScheduleColumn, ByVal RawText As String) As Inf_ScheduleRow
    Dim Result As Inf_ScheduleRow
    Set Result = New Inf_ScheduleRow
    Result.NormDate = Inf_DateUtility.NormalizeToDate(RawDate)
    If Column.Name <> vbNullString Then
        Result.Name = Column.Name
        Result.Value = RawText
    End If
    Set Create = Result
End Function
