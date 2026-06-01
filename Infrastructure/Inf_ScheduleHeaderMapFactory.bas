Attribute VB_Name = "Inf_ScheduleHeaderMapFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal HeaderRows As VBA.Collection) As Inf_ScheduleHeaderMap
    Dim Result As Inf_ScheduleHeaderMap
    Set Result = New Inf_ScheduleHeaderMap
    Dim C As Long
    For C = 1 To HeaderRows.Count
        Dim Column As Inf_ScheduleColumn
        Set Column = CreateColumn(HeaderRows.Item(C))
        Result.Add C, Column
    Next
    Set Create = Result
End Function

Private Function CreateColumn(ByVal ColumnName As String) As Inf_ScheduleColumn
    Dim Column As Inf_ScheduleColumn
    Set Column = New Inf_ScheduleColumn
    If 0 < VBA.InStr(1, ColumnName, HIZUKE) Then
        Column.RawDate = ColumnName
    Else
        Column.RawID = ColumnName
    End If
    Set CreateColumn = Column
End Function
