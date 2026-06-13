Attribute VB_Name = "Inf_ScheduleRowsFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_ScheduleRows
    Dim Result As Inf_ScheduleRows
    Set Result = New Inf_ScheduleRows
    Dim Map As Inf_ScheduleHeaderMap
    Set Map = Inf_ScheduleHeaderMapFactory.Create(RawRows.GetHeaders)
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 2 To RawRows.ColumnsCount(R)
            Dim RawDate As String
            RawDate = RawRows.GetRow(R, 1)
            Dim Column As Inf_ScheduleColumn
            Set Column = Map.Item(CStr(C))
            Dim Row As Inf_ScheduleRow
            Set Row = Inf_ScheduleRowFactory.Create(RawDate, Column, RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function

