Attribute VB_Name = "Inf_ClassHourRowsFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_ClassHourRows
    Dim Result As Inf_ClassHourRows
    Set Result = New Inf_ClassHourRows
    Dim Map As Inf_ClassHourHeaderMap
    Set Map = Inf_ClassHourHeaderMapFactory.Create(RawRows.GetHeaders)
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Dim Row As Inf_ClassHourRow
            Set Row = Inf_ClassHourRowFactory.Create(Map.Item(C), RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
