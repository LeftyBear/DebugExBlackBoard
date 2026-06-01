Attribute VB_Name = "Inf_PeriodRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_PeriodRow
    Dim Result As Inf_PeriodRow
    Set Result = New Inf_PeriodRow
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Result.Value = RawRows.GetRow(R, C)
        Next
    Next
    Set Create = Result
End Function
