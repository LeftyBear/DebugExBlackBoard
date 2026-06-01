Attribute VB_Name = "Inf_SchoolEventRowsFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_SchoolEventRows
    Dim Result As Inf_SchoolEventRows
    Set Result = New Inf_SchoolEventRows
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Dim Row As Inf_SchoolEventRow
            Set Row = Inf_SchoolEventRowFactory.Create(RawRows.GetHeader(C), RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
