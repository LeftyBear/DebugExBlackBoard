Attribute VB_Name = "Inf_MainStreamRowsFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_MainStreamRows
    Dim Result As Inf_MainStreamRows
    Set Result = New Inf_MainStreamRows
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Dim Row As Inf_MainStreamRow
            Set Row = Inf_MainStreamRowFactory.Create(RawRows.GetColumnName(C), RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
