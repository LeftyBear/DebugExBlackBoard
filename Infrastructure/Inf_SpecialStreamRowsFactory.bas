Attribute VB_Name = "Inf_SpecialStreamRowsFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_SpecialStreamRows
    Dim Result As Inf_SpecialStreamRows
    Set Result = New Inf_SpecialStreamRows
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Dim Row As Inf_SpecialStreamRow
            Set Row = Inf_SpecialStreamRowFactory.Create(RawRows.GetHeader(C), RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
