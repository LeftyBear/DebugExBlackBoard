Attribute VB_Name = "Inf_SubjectRowsFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_SubjectRows
    Dim Result As Inf_SubjectRows
    Set Result = New Inf_SubjectRows
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 1 To RawRows.ColumnsCount(R)
            Dim Row As Inf_SubjectRow
            Set Row = Inf_SubjectRowFactory.Create(RawRows.GetColumnName(C), RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
