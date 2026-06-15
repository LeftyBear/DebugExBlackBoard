Attribute VB_Name = "Inf_EnrollmentRowsFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal RawRows As Inf_RawRows) As Inf_EnrollmentRows
    Dim Result As Inf_EnrollmentRows
    Set Result = New Inf_EnrollmentRows
    Dim Map As Inf_EnrollmentHeaderMap
    Set Map = Inf_EnrollmentHeaderMapFactory.Create(RawRows.GetHeader)
    Dim R As Long
    For R = 2 To RawRows.RowsCount
        Dim C As Long
        For C = 2 To RawRows.ColumnsCount(R)
            Dim RawDate As String
            RawDate = RawRows.GetRow(R, 1)
            Dim Column As Inf_EnrollmentColumn
            Set Column = Map.Item(CStr(C))
            Dim Row As Inf_EnrollmentRow
            Set Row = Inf_EnrollmentRowFactory.Create(RawDate, Column, RawRows.GetRow(R, C))
            Result.Add Row
        Next
    Next
    Set Create = Result
End Function
