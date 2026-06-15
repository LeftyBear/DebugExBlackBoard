Attribute VB_Name = "Inf_ScheduleHeaderMapFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module
Private Const COLUMN_DATE As String = "“ú•t"

Public Function Create(ByRef Header As Variant) As Inf_ScheduleHeaderMap
    Dim Result As Inf_ScheduleHeaderMap
    Set Result = New Inf_ScheduleHeaderMap
    Dim C As Long
    For C = LBound(Header) To UBound(Header)
        Dim Column As Inf_ScheduleColumn
        Set Column = CreateColumn(Header(C))
        Result.Add CStr(C), Column
    Next
    Set Create = Result
End Function

Private Function CreateColumn(ByVal ColumnName As String) As Inf_ScheduleColumn
    Dim Column As Inf_ScheduleColumn
    Set Column = New Inf_ScheduleColumn
    If 0 < VBA.InStr(1, ColumnName, COLUMN_DATE) Then
        Column.RawDate = ColumnName
    Else
        Column.Name = ColumnName
    End If
    Set CreateColumn = Column
End Function
