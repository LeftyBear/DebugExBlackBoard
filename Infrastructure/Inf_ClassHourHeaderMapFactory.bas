Attribute VB_Name = "Inf_ClassHourHeaderMapFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module
Private Const COLUMN_DATE As String = "“ú•t"

Public Function Create(ByRef Headers As Variant) As Inf_ClassHourHeaderMap
    Dim Result As Inf_ClassHourHeaderMap
    Set Result = New Inf_ClassHourHeaderMap
    Dim C As Long
    For C = LBound(Headers) To UBound(Headers)
        Dim Column As Inf_ClassHourColumn
        Set Column = CreateColumn(Headers(C))
        Result.Add CStr(C), Column
    Next
    Set Create = Result
End Function

Private Function CreateColumn(ByVal ColumnName As String) As Inf_ClassHourColumn
    Dim Result As Inf_ClassHourColumn
    Set Result = New Inf_ClassHourColumn
    If 0 < VBA.InStr(1, ColumnName, COLUMN_DATE) Then
        Result.RawDate = ColumnName
    Else
        Result.Name = ColumnName
    End If
    Set CreateColumn = Result
End Function
