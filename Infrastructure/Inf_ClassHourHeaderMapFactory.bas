Attribute VB_Name = "Inf_ClassHourHeaderMapFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal HeaderRows As VBA.Collection) As Inf_ClassHourHeaderMap
    Dim Result As Inf_ClassHourHeaderMap
    Set Result = New Inf_ClassHourHeaderMap
    Dim C As Long
    For C = 1 To HeaderRows.Count
        Dim Column As Inf_ClassHourColumn
        Set Column = CreateColumn(HeaderRows.Item(C))
        Result.Add C, Column
    Next
    Set Create = Result
End Function

Private Function CreateColumn(ByVal ColumnName As String) As Inf_ClassHourColumn
    Dim Result As Inf_ClassHourColumn
    Set Result = New Inf_ClassHourColumn
    If 0 < VBA.InStr(1, ColumnName, HIZUKE) Then
        Result.RawDate = ColumnName
    Else
        Result.RawID = ColumnName
    End If
    Set CreateColumn = Result
End Function
