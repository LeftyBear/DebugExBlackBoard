Attribute VB_Name = "Inf_MainStreamRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module
Private Const COLUMN_UPPER_GRADE As String = "ïÅí äwîNêî"
Private Const COLUMN_UPPER_CLASS  As String = "ïÅí äwãâêî"

Public Function Create(ByVal ColumnName As String, ByVal RawText As String) As Inf_MainStreamRow
    Dim Result As Inf_MainStreamRow
    Set Result = New Inf_MainStreamRow
    If ColumnName = COLUMN_UPPER_GRADE Then
        Result.UpperGrade = RawText
    ElseIf ColumnName = COLUMN_UPPER_CLASS Then
        Result.UpperClassNo = RawText
    End If
    Set Create = Result
End Function
