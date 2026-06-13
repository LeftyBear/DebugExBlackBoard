Attribute VB_Name = "Inf_SpecialStreamRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module
Private Const COLUMN_SPECIAL_NAME As String = "“ÁŽx–¼"

Public Function Create(ByVal ColumnName As String, ByVal RawText As String) As Inf_SpecialStreamRow
    Dim Result As Inf_SpecialStreamRow
    Set Result = New Inf_SpecialStreamRow
    If ColumnName = COLUMN_SPECIAL_NAME Then
        Result.Name = RawText
    End If
    Set Create = Result
End Function
