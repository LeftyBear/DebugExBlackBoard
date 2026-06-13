Attribute VB_Name = "Inf_SchoolEventRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module
Private Const COLUMN_EVENT_NAME As String = "ƒCƒxƒ“ƒg–¼"

Public Function Create(ByVal ColumnName As String, ByVal RawText As String) As Inf_SchoolEventRow
    Dim Result As Inf_SchoolEventRow
    Set Result = New Inf_SchoolEventRow
    If ColumnName = COLUMN_EVENT_NAME Then
        Result.Name = RawText
    End If
    Set Create = Result
End Function
