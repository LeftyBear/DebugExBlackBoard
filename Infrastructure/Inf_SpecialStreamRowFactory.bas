Attribute VB_Name = "Inf_SpecialStreamRowFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Private Const SPECIAL_NAME As String = "“ÁŽx–¼"

Public Function Create(ByVal Header As String, ByVal RawText As String) As Inf_SpecialStreamRow
    Dim Result As Inf_SpecialStreamRow
    Set Result = New Inf_SpecialStreamRow
    If Header = SPECIAL_NAME Then
        Result.Name = RawText
    End If
    Set Create = Result
End Function
