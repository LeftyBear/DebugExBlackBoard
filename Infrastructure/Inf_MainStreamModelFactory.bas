Attribute VB_Name = "Inf_MainStreamModelFactory"
'@Folder("Infrastructure.Factory")
Option Explicit
Option Private Module

Public Function Create(ByVal Rows As Inf_MainStreamRows) As App_MainStreamReadModel
    Dim Result As App_MainStreamReadModel
    Set Result = New App_MainStreamReadModel
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_MainStreamRow
        Set Row = Rows.Item(i)
        If Row.UpperGrade <> vbNullString Then
            Result.UpperGrade = Row.UpperGrade
        ElseIf Row.UpperClassNo <> vbNullString Then
            Result.UpperClassNo = Row.UpperClassNo
        End If
    Next
    Set Create = Result
End Function
