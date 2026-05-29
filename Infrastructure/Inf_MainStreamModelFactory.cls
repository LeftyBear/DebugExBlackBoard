VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_MainStreamModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

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
