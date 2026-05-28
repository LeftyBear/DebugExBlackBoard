VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_SubjectModelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Public Function Create(ByVal Rows As Inf_SubjectRows) As App_SubjectReadModels
    Dim Result As App_SubjectReadModels
    Set Result = New App_SubjectReadModels
    Dim i As Long
    For i = 1 To Rows.Count
        Dim Row As Inf_SubjectRow
        Set Row = Rows.Item(i)
        Dim Model As App_SubjectReadModel
        Set Model = New App_SubjectReadModel
        If Row.Name <> vbNullString Then
            Model.Name = Row.Name
        ElseIf Row.TargetGrade <> vbNullString Then
            Model.TargetGrade = Row.TargetGrade
        ElseIf Row.Mark <> vbNullString Then
            Model.Mark = Row.Mark
        End If
        If i Mod 3 = 0 Then Result.Add Model
    Next
    Set Create = Result
End Function
