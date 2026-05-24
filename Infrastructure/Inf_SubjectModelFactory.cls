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

Public Function Create(ByVal Row As Inf_SubjectRow) As App_SubjectReadModel
    Dim Result As App_SubjectReadModel
    Set Result = New App_SubjectReadModel
    Result.Name = Row.Name
    Result.TargetGrade = Row.TargetGrade
    Result.Mark = Row.Mark
    Set Create = Result
End Function
