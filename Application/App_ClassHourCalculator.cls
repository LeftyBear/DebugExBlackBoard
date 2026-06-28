VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Calculator")
Option Explicit

Public Function CalculatePlanTotal(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels) As App_ClassHourTotals
    Dim Result As App_ClassHourTotals
    Set Result = New App_ClassHourTotals
    Dim i As Long
    For i = 1 To ClassHourReadModels.ClassHourPlanCount
        Dim Model As App_ClassHourReadModel
        Set Model = ClassHourReadModels.GetClassHourPlan(i)
        If IsTarget(Model, SubjectReadModels) Then
            Result.AddInternal Model.Grade, Model.Value
        End If
    Next
    Set CalculatePlanTotal = Result
End Function

Public Function CalculateExeTotal(ByVal ClassHourReadModels As App_ClassHourReadModels, ByVal SubjectReadModels As App_SubjectReadModels) As App_ClassHourTotals
    Dim Result As App_ClassHourTotals
    Set Result = New App_ClassHourTotals
    Dim i As Long
    For i = 1 To ClassHourReadModels.ClassHourExeCount
        Dim Model As App_ClassHourReadModel
        Set Model = ClassHourReadModels.GetClassHourExe(i)
        If IsTarget(Model, SubjectReadModels) Then
            Result.AddInternal Model.Grade, Model.Value
        End If
    Next
    Set CalculateExeTotal = Result
End Function

Private Function IsTarget(ByVal ClassHourReadModel As App_ClassHourReadModel, ByVal SubjectReadModels As App_SubjectReadModels) As Boolean
    Dim Subject As App_SubjectReadModel
    For Each Subject In SubjectReadModels
        If Subject.Name = ClassHourReadModel.Subject Then
            If 0 < VBA.InStr(1, Subject.TargetGrade, ClassHourReadModel.Grade) Then
                IsTarget = True: Exit For
            End If
        End If
    Next
End Function
