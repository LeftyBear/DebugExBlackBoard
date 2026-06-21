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

Public Function CalculateTotal(ByVal Models As App_ClassHourReadModels, ByVal Subjects As App_SubjectReadModels, ByVal TypeKey As String) As App_ClassHourTotals
    Dim Result As App_ClassHourTotals
    Set Result = New App_ClassHourTotals
    Dim FilteredModels As App_ClassHourReadModels
    Set FilteredModels = FilterByType(Models, TypeKey)
    Dim Model As App_ClassHourReadModel
    For Each Model In FilteredModels
        If IsTarget(Model, Subjects) Then
            Result.AddInternal Model.Grade, Model.Value
        End If
    Next
    Set CalculateTotal = Result
End Function

Private Function FilterByType(ByVal Models As App_ClassHourReadModels, ByVal TypeKey As String) As App_ClassHourReadModels
    Dim Result As App_ClassHourReadModels
    Set Result = New App_ClassHourReadModels
    Dim i As Long
    For i = 1 To Models.ClassHourCount
        Dim Model As App_ClassHourReadModel
        Set Model = Models.ClassHourItem(i)
        If Model.TypeKey = TypeKey Then
            Result.AddClassHour Model
        End If
    Next
    Set FilterByType = Result
End Function

Private Function IsTarget(ByVal Model As App_ClassHourReadModel, ByVal Subjects As App_SubjectReadModels) As Boolean
    Dim i As Long
    For i = 1 To Subjects.Count
        Dim Subject As App_SubjectReadModel
        Set Subject = Subjects.Item(i)
        If Subject.Name = Model.Subject Then
            If 0 < VBA.InStr(1, Subject.TargetGrade, Model.Grade) Then
                IsTarget = True: Exit For
            End If
        End If
    Next
End Function
