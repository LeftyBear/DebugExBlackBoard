VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_PeriodCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service.Calculator")
Option Explicit

Public Function CountPeriod(ByVal Models As App_ClassHourReadModels, ByVal TypeKey As String) As App_PeriodTotals
    Dim Result As App_PeriodTotals
    Set Result = New App_PeriodTotals
    Dim FilterdModels As App_ClassHourReadModels
    Set FilterdModels = FilterByType(Models, TypeKey)
    Dim Model As App_TimeTableReadModel
    For Each Model In FilterdModels
        If Model.Mark <> vbNullString Then
            If Result.Item(Model.Grade) = 0 Then
                Result.Add Model.Grade, 1
            Else
                Result.Add Model.Grade, Result.Item(Model.Grade) + 1
            End If
        End If
    Next
    Set CountPeriod = Result
End Function

Private Function FilterByType(ByVal Models As App_ClassHourReadModels, ByVal TypeKey As String) As App_ClassHourReadModels
    Dim Result As App_ClassHourReadModels
    Set Result = New App_ClassHourReadModels
    Dim i As Long
    For i = 1 To Models.TimeTableCount
        Dim Model As App_TimeTableReadModel
        Set Model = Models.TimeTableItem(i)
        If Model.TypeKey = TypeKey Then
            Result.AddTimeTable Model
        End If
    Next
    Set FilterByType = Result
End Function
