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

Public Function CountPeriodPlan(ByVal Models As App_ClassHourReadModels) As App_PeriodTotals
    Dim Result As App_PeriodTotals
    Set Result = New App_PeriodTotals
    Dim i As Long
    For i = 1 To Models.TimeTablePlanCount
        Dim Model As App_TimeTableReadModel
        Set Model = Models.GetTimeTablePlan(i)
        If Model.Mark <> vbNullString Then
            Result.AddInternal Model.Grade
        End If
    Next
    Set CountPeriodPlan = Result
End Function

Public Function CountPeriodExe(ByVal Models As App_ClassHourReadModels) As App_PeriodTotals
    Dim Result As App_PeriodTotals
    Set Result = New App_PeriodTotals
    Dim i As Long
    For i = 1 To Models.TimeTableExeCount
        Dim Model As App_TimeTableReadModel
        Set Model = Models.GetTimeTableExe(i)
        If Model.Mark <> vbNullString Then
            Result.AddInternal Model.Grade
        End If
    Next
    Set CountPeriodExe = Result
End Function
