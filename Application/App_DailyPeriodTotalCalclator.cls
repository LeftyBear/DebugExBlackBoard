VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_DailyPeriodTotalCalclator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.UseCase")
Option Explicit

Public Function Calclate(ByVal ClassHours As App_ClassHourReadModels, ByVal TypeCode As String, ByVal MainStream As App_MainStreamReadModel) As App_DailyPeriodReadModels
    Dim Result As App_DailyPeriodReadModels
    Set Result = New App_DailyPeriodReadModels
    Dim Calculator As App_PeriodCalculator
    Set Calculator = New App_PeriodCalculator
    Dim Grade As Long
    For Grade = 1 To MainStream.UpperGrade
        Dim Total As Long
        Total = Calculator.CalculateTotal(ClassHours, TypeCode, Grade)
        Dim Model As App_DailyPeriodReadModel
        Set Model = New App_DailyPeriodReadModel
        Model.Grade = Grade
        Model.Total = Total
        Result.Add Model
    Next
    Calclate = Result
End Function
