VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function GetDailyMarkCount(ByVal Yearly As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Calculator As Dom_TimeTableCalculator, ByVal Filter As Dom_TimeTableFilter) As Long
    Dim Daily As Dom_ClassHourAggregate
    Set Daily = Yearly.GetByDate(TargetDate)
    Select Case Filter.ClassHourType.Value
    Case Dom_ClassHourTypeEnum.Plan
        Dim Items() As Variant
        Items = Daily.GetPlanMarks
    Case Dom_ClassHourTypeEnum.Execution
        Items = Daily.GetExeMarks
    End Select
    GetDailyMarkCount = Calculator.Calculate(Items, Filter)
End Function
