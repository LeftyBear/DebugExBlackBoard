VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function GetDailyHourTotal(ByVal Yearly As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Calculator As Dom_ClassHourCalculator, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim Daily As Dom_ClassHourAggregate
    Set Daily = Yearly.GetByDate(TargetDate)
    Select Case Filter.ClassHourType.Value
    Case Dom_ClassHourTypeEnum.Plan
        Dim Items() As Variant
        Items = Daily.GetPlanValues
    Case Dom_ClassHourTypeEnum.Execution
        Items = Daily.GetExeValues
    End Select
    GetDailyHourTotal = Calculator.Calculate(Items, Filter)
End Function

Public Function GetWeeklyHourTotal(ByVal Yearly As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim TargetWeeks As Long
    TargetWeeks = Dom_Util_Date.GetSchoolWeek(TargetDate.Value)
    Dim DateKeys() As Variant
    DateKeys = GetDateKeys(Yearly, TargetDate, Filter)
    Dim DateKey As Variant
    For Each DateKey In DateKeys
        If Dom_Util_Date.GetSchoolWeek(CDate(DateKey)) = TargetWeeks Then
            Dim Total As Long
            Total = Total + Me.GetDailyHourTotal(Yearly, TargetDate, New Dom_ClassHourCalculator, Filter)
        End If
    Next
    GetWeeklyHourTotal = Total
End Function

Public Function GetMonthlyHourTotal(ByVal Yearly As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Calculator As Dom_ClassHourCalculator, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim TargetMonth As Long
    TargetMonth = VBA.Month(TargetDate)
    Dim DateKeys() As Variant
    DateKeys = GetDateKeys(Yearly, TargetDate, Filter)
    Dim DateKey As Variant
    For Each DateKey In DateKeys
        If VBA.Month(CDate(DateKey)) = TargetMonth Then
            Dim Total As Long
            Total = Total + Me.GetDailyHourTotal(Yearly, TargetDate, Calculator, Filter)
        End If
    Next
    GetMonthlyHourTotal = Total
End Function

Public Function GetCumulativeHourTotal(ByVal Yearly As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Calculator As Dom_ClassHourCalculator, ByVal Filter As Dom_ClassHourFilter) As Long
    Dim DateKeys() As Variant
    DateKeys = GetDateKeys(Yearly, TargetDate, Filter)
    Dim DateKey As Variant
    For Each DateKey In DateKeys
        If CDate(DateKey) <= TargetDate.Value Then
            Dim Total As Long
            Total = Total + Me.GetDailyHourTotal(Yearly, TargetDate, Calculator, Filter)
        End If
    Next
    GetCumulativeHourTotal = Total
End Function

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

Private Function GetDateKeys(ByVal Yearly As Dom_ClassHourYearAggregate, ByVal TargetDate As Dom_DateID, ByVal Filter As Dom_ClassHourFilter) As Variant
    Dim Daily As Dom_ClassHourAggregate
    Set Daily = Yearly.GetByDate(TargetDate)
    Select Case Filter.ClassHourType.Value
    Case Dom_ClassHourTypeEnum.Plan
        GetDateKeys = Daily.GetPlanValueKeys
    Case Dom_ClassHourTypeEnum.Execution
        GetDateKeys = Daily.GetExeValueKeys
    End Select
End Function
