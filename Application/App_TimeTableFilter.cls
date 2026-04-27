VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function FilterByGrade(ByVal Aggregate As Dom_ClassHourAggregate, ByVal Condition As App_TimeTableFilterCondition) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To Aggregate.TimeTablesCount
        Dim TimeTable As Dom_TimeTable
        Set TimeTable = Aggregate.TimeTablesItem(i)
        If Condition.GradeMatchs(TimeTable) Then
            Result.Add TimeTable
        End If
    Next
    Set FilterByGrade = Result
End Function

Public Function FilterByPeriod(ByVal Aggregate As Dom_ClassHourAggregate, ByVal Condition As App_TimeTableFilterCondition) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To Aggregate.TimeTablesCount
        Dim TimeTable As Dom_TimeTable
        Set TimeTable = Aggregate.TimeTablesItem(i)
        If Condition.PeriodMatchs(TimeTable) Then
            Result.Add TimeTable
        End If
    Next
    Set FilterByPeriod = Result
End Function
