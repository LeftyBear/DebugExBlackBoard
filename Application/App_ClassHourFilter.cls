VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function FilterByGrade(ByVal Aggregate As Dom_DailyClassHour, ByVal Condition As App_ClassHourFilterCondition) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To Aggregate.ClassHoursCount
        Dim ClassHour As Dom_ClassHour
        Set ClassHour = Aggregate.ClassHoursItem(i)
        If Condition.GradeMatchs(ClassHour) Then
            Result.Add ClassHour
        End If
    Next
    Set FilterByGrade = Result
End Function

Public Function FilterBySubject(ByVal Aggregate As Dom_DailyClassHour, ByVal Condition As App_ClassHourFilterCondition) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To Aggregate.ClassHoursCount
        Dim ClassHour As Dom_ClassHour
        Set ClassHour = Aggregate.ClassHoursItem(i)
        If Condition.SubjectMatchs(ClassHour) Then
            Result.Add ClassHour
        End If
    Next
    Set FilterBySubject = Result
End Function
