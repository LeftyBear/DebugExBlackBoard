VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Public Function FilterByGender(ByVal Aggregate As Dom_DailyEnrollment, ByVal Condition As App_EnrollmentIDCondition) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To Aggregate.Count
        Dim Enrollment As Dom_Enrollment
        Set Enrollment = Aggregate.Item(i)
        If Condition.GenderMatchs(Enrollment) Then
            Result.Add Enrollment
        End If
    Next
    Set FilterByGender = Result
End Function

Public Function FilterByClass(ByVal Aggregate As Dom_DailyEnrollment, ByVal Condition As App_EnrollmentIDCondition) As VBA.Collection
    Dim Result As VBA.Collection
    Set Result = New VBA.Collection
    Dim i As Long
    For i = 1 To Aggregate.Count
        Dim Enrollment As Dom_Enrollment
        Set Enrollment = Aggregate.Item(i)
        If Condition.ClassMatchs(Enrollment) Then
            Result.Add Enrollment
        End If
    Next
    Set FilterByClass = Result
End Function
