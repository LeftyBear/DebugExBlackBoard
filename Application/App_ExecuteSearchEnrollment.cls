VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ExecuteSearchEnrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Execute")
Option Explicit

Public Function Execute(ByVal Aggregate As Dom_EnrollmentAggregate, ByVal Keywords As VBA.Collection, ByVal Matcher As App_RemarksKeywordMatcher) As Dom_EnrollmentAggregate
    Dim Filtered As VBA.Collection
    Set Filtered = New VBA.Collection
    Dim Enrollments As VBA.Collection
    Set Enrollments = Aggregate.GetEntitiesInternal
    Dim Enrollment As Dom_Enrollment
    For Each Enrollment In Enrollments
        If Matcher.MatchAll(Enrollment.Remarks, Keywords) Then Filtered.Add Enrollment
    Next
    Dim Result As Dom_EnrollmentAggregate
    Set Result = New Dom_EnrollmentAggregate
    Result.Initialize Filtered
    Set Execute = Result
End Function
