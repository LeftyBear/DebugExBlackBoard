VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Calculator")
Option Explicit

Private Type Member
    Filter As App_EnrollmentFilter
End Type

Private This As Member

Public Sub Initialize(ByVal Filter As App_EnrollmentFilter)
    Set This.Filter = Filter
End Sub

Public Function SumClass(ByVal Aggregate As Dom_EnrollmentAggregate, ByVal Condition As App_EnrollmentFilterCondition) As Dom_EnrollmentValue
    Dim Total As Dom_EnrollmentValue
    Set Total = New Dom_EnrollmentValue
    Total.Initialize 0
    Dim Entities As VBA.Collection
    Set Entities = This.Filter.FilterByClass(Aggregate, Condition)
    Dim i As Long
    For i = 1 To Entities.Count
        Dim Enrollment As Dom_Enrollment
        Set Enrollment = Entities.Item(i)
        Set Total = Total.Add(Enrollment)
    Next
    Set SumClass = Total
End Function

Public Function SumGender(ByVal Aggregate As Dom_EnrollmentAggregate, ByVal Condition As App_EnrollmentFilterCondition) As Dom_EnrollmentValue
    Dim Total As Dom_EnrollmentValue
    Set Total = New Dom_EnrollmentValue
    Total.Initialize 0
    Dim Entities As VBA.Collection
    Set Entities = This.Filter.FilterByGender(Aggregate, Condition)
    Dim i As Long
    For i = 1 To Entities.Count
        Dim Enrollment As Dom_Enrollment
        Set Enrollment = Entities.Item(i)
        Set Total = Total.Add(Enrollment)
    Next
    Set SumGender = Total
End Function
