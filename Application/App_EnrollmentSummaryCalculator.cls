VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentSummaryCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Calculator As App_EnrollmentCalculator
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Calculator = New App_EnrollmentCalculator
End Sub
'
'Public Sub Initialize(ByVal Calculator As App_EnrollmentCalculator)
'    Set This.Calculator = Calculator
'End Sub

Public Function SumClassOf(ByVal Aggregate As Dom_DailyEnrollment, ByVal Condition As App_EnrollmentIDCondition) As Dom_EnrollmentValue
    Set SumClassOf = This.Calculator.SumClassOfOneDay(Aggregate, Condition)
End Function

Public Function SumGenderOf(ByVal Aggregate As Dom_DailyEnrollment, ByVal Condition As App_EnrollmentIDCondition) As Dom_EnrollmentValue
    Set SumGenderOf = This.Calculator.SumGenderOfOneDay(Aggregate, Condition)
End Function
