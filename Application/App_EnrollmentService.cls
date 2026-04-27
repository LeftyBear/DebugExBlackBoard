VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EnrollmentService"
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

Public Sub Initialize(ByVal Calculator As App_EnrollmentCalculator)
    Set This.Calculator = Calculator
End Sub

Public Function GetMainClassTotal(ByVal Aggregate As Dom_EnrollmentAggregate, ByVal Condition As App_EnrollmentFilterCondition) As Dom_EnrollmentValue
    Set GetMainClassTotal = This.Calculator.SumClass(Aggregate, Condition)
End Function

Public Function GetMainGenderTotal(ByVal Aggregate As Dom_EnrollmentAggregate, ByVal Condition As App_EnrollmentFilterCondition) As Dom_EnrollmentValue
    Set GetMainGenderTotal = This.Calculator.SumGender(Aggregate, Condition)
End Function
