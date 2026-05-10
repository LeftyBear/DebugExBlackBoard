VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourSummaryCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Calculator As App_ClassHourCalculator
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Calculator = New App_ClassHourCalculator
End Sub
'
'Public Sub Initialize(ByVal Calculator As App_ClassHourCalculator)
'    Set This.Calculator = Calculator
'End Sub

Public Function SumOf(ByVal YearlyAggregate As VBA.Collection, ByVal BeginIndex As Long, ByVal EndIndex As Long, ByVal Condition As App_ClassHourFilterCondition) As Dom_ClassHourValue
    Dim Total As Dom_ClassHourValue
    Set Total = New Dom_ClassHourValue
    Total.Initialize 0
    Dim Index As Long
    For Index = BeginIndex To EndIndex
        Dim Value As Dom_ClassHourValue
        Set Value = This.Calculator.SumOfOneDay(YearlyAggregate.Item(Index), Condition)
        Set Total = Total.Add(Value)
    Next
    Set SumOf = Total
End Function
