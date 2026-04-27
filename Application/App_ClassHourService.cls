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

Private Type Member
    Calculator As App_ClassHourCalculator
End Type

Private This As Member

Public Sub Initialize(ByVal Calculator As App_ClassHourCalculator)
    Set This.Calculator = Calculator
End Sub

Public Function CalculateTotal(ByVal Yearly As Dom_ClassHourYearAggregate, ByVal BeginIndex As Long, ByVal EndIndex As Long, ByVal Condition As App_ClassHourFilterCondition) As Dom_ClassHourValue
'    Dim BeginDate As Date
'    BeginDate = App_DateUtility.GetMonday(InitialDate.Value)
'    Dim BeginIndex As Long
'    BeginIndex = App_DateUtility.DiffApr1st(BeginDate)
'    Dim EndDate As Date
'    EndDate = App_DateUtility.GetSunday(InitialDate.Value)
'    Dim EndIndex As Long
'    EndIndex = App_DateUtility.DiffApr1st(EndDate)
    Dim Total As Dom_ClassHourValue
    Set Total = New Dom_ClassHourValue
    Total.Initialize 0
    Dim Index As Long
    For Index = BeginIndex To EndIndex
        Dim Value As Dom_ClassHourValue
        Set Value = This.Calculator.SumHours(Yearly.Item(Index), Condition)
        Set Total = Total.Add(Value)
    Next
    Set CalculateTotal = Total
End Function
