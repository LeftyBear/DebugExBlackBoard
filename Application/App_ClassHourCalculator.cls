VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ClassHourCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Filter As App_ClassHourFilter
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Filter = New App_ClassHourFilter
End Sub
'
'Public Sub Initialize(ByVal Filter As App_ClassHourFilter)
'    Set This.Filter = Filter
'End Sub

Public Function SumOfOneDay(ByVal Aggregate As Dom_DailyClassHour, ByVal Condition As App_ClassHourFilterCondition) As Dom_ClassHourValue
    Dim Total As Dom_ClassHourValue
    Set Total = New Dom_ClassHourValue
    Total.Initialize 0
    Dim ClassHours As VBA.Collection
    Set ClassHours = This.Filter.FilterByGrade(Aggregate, Condition)
    Dim i As Long
    For i = 1 To ClassHours.Count
        Dim ClassHour As Dom_ClassHour
        Set ClassHour = ClassHours.Item(i)
        Set Total = Total.Add(ClassHour.Value)
    Next
    Set SumOfOneDay = Total
End Function
