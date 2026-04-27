VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_TimeTableService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Service")
Option Explicit

Private Type Member
    Counter As App_TimeTableCounter
End Type

Private This As Member

Public Sub Initialize(ByVal Counter As App_TimeTableCounter)
    Set This.Counter = Counter
End Sub

Public Function CountHours(ByVal Aggregate As Dom_ClassHourAggregate, ByVal Condition As App_TimeTableFilterCondition) As Long
    CountHours = This.Counter.Count(Aggregate, Condition)
End Function
