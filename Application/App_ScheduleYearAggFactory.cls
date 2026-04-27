VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ScheduleYearAggFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal Rows As VBA.Collection) As Dom_ScheduleYearAggregate
    Dim Yearly As Dom_ScheduleYearAggregate
    Set Yearly = New Dom_ScheduleYearAggregate
    Dim Daily As Dom_ScheduleAggregate
    For Each Daily In Rows
        Yearly.Add Daily
    Next
    Set Create = Yearly
End Function
