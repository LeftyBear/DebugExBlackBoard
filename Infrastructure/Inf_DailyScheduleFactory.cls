VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.Factory")
Option Explicit

Private Type Member
    Factory As Inf_ScheduleFactory
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.Factory = New Inf_ScheduleFactory
End Sub

Public Function Create(ByVal DailyRows As Inf_ScheduleRows) As Dom_DailySchedules
    Dim Result As Dom_DailySchedules
    Set Result = New Dom_DailySchedules
    Dim i As Long
    For i = 1 To DailyRows.Count
        Dim Entry As Dom_Schedule
        Set Entry = This.Factory.Create(DailyRows.Item(i))
        If Not Entry Is Nothing Then Result.Add Entry
    Next
    Set Create = Result
End Function
