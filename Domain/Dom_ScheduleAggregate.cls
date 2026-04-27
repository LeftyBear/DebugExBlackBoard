VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule UseMeaningfulName
'@Folder("Domain.Entity")
Option Explicit

Private Type Member
    DateID      As Dom_DateID
    Schedules   As VBA.Collection
End Type

Private This As Member

Public Sub Initialize(ByVal DateID As Dom_DateID, ByVal Schedules As VBA.Collection)
    Set This.DateID = DateID
    Set This.Schedules = Schedules
 End Sub

Public Sub UpdateSchedule(ByVal ID As Dom_ScheduleID, ByVal NewValue As String)
    Dim Target As Dom_Schedule
    Set Target = FindSchedule(ID)
    If Target Is Nothing Then Err.Raise DomErrNothingObject, "Dom_ScheduleAggregate", "Scheduleオブジェクトが必要です。"
    Dim Value As Dom_ScheduleValue
    Set Value = New Dom_ScheduleValue
    Value.Initialize NewValue
    Target.ChangeValueInternal Value
End Sub

Public Property Get DateID() As Dom_DateID
    Set DateID = This.DateID
End Property

Public Function ScheduleCount() As Long
    ScheduleCount = This.Schedules.Count
End Function

Public Function ScheduleItem(ByVal Index As Long) As Dom_Schedule
    Set ScheduleItem = This.Schedules.Item(Index)
End Function

Private Function FindSchedule(ByVal ID As Dom_ScheduleID) As Dom_Schedule
    Dim i As Long
    For i = 1 To This.Schedules.Count
        Dim Schedule As Dom_Schedule
        Set Schedule = This.Schedules.Item(i)
        If Schedule.ID.Equals(ID) Then
            Set FindSchedule = Schedule: Exit Function
        End If
    Next
End Function
