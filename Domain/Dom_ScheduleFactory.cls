VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.Factory")
Option Explicit

Public Function Create(ByVal ID As Dom_ScheduleID, ByVal RawText As String) As Dom_Schedule
    Dim ScheduleValue As Dom_ScheduleValue
    Set ScheduleValue = CreateSchedlueValue(RawText)
    Dim Result As Dom_Schedule
    Set Result = New Dom_Schedule
    Result.Initialize ID, ScheduleValue
    Set Create = Result
End Function

Private Function CreateSchedlueValue(ByVal RawText As String) As Dom_ScheduleValue
    With New Dom_ScheduleValue
        Set CreateSchedlueValue = .Create(RawText)
    End With
End Function

