VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_ScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Private Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByVal ID As Dom_ScheduleID, ByVal RawText As String) As Dom_Schedule
    'TODO Date
    Dim DateID As Dom_EnrollmentID
    Set DateID = This.Factory.CreateDateID(Date)
    Dim ScheduleValue As Dom_ScheduleValue
    Set ScheduleValue = This.Factory.CreateSchedlueValue(RawText)
    Dim Result As Dom_Schedule
    Set Result = New Dom_Schedule
    Result.Initialize DateID, ID, ScheduleValue
    Set Create = Result
End Function
