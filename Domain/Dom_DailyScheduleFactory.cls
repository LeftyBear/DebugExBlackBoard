VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_DailyScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit

Private Type Member
    Factory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal Factory As Dom_ValueObjectFactory)
    Set This.Factory = Factory
End Sub

Public Function Create(ByVal Row As App_ScheduleRow) As Dom_Schedule
    Dim ID As Dom_ScheduleID
    Set ID = This.Factory.CreateScheduleID(Row.EventName)
    If ID Is Nothing Then Exit Function
    Dim DateID As Dom_DateID
    Set DateID = This.Factory.CreateDateID(Row.NormDate)
    Dim Value As Dom_ScheduleValue
    Set Value = This.Factory.CreateScheduleValue(Row.Value)
    Dim Result As Dom_Schedule
    Set Result = New Dom_Schedule
    Result.Initialize DateID, ID, Value
    Set Create = Result
End Function
