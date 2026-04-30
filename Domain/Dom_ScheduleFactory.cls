VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dom_ScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Domain.EntityFactory")
Option Explicit

Private Type Member
    ValueObjectFactory As Dom_ValueObjectFactory
End Type

Private This As Member

Public Sub Initialize(ByVal ValueObjectFactory As Dom_ValueObjectFactory)
    Set This.ValueObjectFactory = ValueObjectFactory
End Sub

Public Function Create(ByVal Row As App_ScheduleRow) As Dom_Schedule
    Dim DateID As Dom_ScheduleID
    Set DateID = This.ValueObjectFactory.CreateDateID(Row.NormDate)
    Dim ID As Dom_ScheduleID
    Set ID = This.ValueObjectFactory.CreateScheduleID(Row.EventName)
    Dim Value As Dom_ScheduleValue
    Set Value = This.ValueObjectFactory.CreateScheduleValue(Row.Value)
    If ID Is Nothing Then Exit Function
    Dim Entity As Dom_Schedule
    Set Entity = New Dom_Schedule
    Entity.Initialize DateID, ID, Value
    Set Create = Entity
End Function
