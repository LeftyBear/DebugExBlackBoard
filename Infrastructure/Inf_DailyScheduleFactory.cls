VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_DailyScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit

Private Type Member
    IDMapper        As Inf_ScheduleIDMapper
    TransferMapper  As Inf_TransferMapper
End Type

Private This As Member

Private Sub Class_Initialize()
    Set This.IDMapper = New Inf_ScheduleIDMapper
    Set This.TransferMapper = New Inf_TransferMapper
End Sub

Public Function Create(ByVal TargetDate As Date, ByVal Row As Inf_ScheduleRow) As Dom_Schedule
    Dim DateID As Dom_DateID
    Set DateID = Dom_ValueObjectFactory.CreateDateID(TargetDate)
    Dim ID As Dom_ScheduleID
    Set ID = This.IDMapper.ToDomain(Row.RawID)
    Dim Value As Dom_ScheduleName
    Set Value = Dom_ValueObjectFactory.CreateScheduleValue(Row.Value)
    Dim Entity As Dom_Schedule
    Set Entity = New Dom_Schedule
    Entity.Initialize DateID, ID, Value
    Set Create = Entity
End Function
