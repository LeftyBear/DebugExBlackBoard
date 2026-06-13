VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Inf_ScheduleFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Infrastructure.EntityFactory")
Option Explicit

Public Function Create(ByVal Row As Inf_ScheduleRow) As Dom_Schedule
    Dim DateID As Dom_DateID
    Set DateID = Dom_ValueObjectFactory.CreateDateID(Row.NormDate)
    Dim Name As Dom_ScheduleName
    Set Name = Dom_ValueObjectFactory.CreateScheduleName(Row.Name)
    Dim Value As Dom_ScheduleValue
    Set Value = Dom_ValueObjectFactory.CreateScheduleValue(Row.Value)
    Dim Entry As Dom_Schedule
    Set Entry = New Dom_Schedule
    Entry.Initialize DateID, Name, Value
    Set Create = Entry
End Function

